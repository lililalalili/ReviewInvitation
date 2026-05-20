from pathlib import Path
import sys
import types
import zipfile

import pytest

import nb_review_invitation_agent.mailer_outlook as mailer_module
from nb_review_invitation_agent.mailer_outlook import DraftMessage, OutlookMailer


def _make_docx(path: Path, xml: str) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types></Types>")
        zf.writestr("word/document.xml", xml)


def test_win32com_not_imported_at_module_import_time():
    assert "win32com" not in mailer_module.__dict__


def test_render_template_docx_with_ooxml_replaces_and_escapes(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, "<w:t>Aaaaa Ttttt</w:t>")
    rendered = OutlookMailer()._render_template_docx_with_ooxml(source, {"Aaaaa": "Li", "Ttttt": "Title & Value"})
    try:
        with zipfile.ZipFile(rendered, "r") as zf:
            content = zf.read("word/document.xml").decode("utf-8")
        assert "Li" in content
        assert "Title &amp; Value" in content
    finally:
        rendered.unlink(missing_ok=True)


def test_remaining_placeholders_detected(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, "<w:t>Aaa</w:t><w:t>aa Dddd</w:t><w:t>din</w:t>")
    assert OutlookMailer()._remaining_placeholders_in_docx(source, {"Aaaaa": "Li"}) == ["Aaaaa", "Dddddin"]


def test_replace_placeholders_in_range_uses_replace_all():
    calls = []

    class FakeReplacement:
        def ClearFormatting(self):
            return None

    class FakeFind:
        Replacement = FakeReplacement()

        def ClearFormatting(self):
            return None

        def Execute(self, **kwargs):
            calls.append(kwargs)

    mailer = OutlookMailer()
    mailer._replace_placeholders_in_range(types.SimpleNamespace(Find=FakeFind()), {"Aaaaa": "Li"})
    assert calls and calls[0]["Replace"] == 2


def test_render_formatted_body_word_first_success_no_fallback(monkeypatch, tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, "<w:t>Aaaaa</w:t>")
    copied = {"copy": 0, "paste": 0, "opened": []}

    class FakeRange:
        def __init__(self, text):
            self.Text = text
            self.NextStoryRange = None
            self.Find = types.SimpleNamespace(
                ClearFormatting=lambda: None,
                Replacement=types.SimpleNamespace(ClearFormatting=lambda: None),
                Execute=lambda **_: None,
            )

        def Copy(self):
            return None

    class FakeDoc:
        def __init__(self):
            self.Content = FakeRange("clean")
            second = FakeRange("story2")
            first = FakeRange("story1")
            first.NextStoryRange = second
            self.StoryRanges = [first]
            self.Shapes = [types.SimpleNamespace(TextFrame=types.SimpleNamespace(HasText=True, TextRange=FakeRange("shape")))]

        def Close(self, _):
            return None

    class FakeDocuments:
        def Open(self, p):
            copied["opened"].append(p)
            return FakeDoc()

    class FakeWord:
        def __init__(self):
            self.Visible = False
            self.Documents = FakeDocuments()

        def Quit(self):
            return None

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return FakeWord()

    class FakeEditor:
        def Range(self, _a, _b):
            return types.SimpleNamespace(Paste=lambda: copied.__setitem__("paste", copied["paste"] + 1))

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_render_template_docx_with_ooxml", lambda *_: pytest.fail("fallback should not be called"))
    monkeypatch.setattr(
        mailer,
        "_replace_placeholders_in_range",
        lambda rng, _p: copied.__setitem__("copy", copied["copy"] + (1 if rng is not None else 0)),
    )

    draft = DraftMessage("to@example.com", "", "s", template_path=source, placeholders={"Aaaaa": "Li"})
    mailer._render_formatted_body_to_clipboard_via_word_com(draft)
    assert len(copied["opened"]) == 1
    assert copied["copy"] >= 4


def test_render_formatted_body_uses_fallback_when_word_still_has_placeholders(monkeypatch, tmp_path: Path):
    source = tmp_path / "template.docx"
    rendered = tmp_path / "rendered.docx"
    _make_docx(source, "<w:t>Aaaaa</w:t>")
    _make_docx(rendered, "<w:t>done</w:t>")
    opens = []

    class FakeDoc:
        def __init__(self):
            self.Content = types.SimpleNamespace(Copy=lambda: None)

        def Close(self, _):
            return None

    class FakeDocuments:
        def Open(self, p):
            opens.append(Path(p))
            return FakeDoc()

    class FakeWord:
        def __init__(self):
            self.Visible = False
            self.Documents = FakeDocuments()

        def Quit(self):
            return None

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return FakeWord()

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_render_template_docx_with_ooxml", lambda *_: rendered)
    states = iter([["Aaaaa"], []])
    monkeypatch.setattr(mailer, "_replace_placeholders_in_word_doc", lambda *_: None)
    monkeypatch.setattr(mailer, "_remaining_placeholders_in_word_doc", lambda _doc: next(states))
    monkeypatch.setattr(mailer, "_remaining_placeholders_in_docx", lambda *_: [])
    mailer._render_formatted_body_to_clipboard_via_word_com(DraftMessage("to@example.com", "", "s", source, {"Aaaaa": "Li"}))
    assert opens[0] == source
    assert opens[1] == rendered


def test_render_formatted_body_raises_when_still_remaining_after_fallback(monkeypatch, tmp_path: Path):
    source = tmp_path / "template.docx"
    rendered = tmp_path / "rendered.docx"
    _make_docx(source, "<w:t>Aaaaa</w:t>")
    _make_docx(rendered, "<w:t>Aaaaa</w:t>")

    class FakeWord:
        def __init__(self):
            self.Visible = False
            self.Documents = types.SimpleNamespace(Open=lambda *_: types.SimpleNamespace(Content=types.SimpleNamespace(Copy=lambda: None), Close=lambda *_: None))

        def Quit(self):
            return None

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return FakeWord()

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_render_template_docx_with_ooxml", lambda *_: rendered)
    monkeypatch.setattr(mailer, "_replace_placeholders_in_word_doc", lambda *_: None)
    monkeypatch.setattr(mailer, "_remaining_placeholders_in_word_doc", lambda _doc: ["Aaaaa"])
    monkeypatch.setattr(mailer, "_remaining_placeholders_in_docx", lambda *_: ["Aaaaa"])
    with pytest.raises(RuntimeError, match="Template placeholders were not fully replaced: Aaaaa"):
        mailer._render_formatted_body_to_clipboard_via_word_com(DraftMessage("to@example.com", "", "s", source, {"Aaaaa": "Li"}))


def test_confirmation_happens_before_send_with_fake_mailer():
    d = DraftMessage("to@example.com", "cc@example.com", "s", template_path=Path("templates/NB_Template_Insight.docx"), placeholders={})
    calls = {"send": 0, "confirm": 0}

    class FakeOutlookMailer(OutlookMailer):
        def create_draft_and_maybe_send(self, draft, confirm):
            calls["confirm"] += 1
            if not confirm("确认发送邮件？"):
                return False
            calls["send"] += 1
            return True

    m = FakeOutlookMailer()
    assert m.create_draft_and_maybe_send(d, lambda _: False) is False
    assert calls["confirm"] == 1
    assert calls["send"] == 0

def test_create_draft_displays_only_after_render_success(monkeypatch, tmp_path: Path):
    calls = []

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

    class Mail:
        def __init__(self):
            self.To = self.CC = self.Subject = ""
            self.SendUsingAccount = None

        def Display(self):
            calls.append("display")

        def Send(self):
            calls.append("send")

    class Outlook:
        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            return Mail()

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return Outlook()

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_is_windows", lambda: True)
    monkeypatch.setattr(mailer, "_render_formatted_body_to_clipboard_via_word_com", lambda _d: calls.append("render"))
    monkeypatch.setattr(mailer, "_paste_clipboard_into_mail", lambda _m: calls.append("paste"))

    sent = mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: True)
    assert sent is True
    assert calls[:3] == ["render", "display", "paste"]


def test_create_draft_does_not_display_when_render_fails(monkeypatch, tmp_path: Path):
    calls = []

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

    class Mail:
        def Display(self):
            calls.append("display")

    class Outlook:
        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            return Mail()

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return Outlook()

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_is_windows", lambda: True)
    monkeypatch.setattr(mailer, "_render_formatted_body_to_clipboard_via_word_com", lambda _d: (_ for _ in ()).throw(RuntimeError("Template placeholders were not fully replaced: Aaaaa")))
    with pytest.raises(RuntimeError, match="Template placeholders"):
        mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: True)
    assert calls == []


def test_ooxml_fallback_replaces_split_placeholder_runs(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>Aaa</w:t><w:t>aa</w:t></w:document>')
    rendered = OutlookMailer()._render_template_docx_with_ooxml(source, {"Aaaaa": "Li"})
    try:
        assert OutlookMailer()._remaining_placeholders_in_docx(rendered, {"Aaaaa": "Li"}) == []
    finally:
        rendered.unlink(missing_ok=True)


def test_ooxml_fallback_preserves_xml_prefixes_and_attributes(tmp_path: Path):
    source = tmp_path / "template.docx"
    original = '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p custom="1"><w:r><w:t xml:space="preserve">Aaaaa</w:t></w:r></w:p></w:document>'
    _make_docx(source, original)
    rendered = OutlookMailer()._render_template_docx_with_ooxml(source, {"Aaaaa": "Li"})
    try:
        with zipfile.ZipFile(rendered, "r") as zf:
            text = zf.read("word/document.xml").decode("utf-8")
        assert "w:document" in text
        assert 'xml:space="preserve"' in text
        assert 'custom="1"' in text
    finally:
        rendered.unlink(missing_ok=True)


def test_ooxml_escape_special_characters(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>Aaaaa</w:t></w:document>')
    rendered = OutlookMailer()._render_template_docx_with_ooxml(source, {"Aaaaa": 'A&B<>"\''})
    try:
        with zipfile.ZipFile(rendered, "r") as zf:
            text = zf.read("word/document.xml").decode("utf-8")
        assert "A&amp;B&lt;&gt;&quot;&apos;" in text
    finally:
        rendered.unlink(missing_ok=True)


def test_ooxml_fallback_result_zip_is_valid(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>Aaaaa</w:t></w:document>')
    rendered = OutlookMailer()._render_template_docx_with_ooxml(source, {"Aaaaa": "Li"})
    try:
        with zipfile.ZipFile(rendered, "r") as zf:
            assert zf.testzip() is None
    finally:
        rendered.unlink(missing_ok=True)


def test_render_formatted_body_fallback_word_open_failure_raises_clear_error(monkeypatch, tmp_path: Path):
    source = tmp_path / "template.docx"
    rendered = tmp_path / "rendered.docx"
    _make_docx(source, "<w:t>Aaaaa</w:t>")
    _make_docx(rendered, "<w:t>done</w:t>")

    class FakeDocuments:
        def Open(self, p):
            if Path(p) == rendered:
                raise RuntimeError("Word says corrupt")
            return types.SimpleNamespace(Content=types.SimpleNamespace(Copy=lambda: None), Close=lambda *_: None)

    class FakeWord:
        def __init__(self):
            self.Visible = False
            self.Documents = FakeDocuments()

        def Quit(self):
            return None

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return FakeWord()

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)
    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_render_template_docx_with_ooxml", lambda *_: rendered)
    monkeypatch.setattr(mailer, "_replace_placeholders_in_word_doc", lambda *_: None)
    states = iter([["Aaaaa"]])
    monkeypatch.setattr(mailer, "_remaining_placeholders_in_word_doc", lambda _doc: next(states, []))
    monkeypatch.setattr(mailer, "_remaining_placeholders_in_docx", lambda *_: [])
    with pytest.raises(RuntimeError, match="Rendered Word template could not be opened"):
        mailer._render_formatted_body_to_clipboard_via_word_com(DraftMessage("to@example.com", "", "s", source, {"Aaaaa": "Li"}))
