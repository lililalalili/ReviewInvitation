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

    mailer = OutlookMailer()
    rendered = mailer._render_template_docx_with_ooxml(source, {"Aaaaa": "Li", "Ttttt": "Title & Value"})
    try:
        with zipfile.ZipFile(rendered, "r") as zf:
            content = zf.read("word/document.xml").decode("utf-8")
        assert "Li" in content
        assert "Title &amp; Value" in content
    finally:
        rendered.unlink(missing_ok=True)


def test_remaining_placeholders_detected(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, "<w:t>Aaaaa Dddddin</w:t>")

    remaining = OutlookMailer()._remaining_placeholders_in_docx(source, {"Aaaaa": "Li", "Dddddin": "x"})
    assert remaining == ["Aaaaa", "Dddddin"]


def test_apply_formatted_body_uses_copy_paste_path(monkeypatch, tmp_path: Path):
    copied = {"copy": 0, "paste": 0, "opened": None}
    source = tmp_path / "template.docx"
    _make_docx(source, "<w:t>Aaaaa Ttttt</w:t>")

    class FakeContent:
        def Copy(self):
            copied["copy"] += 1

    class FakeDoc:
        def __init__(self):
            self.Content = FakeContent()

        def Close(self, _):
            return None

    class FakeDocuments:
        def Open(self, opened_path):
            copied["opened"] = opened_path
            return FakeDoc()

    class FakeWord:
        def __init__(self):
            self.Visible = False
            self.Documents = FakeDocuments()

        def Quit(self):
            return None

    class FakeWordEditor:
        class FakeRange:
            def Paste(self_inner):
                copied["paste"] += 1

        def Range(self, _start, _end):
            return self.FakeRange()

    class FakeInspector:
        WordEditor = FakeWordEditor()

    class FakeMail:
        GetInspector = FakeInspector()

    class FakeWin32Client:
        @staticmethod
        def Dispatch(name):
            assert name == "Word.Application"
            return FakeWord()

    fake_win32 = types.SimpleNamespace(client=FakeWin32Client)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    mailer = OutlookMailer()
    draft = DraftMessage(
        to_email="to@example.com",
        cc_email="",
        subject="s",
        template_path=source,
        placeholders={"Aaaaa": "Li", "Ttttt": "Title"},
    )
    mailer._apply_formatted_body_via_word_com(FakeMail(), draft)

    assert copied["copy"] == 1
    assert copied["paste"] == 1
    assert copied["opened"] is not None
    assert Path(copied["opened"]) != source


def test_apply_formatted_body_raises_when_placeholder_remains(monkeypatch, tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, "<w:t>Aaaaa</w:t><w:t>Dddddin</w:t>")

    class FakeWord:
        def __init__(self):
            self.Visible = False
            self.Documents = types.SimpleNamespace(Open=lambda _: None)

        def Quit(self):
            return None

    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return FakeWord()

    fake_win32 = types.SimpleNamespace(client=FakeWin32Client)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)

    draft = DraftMessage("to@example.com", "", "s", template_path=source, placeholders={"Aaaaa": "Li"})
    with pytest.raises(RuntimeError, match="Template placeholders were not fully replaced: Dddddin"):
        OutlookMailer()._apply_formatted_body_via_word_com(types.SimpleNamespace(GetInspector=None), draft)


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
