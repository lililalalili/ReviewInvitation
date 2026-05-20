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


def test_replace_placeholders_in_range_uses_replace_all():
    calls = []

    class FakeReplacement:
        def ClearFormatting(self): ...

    class FakeFind:
        Replacement = FakeReplacement()

        def ClearFormatting(self): ...

        def Execute(self, **kwargs):
            calls.append(kwargs)

    OutlookMailer()._replace_placeholders_in_range(types.SimpleNamespace(Find=FakeFind()), {"Aaaaa": "Li"})
    assert calls and calls[0]["Replace"] == 2


def test_create_draft_assigns_htmlbody_before_display(monkeypatch, tmp_path: Path):
    calls = []

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

    class Mail:
        def __init__(self):
            self.To = self.CC = self.Subject = ""
            self.SendUsingAccount = None
            self.Body = ""

        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == "HTMLBody":
                object.__setattr__(self, "Body", "plain content from html")
                calls.append("html")

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
    monkeypatch.setattr(mailer, "_render_template_to_filtered_html_via_word_com", lambda _d: "<html><body>hello world body content</body></html>")

    sent = mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: True)
    assert sent is True
    assert "html" in calls and "display" in calls
    assert calls.index("html") < calls.index("display")


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
    monkeypatch.setattr(mailer, "_render_template_to_filtered_html_via_word_com", lambda _d: (_ for _ in ()).throw(RuntimeError("Rendered email body is empty.")))
    with pytest.raises(RuntimeError, match="Rendered email body is empty"):
        mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: True)
    assert calls == []


def test_send_not_called_before_confirmation(monkeypatch, tmp_path: Path):
    sent = {"send": 0}

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

    class Mail:
        HTMLBody = ""
        Body = "body with enough length"

        def Display(self): ...

        def Send(self):
            sent["send"] += 1

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
    monkeypatch.setattr(mailer, "_render_template_to_filtered_html_via_word_com", lambda _d: "<html><body>body with enough length</body></html>")
    assert mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: False) is False
    assert sent["send"] == 0


def test_validate_rendered_html_failures():
    mailer = OutlookMailer()
    with pytest.raises(RuntimeError, match="Rendered email body is empty"):
        mailer._validate_rendered_html("  ")
    with pytest.raises(RuntimeError, match="still contains placeholders"):
        mailer._validate_rendered_html("<html><body>Aaaaa placeholder remains in this long html body text</body></html>")


def test_validate_outlook_body_failures():
    mailer = OutlookMailer()
    with pytest.raises(RuntimeError, match="Outlook email body is empty after HTMLBody assignment"):
        mailer._validate_outlook_body(types.SimpleNamespace(Body=""))
    with pytest.raises(RuntimeError, match="Outlook email body still contains placeholders"):
        mailer._validate_outlook_body(types.SimpleNamespace(Body="hello Aaaaa and enough long text for test"))


def test_ooxml_fallback_replaces_split_placeholder_runs(tmp_path: Path):
    source = tmp_path / "template.docx"
    _make_docx(source, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>Aaa</w:t><w:t>aa</w:t></w:document>')
    rendered = OutlookMailer()._render_template_docx_with_ooxml(source, {"Aaaaa": "Li"})
    try:
        assert OutlookMailer()._remaining_placeholders_in_docx(rendered, {"Aaaaa": "Li"}) == []
    finally:
        rendered.unlink(missing_ok=True)
