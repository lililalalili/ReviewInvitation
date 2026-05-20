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


def _install_fake_outlook(monkeypatch, outlook_obj):
    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return outlook_obj

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)


def test_sender_account_selected_and_properties_set(monkeypatch, tmp_path: Path):
    calls = []

    class Folder:
        def __init__(self):
            self.Items = types.SimpleNamespace(Add=lambda _msg: Mail("from_drafts"))

    class DeliveryStore:
        def GetDefaultFolder(self, folder_id):
            if folder_id == 16:
                return Folder()
            if folder_id == 5:
                return "sent-folder"
            raise RuntimeError("unexpected")

    class Account:
        def __init__(self, smtp):
            self.SmtpAddress = smtp
            self.DeliveryStore = DeliveryStore()

    class Mail:
        def __init__(self, source):
            self.source = source
            self.Body = ""
            self.SendUsingAccount = None
            self.SaveSentMessageFolder = None

        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == "HTMLBody":
                object.__setattr__(self, "Body", "rendered enough body length text")
                calls.append("html")

        def Display(self):
            calls.append("display")

        def Send(self):
            calls.append("send")

    class Outlook:
        def __init__(self):
            self.created = []
            self.accounts = [Account("other@example.com"), Account("NSB@ION.AC.CN")]

        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=self.accounts)

        def CreateItem(self, _):
            self.created.append("fallback")
            return Mail("fallback")

    outlook = Outlook()
    _install_fake_outlook(monkeypatch, outlook)
    monkeypatch.setenv("OUTLOOK_SEND_ACCOUNT", "nsb@ion.ac.cn")
    monkeypatch.setenv("OUTLOOK_FORCE_FROM_ADDRESS", "nsb@ion.ac.cn")

    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_is_windows", lambda: True)
    monkeypatch.setattr(mailer, "_render_template_to_filtered_html_via_word_com", lambda _d: "<html><body>hello world body content</body></html>")

    assert mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: False) is False
    assert "display" in calls
    assert "send" not in calls
    assert outlook.created == []


def test_drafts_creation_fallbacks_to_create_item(monkeypatch, tmp_path: Path):
    class DeliveryStore:
        def GetDefaultFolder(self, _folder_id):
            raise RuntimeError("cannot access drafts")

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

        def __init__(self):
            self.DeliveryStore = DeliveryStore()

    class Mail:
        Body = "good enough body text for validation"

        def Display(self): ...

        def Send(self): ...

    class Outlook:
        def __init__(self):
            self.create_item_called = 0

        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            self.create_item_called += 1
            return Mail()

    outlook = Outlook()
    _install_fake_outlook(monkeypatch, outlook)
    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_is_windows", lambda: True)
    monkeypatch.setattr(mailer, "_render_template_to_filtered_html_via_word_com", lambda _d: "<html><body>body with enough length for rendering</body></html>")
    mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: False)
    assert outlook.create_item_called == 1


def test_no_matching_account_raises_before_display(monkeypatch, tmp_path: Path):
    class Account:
        SmtpAddress = "other@example.com"

    class Outlook:
        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            raise AssertionError("should not create mail")

    _install_fake_outlook(monkeypatch, Outlook())
    monkeypatch.setenv("OUTLOOK_SEND_ACCOUNT", "nsb@ion.ac.cn")
    mailer = OutlookMailer()
    with pytest.raises(RuntimeError, match="Outlook account nsb@ion.ac.cn not found"):
        mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: True)


def test_set_sent_on_behalf_error_raised_before_send(monkeypatch, tmp_path: Path):
    class Folder:
        Items = types.SimpleNamespace(Add=lambda _msg: Mail())

    class DeliveryStore:
        def GetDefaultFolder(self, _id):
            return Folder()

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

        def __init__(self):
            self.DeliveryStore = DeliveryStore()

    class Mail:
        Body = "good enough body text for validation"

        def __setattr__(self, name, value):
            if name == "SentOnBehalfOfName":
                raise RuntimeError("no permission")
            object.__setattr__(self, name, value)

        def Display(self):
            raise AssertionError("must not display")

    class Outlook:
        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            return Mail()

    _install_fake_outlook(monkeypatch, Outlook())
    mailer = OutlookMailer()
    monkeypatch.setattr(mailer, "_is_windows", lambda: True)
    monkeypatch.setattr(mailer, "_render_template_to_filtered_html_via_word_com", lambda _d: "<html><body>body with enough length for rendering</body></html>")
    with pytest.raises(RuntimeError, match="Unable to set Outlook sender identity"):
        mailer.create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.docx", {}), lambda _: True)


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
