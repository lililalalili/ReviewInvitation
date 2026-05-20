from pathlib import Path
import sys
import types

import pytest

import nb_review_invitation_agent.mailer_outlook as mailer_module
from nb_review_invitation_agent.mailer_outlook import DraftMessage, OutlookMailer


def test_win32com_not_imported_at_module_import_time():
    assert "win32com" not in mailer_module.__dict__


def _install_fake_outlook(monkeypatch, outlook_obj):
    class FakeWin32Client:
        @staticmethod
        def Dispatch(_name):
            return outlook_obj

    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=FakeWin32Client))
    monkeypatch.setitem(sys.modules, "win32com.client", FakeWin32Client)


def test_htmlbody_assigned_before_display_and_send_after_confirm(monkeypatch, tmp_path: Path):
    calls = []

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"
        DeliveryStore = types.SimpleNamespace(GetDefaultFolder=lambda _id: types.SimpleNamespace(Items=types.SimpleNamespace(Add=lambda _: Mail())))

    class Mail:
        Body = ""

        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == "HTMLBody":
                self.Body = "sufficient body text after html assignment"
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

    _install_fake_outlook(monkeypatch, Outlook())
    mailer = OutlookMailer()
    sent = mailer.create_draft_and_maybe_send(
        DraftMessage("to@example.com", "", "s", tmp_path / "x.html", {}, body_html="<html><body>hello world body content</body></html>"),
        lambda _: True,
    )
    assert sent is True
    assert calls == ["html", "display", "send"]


def test_display_not_called_if_html_rendering_fails(monkeypatch, tmp_path: Path):
    class Account:
        SmtpAddress = "nsb@ion.ac.cn"

    class Outlook:
        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            raise AssertionError("must not create draft")

    _install_fake_outlook(monkeypatch, Outlook())
    with pytest.raises(RuntimeError, match="Rendered email body is empty"):
        OutlookMailer().create_draft_and_maybe_send(DraftMessage("to@example.com", "", "s", tmp_path / "x.html", {}, body_html=""), lambda _: True)


def test_send_not_called_before_confirmation(monkeypatch, tmp_path: Path):
    calls = []

    class Account:
        SmtpAddress = "nsb@ion.ac.cn"
        DeliveryStore = types.SimpleNamespace(GetDefaultFolder=lambda _id: types.SimpleNamespace(Items=types.SimpleNamespace(Add=lambda _: Mail())))

    class Mail:
        Body = ""

        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == "HTMLBody":
                self.Body = "sufficient body text after html assignment"

        def Display(self):
            calls.append("display")

        def Send(self):
            calls.append("send")

    class Outlook:
        def GetNamespace(self, _):
            return types.SimpleNamespace(Accounts=[Account()])

        def CreateItem(self, _):
            return Mail()

    _install_fake_outlook(monkeypatch, Outlook())
    sent = OutlookMailer().create_draft_and_maybe_send(
        DraftMessage("to@example.com", "", "s", tmp_path / "x.html", {}, body_html="<html><body>hello world body content</body></html>"),
        lambda _: False,
    )
    assert sent is False
    assert calls == ["display"]
