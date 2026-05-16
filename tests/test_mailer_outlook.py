import sys
import types
import importlib
import inspect

from nb_review_invitation_agent.mailer_outlook import DraftMessage, OutlookMailer
from nb_review_invitation_agent.template_renderer import RenderedInvitation


def test_draft_message_dataclass():
    rendered = RenderedInvitation(template_path=__file__, template_name="x.docx", subject="s", placeholders={}, body_text="")
    d = DraftMessage("to@example.com", "cc@example.com", rendered)
    assert d.to_email == "to@example.com"


def test_mailer_has_lazy_import_in_method():
    import nb_review_invitation_agent.mailer_outlook as mod

    source = inspect.getsource(mod)
    assert "import win32com.client" in source
    assert "from win32com.client" not in source


def test_confirmation_required_before_send(monkeypatch):
    events = []

    class FakeMail:
        HTMLBody = ""
        def Display(self): events.append("display")
        def Send(self): events.append("send")

    class FakeNamespace:
        Accounts = [types.SimpleNamespace(SmtpAddress="nsb@ion.ac.cn")]

    class FakeOutlook:
        def GetNamespace(self, _): return FakeNamespace()
        def CreateItem(self, _): return FakeMail()

    class FakeWordDoc:
        def __init__(self):
            self.Content = types.SimpleNamespace(
                Find=types.SimpleNamespace(
                    ClearFormatting=lambda: None,
                    Replacement=types.SimpleNamespace(ClearFormatting=lambda: None),
                    Execute=lambda **kwargs: None,
                ),
                HTML="<p>x</p>",
            )
        def Close(self, _): pass

    class FakeWord:
        Visible = False
        Documents = types.SimpleNamespace(Open=lambda _: FakeWordDoc())
        def Quit(self): pass

    class FakeClient:
        def Dispatch(self, name):
            if name == "Outlook.Application":
                return FakeOutlook()
            return FakeWord()

    fake_mod = types.SimpleNamespace(client=FakeClient())
    monkeypatch.setitem(sys.modules, "win32com", fake_mod)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_mod.client)

    rendered = RenderedInvitation(template_path=__file__, template_name="x.docx", subject="S", placeholders={"Aaaaa": "Li"})
    draft = DraftMessage("to@example.com", "", rendered)
    sent = OutlookMailer().create_draft_and_maybe_send(draft, lambda _: False)
    assert sent is False
    assert events == ["display"]
