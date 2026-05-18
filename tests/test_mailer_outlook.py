from pathlib import Path
import sys
import types

import nb_review_invitation_agent.mailer_outlook as mailer_module
from nb_review_invitation_agent.mailer_outlook import DraftMessage, OutlookMailer


def test_win32com_not_imported_at_module_import_time():
    assert "win32com" not in mailer_module.__dict__


def test_apply_formatted_body_uses_copy_paste_path(monkeypatch):
    copied = {"copy": 0, "paste": 0, "find": []}

    class FakeFind:
        def ClearFormatting(self):
            return None

        class ReplacementObj:
            def ClearFormatting(self):
                return None

        Replacement = ReplacementObj()

        def Execute(self, **kwargs):
            copied["find"].append(kwargs)

    class FakeContent:
        def __init__(self):
            self.Find = FakeFind()

        def Copy(self):
            copied["copy"] += 1

    class FakeDoc:
        def __init__(self):
            self.Content = FakeContent()

        def Close(self, _):
            return None

    class FakeDocuments:
        def Open(self, _):
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
        template_path=Path("templates/NB_Template_Insight.docx"),
        placeholders={"Aaaaa": "Li", "Ttttt": "Title"},
    )
    mailer._apply_formatted_body_via_word_com(FakeMail(), draft)

    assert copied["copy"] == 1
    assert copied["paste"] == 1
    assert len(copied["find"]) == 2


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
