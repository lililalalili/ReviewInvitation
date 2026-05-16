from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol


@dataclass
class DraftMessage:
    to_email: str
    cc_email: str
    subject: str
    template_path: Path
    placeholders: dict[str, str]
    body_text: str | None = None


class ConfirmationProvider(Protocol):
    def __call__(self, prompt: str) -> bool: ...


class OutlookMailer:
    def create_draft_and_maybe_send(self, draft: DraftMessage, confirm: ConfirmationProvider) -> bool:
        try:
            import win32com.client  # type: ignore[import-not-found]
        except Exception as exc:  # pragma: no cover - windows-only
            raise RuntimeError("pywin32/win32com is required on Windows for Outlook automation") from exc

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        sender = None
        for account in namespace.Accounts:
            smtp = str(getattr(account, "SmtpAddress", "")).strip().lower()
            if smtp == "nsb@ion.ac.cn":
                sender = account
                break
        if sender is None:
            raise RuntimeError("Outlook account nsb@ion.ac.cn not found")

        mail = outlook.CreateItem(0)
        mail.To = draft.to_email
        mail.CC = draft.cc_email
        mail.Subject = draft.subject
        mail.SendUsingAccount = sender
        mail.Display()
        if self._is_windows():
            self._apply_formatted_body_via_word_com(mail, draft)
        else:
            mail.Body = draft.body_text or ""

        if not confirm("确认发送邮件？"):
            return False

        mail.Send()
        return True

    def _is_windows(self) -> bool:
        import platform

        return platform.system().lower().startswith("win")

    def _apply_formatted_body_via_word_com(self, mail, draft: DraftMessage) -> None:  # pragma: no cover - windows-only
        import win32com.client  # type: ignore[import-not-found]

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = None
        try:
            doc = word.Documents.Open(str(draft.template_path))
            for key, value in draft.placeholders.items():
                find = doc.Content.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                find.Execute(FindText=key, ReplaceWith=value, Replace=2)

            doc.Content.Copy()
            inspector = mail.GetInspector
            editor = inspector.WordEditor
            editor.Range(0, 0).Paste()
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()
