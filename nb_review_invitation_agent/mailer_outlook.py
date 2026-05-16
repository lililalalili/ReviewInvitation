from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from .template_renderer import RenderedInvitation


@dataclass
class DraftMessage:
    to_email: str
    cc_email: str
    rendered: RenderedInvitation


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
        mail.Subject = draft.rendered.subject
        mail.SendUsingAccount = sender
        self._render_template_into_mail_body(mail, draft.rendered)
        mail.Display()

        if not confirm("确认发送邮件？"):
            return False

        mail.Send()
        return True

    def _render_template_into_mail_body(self, mail, rendered: RenderedInvitation) -> None:
        try:
            import win32com.client  # type: ignore[import-not-found]
        except Exception as exc:  # pragma: no cover - windows-only
            raise RuntimeError("pywin32/win32com is required on Windows for Word template rendering") from exc

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(Path(rendered.template_path).resolve()))
        try:
            for key, value in rendered.placeholders.items():
                find = doc.Content.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                find.Execute(FindText=key, ReplaceWith=value, Replace=2)
            mail.HTMLBody = doc.Content.HTML + mail.HTMLBody
        finally:
            doc.Close(False)
            word.Quit()
