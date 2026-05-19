from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tempfile
from typing import Protocol
from xml.sax.saxutils import escape
import zipfile


KNOWN_PLACEHOLDERS = (
    "Aaaaa",
    "Jjjjj",
    "Ttttt",
    "Fffff",
    "Pppppyes",
    "Pppppno",
    "Dddddre",
    "Dddddin",
)


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

    def _render_template_docx_with_ooxml(self, template_path: Path, placeholders: dict[str, str]) -> Path:
        temp_file = tempfile.NamedTemporaryFile(prefix="nb_rendered_", suffix=".docx", delete=False)
        temp_path = Path(temp_file.name)
        temp_file.close()

        with zipfile.ZipFile(template_path, "r") as source_zip, zipfile.ZipFile(temp_path, "w") as dest_zip:
            for item in source_zip.infolist():
                content = source_zip.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    xml_text = content.decode("utf-8")
                    for key, value in placeholders.items():
                        xml_text = xml_text.replace(key, escape("" if value is None else str(value), {'"': '&quot;', "'": '&apos;' }))
                    content = xml_text.encode("utf-8")
                dest_zip.writestr(item, content)

        return temp_path

    def _remaining_placeholders_in_docx(self, docx_path: Path, placeholders: dict[str, str]) -> list[str]:
        remaining: set[str] = set()
        keys = tuple(dict.fromkeys([*KNOWN_PLACEHOLDERS, *placeholders.keys()]))

        with zipfile.ZipFile(docx_path, "r") as zf:
            for item in zf.infolist():
                if not (item.filename.startswith("word/") and item.filename.endswith(".xml")):
                    continue
                xml_text = zf.read(item.filename).decode("utf-8")
                for key in keys:
                    if key in xml_text:
                        remaining.add(key)

        return sorted(remaining)

    def _apply_formatted_body_via_word_com(self, mail, draft: DraftMessage) -> None:  # pragma: no cover - windows-only
        import win32com.client  # type: ignore[import-not-found]

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = None
        rendered_template_path = self._render_template_docx_with_ooxml(draft.template_path, draft.placeholders)
        try:
            remaining = self._remaining_placeholders_in_docx(rendered_template_path, draft.placeholders)
            if remaining:
                raise RuntimeError(f"Template placeholders were not fully replaced: {', '.join(remaining)}")

            doc = word.Documents.Open(str(rendered_template_path))
            doc.Content.Copy()
            inspector = mail.GetInspector
            editor = inspector.WordEditor
            editor.Range(0, 0).Paste()
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()
            rendered_template_path.unlink(missing_ok=True)
