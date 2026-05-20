from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tempfile
from typing import Protocol
import zipfile
from html import unescape as html_unescape
import re
from xml.sax.saxutils import escape


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
TEXT_NODE_PATTERN = re.compile(r"(<(?P<tag>w:t|a:t)\b[^>]*>)(?P<text>.*?)(</(?P=tag)>)", re.DOTALL)
WD_FORMAT_FILTERED_HTML = 10


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

        rendered_html = self._render_template_to_filtered_html_via_word_com(draft) if self._is_windows() else (draft.body_text or "")
        if not self._is_windows() and len(rendered_html.strip()) <= 20:
            raise RuntimeError("Rendered email body is empty.")

        mail = outlook.CreateItem(0)
        mail.To = draft.to_email
        mail.CC = draft.cc_email
        mail.Subject = draft.subject
        mail.SendUsingAccount = sender
        mail.HTMLBody = rendered_html
        self._validate_outlook_body(mail)
        mail.Display()

        if not confirm("确认发送邮件？"):
            return False

        mail.Send()
        return True

    def _is_windows(self) -> bool:
        import platform

        return platform.system().lower().startswith("win")

    def _replace_placeholders_in_range(self, word_range, placeholders: dict[str, str]) -> None:
        wd_find_continue = 1
        wd_replace_all = 2
        for key, value in placeholders.items():
            replacement = "" if value is None else str(value)
            find = word_range.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Execute(FindText=key, MatchCase=False, MatchWholeWord=False, MatchWildcards=False, Forward=True, Wrap=wd_find_continue, Format=False, ReplaceWith=replacement, Replace=wd_replace_all)

    def _replace_placeholders_in_shapes(self, shapes, placeholders: dict[str, str]) -> None:
        for shape in shapes:
            try:
                frame = getattr(shape, "TextFrame", None)
                if frame is None or not frame.HasText:
                    continue
                text_range = frame.TextRange
            except Exception:
                continue
            self._replace_placeholders_in_range(text_range, placeholders)

    def _replace_placeholders_in_word_doc(self, doc, placeholders: dict[str, str]) -> None:
        self._replace_placeholders_in_range(doc.Content, placeholders)
        for story_range in getattr(doc, "StoryRanges", []):
            current = story_range
            while current is not None:
                self._replace_placeholders_in_range(current, placeholders)
                current = getattr(current, "NextStoryRange", None)
        shapes = getattr(doc, "Shapes", None)
        if shapes is not None:
            self._replace_placeholders_in_shapes(shapes, placeholders)

    def _remaining_placeholders_in_word_doc(self, doc, known_placeholders: tuple[str, ...] = KNOWN_PLACEHOLDERS) -> list[str]:
        haystacks = [str(getattr(doc.Content, "Text", ""))]
        for story_range in getattr(doc, "StoryRanges", []):
            current = story_range
            while current is not None:
                haystacks.append(str(getattr(current, "Text", "")))
                current = getattr(current, "NextStoryRange", None)
        shapes = getattr(doc, "Shapes", None)
        if shapes is not None:
            for shape in shapes:
                try:
                    frame = getattr(shape, "TextFrame", None)
                    if frame is None or not frame.HasText:
                        continue
                    haystacks.append(str(getattr(frame.TextRange, "Text", "")))
                except Exception:
                    continue
        return sorted({k for k in known_placeholders if any(k in text for text in haystacks)})

    def _save_word_doc_as_filtered_html(self, doc) -> str:
        tmp = tempfile.NamedTemporaryFile(prefix="nb_rendered_", suffix=".htm", delete=False)
        tmp_path = Path(tmp.name)
        tmp.close()
        doc.SaveAs(str(tmp_path), FileFormat=WD_FORMAT_FILTERED_HTML)
        return str(tmp_path)

    def _validate_rendered_html(self, html: str) -> None:
        if len(html.strip()) <= 20:
            raise RuntimeError("Rendered email body is empty.")
        remaining = [k for k in KNOWN_PLACEHOLDERS if k in html]
        if remaining:
            raise RuntimeError(f"Rendered email body still contains placeholders: {', '.join(remaining)}")

    def _validate_outlook_body(self, mail) -> None:
        body = str(getattr(mail, "Body", "") or "")
        if len(body.strip()) <= 20:
            raise RuntimeError("Outlook email body is empty after HTMLBody assignment.")
        remaining = [k for k in KNOWN_PLACEHOLDERS if k in body]
        if remaining:
            raise RuntimeError(f"Outlook email body still contains placeholders: {', '.join(remaining)}")

    def _render_template_to_filtered_html_via_word_com(self, draft: DraftMessage) -> str:  # pragma: no cover - windows-only
        import win32com.client  # type: ignore[import-not-found]

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = None
        rendered_template_path: Path | None = None
        temp_html_path: Path | None = None
        try:
            doc = word.Documents.Open(str(draft.template_path))
            self._replace_placeholders_in_word_doc(doc, draft.placeholders)
            remaining = self._remaining_placeholders_in_word_doc(doc)
            if remaining:
                doc.Close(False)
                doc = None
                rendered_template_path = self._render_template_docx_with_ooxml(draft.template_path, draft.placeholders)
                if self._remaining_placeholders_in_docx(rendered_template_path, draft.placeholders):
                    raise RuntimeError(f"Rendered email body still contains placeholders: {', '.join(self._remaining_placeholders_in_docx(rendered_template_path, draft.placeholders))}")
                try:
                    doc = word.Documents.Open(str(rendered_template_path))
                except Exception as exc:
                    raise RuntimeError("Rendered Word template could not be opened. The generated document may be invalid.") from exc
                remaining = self._remaining_placeholders_in_word_doc(doc)
                if remaining:
                    raise RuntimeError(f"Rendered email body still contains placeholders: {', '.join(remaining)}")

            text = str(getattr(doc.Content, "Text", "") or "")
            if len(text.strip()) <= 20:
                raise RuntimeError("Rendered email body is empty.")
            temp_html_path = Path(self._save_word_doc_as_filtered_html(doc))
            html = temp_html_path.read_text(encoding="utf-8", errors="ignore")
            self._validate_rendered_html(html)
            return html
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()
            if rendered_template_path is not None:
                rendered_template_path.unlink(missing_ok=True)
            if temp_html_path is not None:
                temp_html_path.unlink(missing_ok=True)

    # optional fallback retained
    def _paste_clipboard_into_mail(self, mail) -> None:
        inspector = mail.GetInspector
        editor = inspector.WordEditor
        editor.Range(0, 0).Paste()

    def _extract_text_nodes(self, xml_text: str) -> list[dict[str, str | int]]:
        nodes: list[dict[str, str | int]] = []
        for match in TEXT_NODE_PATTERN.finditer(xml_text):
            text_raw = match.group("text")
            nodes.append({"start": match.start("text"), "end": match.end("text"), "raw": text_raw, "text": html_unescape(text_raw)})
        return nodes

    def _replace_placeholders_in_xml_text_nodes(self, xml_text: str, placeholders: dict[str, str]) -> str:
        nodes = self._extract_text_nodes(xml_text)
        if not nodes:
            return xml_text
        full_text = "".join(str(node["text"]) for node in nodes)
        for key, value in placeholders.items():
            full_text = full_text.replace(key, "" if value is None else str(value))
        rendered_node_texts = []
        cursor = 0
        for index, node in enumerate(nodes):
            source_len = len(str(node["text"]))
            if index == len(nodes) - 1:
                rendered_node_texts.append(full_text[cursor:])
            else:
                rendered_node_texts.append(full_text[cursor : cursor + source_len])
                cursor += source_len
        chunks = []
        last = 0
        for node, replacement_text in zip(nodes, rendered_node_texts):
            start, end = int(node["start"]), int(node["end"])
            chunks.append(xml_text[last:start])
            chunks.append(escape(replacement_text, {'"': "&quot;", "'": "&apos;"}))
            last = end
        chunks.append(xml_text[last:])
        return "".join(chunks)

    def _render_template_docx_with_ooxml(self, template_path: Path, placeholders: dict[str, str]) -> Path:
        temp_file = tempfile.NamedTemporaryFile(prefix="nb_rendered_", suffix=".docx", delete=False)
        temp_path = Path(temp_file.name)
        temp_file.close()
        with zipfile.ZipFile(template_path, "r") as source_zip, zipfile.ZipFile(temp_path, "w") as dest_zip:
            for item in source_zip.infolist():
                content = source_zip.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    xml_text = content.decode("utf-8")
                    if self._is_target_xml_part(item.filename):
                        xml_text = self._replace_placeholders_in_xml_text_nodes(xml_text, placeholders)
                    content = xml_text.encode("utf-8")
                dest_zip.writestr(item, content)
        with zipfile.ZipFile(temp_path, "r") as rendered_zip:
            if rendered_zip.testzip() is not None:
                raise RuntimeError("Rendered Word template could not be opened. The generated document may be invalid.")
        return temp_path

    def _is_target_xml_part(self, filename: str) -> bool:
        name = Path(filename).name
        return filename == "word/document.xml" or name.startswith("header") or name.startswith("footer") or filename in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}

    def _remaining_placeholders_in_docx(self, docx_path: Path, placeholders: dict[str, str]) -> list[str]:
        remaining: set[str] = set()
        keys = tuple(dict.fromkeys([*KNOWN_PLACEHOLDERS, *placeholders.keys()]))
        with zipfile.ZipFile(docx_path, "r") as zf:
            for item in zf.infolist():
                if not (item.filename.startswith("word/") and item.filename.endswith(".xml")):
                    continue
                xml_text = zf.read(item.filename).decode("utf-8")
                combined_text = "".join(str(node["text"]) for node in self._extract_text_nodes(xml_text))
                for key in keys:
                    if key in combined_text:
                        remaining.add(key)
        return sorted(remaining)
