from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Callable

from .mailer_outlook import DraftMessage
from .template_renderer import TemplateRenderer
from .workbook import DATE_OF_INVITAION


@dataclass
class InvitationResult:
    status: str
    message: str
    row_number: int
    recipient: str = ""
    subject: str = ""
    template_name: str = ""
    date_written: bool = False


class InvitationService:
    def __init__(self, controller, renderer: TemplateRenderer, mailer, confirm_send: Callable[[str], bool], today_provider: Callable[[], date] | None = None) -> None:
        self.controller = controller
        self.renderer = renderer
        self.mailer = mailer
        self.confirm_send = confirm_send
        self.today_provider = today_provider or date.today

    def invite_current(self, row_context=None) -> InvitationResult:
        row = row_context or self.controller.get_current_row()
        rv = row.values
        row_number = row.row_number
        invited = rv.get(DATE_OF_INVITAION, "").strip()
        if rv.get("Manual Decision", "").strip() == "No":
            return InvitationResult("Skipped", "Manual Decision = No，请确认是否邀请", row_number)
        if invited:
            return InvitationResult("Skipped", "已邀请过", row_number)
        recipient = rv.get("Email of the Last Author", "").strip()
        if not recipient:
            return InvitationResult("Error", "缺少 Email of the Last Author", row_number)
        save_path = getattr(self.controller, "save_path", None)
        if not save_path:
            return InvitationResult("Error", "无法发送：controller.save_path 未设置", row_number, recipient=recipient)

        try:
            rendered = self.renderer.render_for_row(rv, self.today_provider())
        except Exception as exc:
            return InvitationResult("Error", str(exc), row_number, recipient=recipient)

        cc_email = rv.get("First Author Email", "").strip()
        draft = DraftMessage(
            to_email=recipient,
            cc_email=cc_email,
            subject=rendered.subject,
            template_path=rendered.template_path,
            placeholders=rendered.placeholders,
            body_text=rendered.body_text,
        )
        try:
            sent = self.mailer.create_draft_and_maybe_send(draft, self.confirm_send)
        except Exception as exc:
            return InvitationResult("Error", str(exc), row_number, recipient=recipient, subject=rendered.subject, template_name=rendered.template_name)
        if not sent:
            return InvitationResult("Cancelled", "用户取消发送", row_number, recipient=recipient, subject=rendered.subject, template_name=rendered.template_name)

        invite_date = self.today_provider().isoformat()
        try:
            self.controller.set_row_value(row_number, DATE_OF_INVITAION, invite_date)
            self.controller.save_workbook(save_path)
        except Exception as exc:
            return InvitationResult("Error", f"邮件已发送，但保存工作簿失败: {exc}", row_number, recipient=recipient, subject=rendered.subject, template_name=rendered.template_name)

        return InvitationResult("Sent", "发送成功", row_number, recipient=recipient, subject=rendered.subject, template_name=rendered.template_name, date_written=True)

    def invite_batch(self, batch_rows) -> list[InvitationResult]:
        results: list[InvitationResult] = []
        for row in batch_rows:
            result = self.invite_current(row)
            results.append(result)
            if result.status == "Error" and "nsb@ion.ac.cn" in result.message:
                break
        return results
