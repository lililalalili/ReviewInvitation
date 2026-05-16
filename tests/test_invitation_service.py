from datetime import date

from nb_review_invitation_agent.invitation_service import InvitationService
from nb_review_invitation_agent.template_renderer import TemplateRenderer


class FakeMailer:
    def __init__(self, send_result=True, raise_error=False):
        self.send_result = send_result
        self.raise_error = raise_error
        self.drafts = []

    def create_draft_and_maybe_send(self, draft, confirm):
        self.drafts.append(draft)
        if self.raise_error:
            raise RuntimeError("send failed")
        return self.send_result and confirm("ok")


class FakeController:
    def __init__(self, row, save_path="fake.xlsm"):
        self.row = row
        self.written = []
        self.saved = 0
        self.save_path = save_path

    def get_current_row(self):
        return self.row

    def set_row_value(self, row_number, field, value):
        self.written.append((row_number, field, value))
        self.row.values[field] = value

    def save_workbook(self, _):
        self.saved += 1


class Row:
    def __init__(self, row_number, values):
        self.row_number = row_number
        self.values = values


def base_values():
    return {
        "Manual Decision": "Review",
        "Overseas": "Yes",
        "Email of the Last Author": "last@example.com",
        "First Author Email": "first@example.com",
        "Date of Invitaion": "",
        "Journal": "nature neuroscience",
        "Title": "Title.",
        "Research field": "RF",
        "Full Name of the Last Author": "Alice Smith",
    }


def test_skip_manual_no_and_already_invited_and_missing_recipient():
    renderer = TemplateRenderer()
    row1 = Row(2, {**base_values(), "Manual Decision": "No"})
    svc1 = InvitationService(FakeController(row1), renderer, FakeMailer(), lambda _: True, today_provider=lambda: date(2026, 5, 16))
    assert svc1.invite_current().status == "Skipped"

    row2 = Row(2, {**base_values(), "Date of Invitaion": "2026-01-01"})
    svc2 = InvitationService(FakeController(row2), renderer, FakeMailer(), lambda _: True, today_provider=lambda: date(2026, 5, 16))
    assert "已邀请过" in svc2.invite_current().message

    row3 = Row(2, {**base_values(), "Email of the Last Author": ""})
    svc3 = InvitationService(FakeController(row3), renderer, FakeMailer(), lambda _: True, today_provider=lambda: date(2026, 5, 16))
    assert svc3.invite_current().status == "Error"


def test_confirmation_and_send_failures_and_success():
    renderer = TemplateRenderer()

    row_cancel = Row(2, base_values())
    controller_cancel = FakeController(row_cancel)
    svc_cancel = InvitationService(controller_cancel, renderer, FakeMailer(), lambda _: False, today_provider=lambda: date(2026, 5, 16))
    res_cancel = svc_cancel.invite_current()
    assert res_cancel.status == "Cancelled"
    assert controller_cancel.written == []

    row_fail = Row(2, base_values())
    controller_fail = FakeController(row_fail)
    svc_fail = InvitationService(controller_fail, renderer, FakeMailer(raise_error=True), lambda _: True, today_provider=lambda: date(2026, 5, 16))
    res_fail = svc_fail.invite_current()
    assert res_fail.status == "Error"
    assert controller_fail.written == []

    row_ok = Row(2, base_values())
    controller_ok = FakeController(row_ok)
    mailer_ok = FakeMailer()
    svc_ok = InvitationService(controller_ok, renderer, mailer_ok, lambda _: True, today_provider=lambda: date(2026, 5, 16))
    res_ok = svc_ok.invite_current()
    assert res_ok.status == "Sent"
    assert controller_ok.saved == 1
    assert controller_ok.written and controller_ok.written[0][1] == "Date of Invitaion"
    assert mailer_ok.drafts[0].cc_email == "first@example.com"
    assert mailer_ok.drafts[0].rendered.template_name == "NB_Template_Review_Yes.docx"

    row_nocc = Row(3, {**base_values(), "First Author Email": ""})
    mailer_nocc = FakeMailer()
    InvitationService(FakeController(row_nocc), renderer, mailer_nocc, lambda _: True, today_provider=lambda: date(2026, 5, 16)).invite_current()
    assert mailer_nocc.drafts[0].cc_email == ""


def test_missing_save_path_returns_clear_error():
    renderer = TemplateRenderer()
    row_ok = Row(2, base_values())
    controller = FakeController(row_ok, save_path=None)
    svc = InvitationService(controller, renderer, FakeMailer(), lambda _: True, today_provider=lambda: date(2026, 5, 16))
    res = svc.invite_current()
    assert res.status == "Error"
    assert "save path is missing" in res.message
    assert controller.written == []
