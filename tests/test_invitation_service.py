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
    def __init__(self, row):
        self.row = row
        self.written = []
        self.saved = 0
        self.save_path = "fake.xlsm"
        self.fail_save = False

    def get_current_row(self):
        return self.row

    def set_row_value(self, row_number, field, value):
        self.written.append((row_number, field, value))
        self.row.values[field] = value

    def save_workbook(self, _):
        if self.fail_save:
            raise RuntimeError("Failed to save workbook. Please close NB_Author_2026.xlsm in Excel and try again.")
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
    controller2 = FakeController(row2)
    mailer2 = FakeMailer()
    svc2 = InvitationService(controller2, renderer, mailer2, lambda _: True, today_provider=lambda: date(2026, 5, 16))
    res2 = svc2.invite_current()
    assert "已邀请过" in res2.message
    assert res2.status == "Skipped"
    assert mailer2.drafts == []
    assert controller2.written == []

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
    assert controller_ok.saved == 2
    assert controller_ok.written and controller_ok.written[0][1] == "Date of Invitaion"
    assert mailer_ok.drafts[0].cc_email == "first@example.com"

    row_nocc = Row(3, {**base_values(), "First Author Email": ""})
    mailer_nocc = FakeMailer()
    InvitationService(FakeController(row_nocc), renderer, mailer_nocc, lambda _: True, today_provider=lambda: date(2026, 5, 16)).invite_current()
    assert mailer_nocc.drafts[0].cc_email == ""


def test_missing_save_path_preflight_blocks_send_and_write():
    renderer = TemplateRenderer()
    row_ok = Row(2, base_values())
    controller = FakeController(row_ok)
    controller.save_path = None
    svc = InvitationService(controller, renderer, FakeMailer(), lambda _: True, today_provider=lambda: date(2026, 5, 16))
    res = svc.invite_current()
    assert res.status == "Error"
    assert "save_path" in res.message
    assert "无法发送" in res.message
    assert svc.mailer.drafts == []
    assert controller.written == []


def test_preflight_save_failure_blocks_renderer_and_mailer_and_date_write():
    class FakeRenderer:
        def __init__(self):
            self.calls = 0

        def render_for_row(self, *_):
            self.calls += 1
            raise AssertionError("renderer should not be called")

    row_ok = Row(2, base_values())
    controller = FakeController(row_ok)
    controller.fail_save = True
    renderer = FakeRenderer()
    mailer = FakeMailer()
    svc = InvitationService(controller, renderer, mailer, lambda _: True, today_provider=lambda: date(2026, 5, 16))

    res = svc.invite_current()
    assert res.status == "Error"
    assert "Please close NB_Author_2026.xlsm" in res.message
    assert renderer.calls == 0
    assert mailer.drafts == []
    assert controller.written == []
