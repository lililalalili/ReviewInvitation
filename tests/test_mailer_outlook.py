from nb_review_invitation_agent.mailer_outlook import DraftMessage


def test_draft_message_dataclass():
    d = DraftMessage("to@example.com", "cc@example.com", "s", "b")
    assert d.to_email == "to@example.com"
