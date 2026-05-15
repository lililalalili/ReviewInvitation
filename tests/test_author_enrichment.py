from nb_review_invitation_agent.author_enrichment import clean_space, extract_email, mask_email


def test_email_extract_and_mask():
    text = "Electronic address: alice.smith@example.edu"
    email = extract_email(text)
    assert email == "alice.smith@example.edu"
    assert mask_email(email) == "a***@example.edu"


def test_clean_space():
    assert clean_space("a   b\n c") == "a b c"
