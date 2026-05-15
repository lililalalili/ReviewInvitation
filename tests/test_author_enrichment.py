import json

from nb_review_invitation_agent.author_enrichment import AuthorEnricher, clean_space, extract_email, mask_email
from nb_review_invitation_agent.deepseek_client import DeepSeekConfig, FakeDeepSeekClient
from nb_review_invitation_agent.search_provider import FakeSearchProvider, SearchResult


def _base_row():
    return {
        "Full Name of the Last Author": "Alice Smith",
        "Email of the Last Author": "",
        "Affiliation of the Last Author": "Example University",
        "First Author Full name": "Bob Chen",
        "Affiliation of the First Author": "Example Institute",
        "First Author Email": "",
        "Date of Invitaion": "",
        "Manual Decision": "Do not include",
    }


def test_email_extract_and_mask():
    text = "Electronic address: alice.smith@example.edu"
    email = extract_email(text)
    assert email == "alice.smith@example.edu"
    assert mask_email(email) == "a***@example.edu"


def test_clean_space():
    assert clean_space("a   b\n c") == "a b c"


def test_deepseek_default_model_is_v4_pro():
    cfg = DeepSeekConfig(api_key="k")
    assert cfg.model == "deepseek-v4-pro"


def test_high_confidence_populates_web_research_and_blank_emails_only():
    row = _base_row()
    search = FakeSearchProvider(
        {
            "Alice Smith Example University": [SearchResult("Prof Alice", "https://u.edu/alice", "Neuroscience lab")],
            "Bob Chen Example Institute": [SearchResult("Bob", "https://i.edu/bob", "postdoc")],
        }
    )
    llm = FakeDeepSeekClient(
        {
            "personal_web_url": "https://u.edu/alice",
            "research_quote": "Studies synaptic plasticity.",
            "last_author_email": "alice@u.edu",
            "first_author_email": "bob@i.edu",
            "confidence": 0.92,
            "evidence": ["official faculty page"],
            "status": "Enriched",
        }
    )
    out = AuthorEnricher(search_provider=search, llm_client=llm).enrich_row(row)
    assert out["Last Author Web"] == "https://u.edu/alice"
    assert out["Last Author Research"] == "Studies synaptic plasticity."
    assert out["Email of the Last Author"] == "alice@u.edu"
    assert out["First Author Email"] == "bob@i.edu"
    assert out["Author Enrichment Status"] == "Enriched"
    assert out["Date of Invitaion"] == ""


def test_existing_emails_not_overwritten():
    row = _base_row()
    row["Email of the Last Author"] = "existing_last@u.edu"
    row["First Author Email"] = "existing_first@i.edu"
    llm = FakeDeepSeekClient({"last_author_email": "new_last@u.edu", "first_author_email": "new_first@i.edu", "confidence": 0.99})
    out = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=llm).enrich_row(row)
    assert out["Email of the Last Author"] == "existing_last@u.edu"
    assert out["First Author Email"] == "existing_first@i.edu"


def test_blank_first_author_email_generates_first_author_queries():
    row = _base_row()
    enricher = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=FakeDeepSeekClient({"confidence": 0.0}))
    queries = enricher._build_queries(row)
    assert "Alice Smith Example University" in queries
    assert "Bob Chen Example Institute" in queries


def test_existing_first_author_email_skips_first_author_query():
    row = _base_row()
    row["First Author Email"] = "present@i.edu"
    enricher = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=FakeDeepSeekClient({"confidence": 0.0}))
    queries = enricher._build_queries(row)
    assert "Bob Chen Example Institute" not in queries


def test_blank_names_do_not_generate_noisy_queries():
    row = _base_row()
    row["Full Name of the Last Author"] = ""
    row["Affiliation of the Last Author"] = ""
    row["First Author Full name"] = "  "
    row["Affiliation of the First Author"] = ""
    enricher = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=FakeDeepSeekClient({"confidence": 0.0}))
    queries = enricher._build_queries(row)
    assert queries == []


def test_prompt_uses_sanitized_subset_and_excludes_invitation_date():
    row = _base_row()
    row["Email of the Last Author"] = "alice@u.edu"
    llm = FakeDeepSeekClient({"confidence": 0.1})
    AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=llm).enrich_row(row)
    assert "Date of Invitaion" not in llm.last_user_prompt
    assert "Manual Decision" not in llm.last_user_prompt
    prompt_json = llm.last_user_prompt.split("Author row: ", 1)[1].split("\nCandidates:", 1)[0]
    payload = json.loads(prompt_json)
    assert sorted(payload.keys()) == sorted(
        [
            "Full Name of the Last Author",
            "Email of the Last Author",
            "Affiliation of the Last Author",
            "First Author Full name",
            "Affiliation of the First Author",
            "First Author Email",
        ]
    )


def test_low_confidence_marks_needs_review_and_does_not_write_uncertain_data():
    row = _base_row()
    llm = FakeDeepSeekClient({"personal_web_url": "https://bad.example", "last_author_email": "maybe@example.com", "confidence": 0.2, "evidence": ["weak"]})
    out = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=llm).enrich_row(row)
    assert out["Author Enrichment Status"] == "Needs Review"
    assert out.get("Last Author Web", "") in ("", row.get("Last Author Web", ""))
    assert out["Email of the Last Author"] == ""
    assert out["Date of Invitaion"] == ""


def test_provider_error_captured():
    row = _base_row()
    out = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=FakeDeepSeekClient(RuntimeError("boom"))).enrich_row(row)
    assert out["Author Enrichment Status"] == "Error"
    assert "boom" in out["Author Enrichment Evidence"]
    assert out["Date of Invitaion"] == ""


def test_evidence_string_becomes_single_item_and_percentage_confidence_is_high():
    row = _base_row()
    llm = FakeDeepSeekClient(
        {
            "personal_web_url": "https://u.edu/alice",
            "research_quote": "Studies synaptic plasticity.",
            "confidence": "92%",
            "evidence": "official faculty profile",
        }
    )
    out = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=llm).enrich_row(row)
    assert out["Author Enrichment Status"] == "Enriched"
    payload = json.loads(out["Author Enrichment Evidence"])
    assert payload["evidence"] == ["official faculty profile"]


def test_evidence_none_does_not_crash_and_invitation_date_preserved_when_non_empty():
    row = _base_row()
    row["Date of Invitaion"] = "2026-05-01"
    llm = FakeDeepSeekClient({"confidence": "", "evidence": None})
    out = AuthorEnricher(search_provider=FakeSearchProvider(), llm_client=llm).enrich_row(row)
    assert out["Author Enrichment Status"] == "Needs Review"
    payload = json.loads(out["Author Enrichment Evidence"])
    assert payload["evidence"] == []
    assert out["Date of Invitaion"] == "2026-05-01"
