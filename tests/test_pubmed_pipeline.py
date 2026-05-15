from datetime import date

from nb_review_invitation_agent.pubmed_pipeline import build_pubmed_query


def test_build_pubmed_query_has_expected_sections():
    q = build_pubmed_query(date(2026, 1, 1), date(2026, 1, 31))
    assert '"Nature"[Journal]' in q
    assert '"2026/01/01"[Date - Publication]' in q
    assert 'NOT (erratum[pt] OR editorial[pt] OR comment[pt] OR review[pt])' in q
