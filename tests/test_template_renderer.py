import pytest
from datetime import date
from pathlib import Path

from nb_review_invitation_agent.template_renderer import (
    TemplateRenderer,
    add_calendar_months,
    build_placeholder_map,
    derive_family_name,
    normalize_journal_name,
    select_template_and_subject,
)


def test_template_selection_cases():
    c1 = select_template_and_subject({"Manual Decision": "Review", "Overseas": "Yes"})
    assert c1.template_name == "NB_Template_Review_Yes.html"

    c2 = select_template_and_subject({"Manual Decision": "Review", "Overseas": "No"})
    assert c2.template_name == "NB_Template_Review_No.html"

    c3 = select_template_and_subject({"Manual Decision": "Insight", "Overseas": "No"})
    assert c3.template_name == "NB_Template_Insight.html"


def test_placeholders_and_dates():
    today = date(2026, 5, 16)
    m = build_placeholder_map(
        {
            "Family Name of the Last Author": "",
            "Full Name of the Last Author": "Jane van Rossum",
            "Journal": "nature neuroscience",
            "Title": "A great paper.",
            "Research field": "Systems",
            "Overseas": "Yes",
            "Manual Decision": "Review",
        },
        today,
    )
    assert m["Aaaaa"] == "Rossum"
    assert m["Jjjjj"] == "Nature Neuroscience"
    assert m["Ttttt"] == "A great paper"
    assert m["Fffff"] == "Systems"
    assert "1,000 USD" in m["Pppppyes"]
    assert m["Pppppno"] == ""
    assert m["Dddddre"] == "November 16, 2026"
    assert m["Dddddin"] == "September 18, 2026"


def test_add_calendar_months_end_of_month():
    assert add_calendar_months(date(2026, 8, 31), 6) == date(2027, 2, 28)


def test_derive_family_name_prefers_explicit():
    assert derive_family_name({"Family Name of the Last Author": "Li", "Full Name of the Last Author": "Mei Li"}) == "Li"


def test_insight_overseas_yes_keeps_pppppno():
    m = build_placeholder_map({"Manual Decision": "Insight", "Overseas": "Yes"}, date(2026, 5, 16))
    assert m["Pppppyes"] == ""
    assert m["Pppppno"] == "For invited authors, article publication charges will be covered by the journal."


def test_rendered_invitation_includes_template_path_and_placeholders():
    rendered = TemplateRenderer(templates_dir=Path("templates")).render_for_row(
        {"Manual Decision": "Insight", "Overseas": "Yes"},
        date(2026, 5, 16),
    )
    assert rendered.template_path.is_absolute()
    assert rendered.template_path.name == "NB_Template_Insight.html"
    assert rendered.placeholders["Pppppno"].startswith("For invited authors")
    assert rendered.rendered_html


def test_dynamic_values_are_escaped_in_rendered_html(tmp_path: Path):
    tpl = tmp_path / "NB_Template_Insight.html"
    tpl.write_text("<html><body>Aaaaa|Jjjjj|Ttttt|Fffff|Pppppno|Dddddin</body></html>", encoding="utf-8")
    rendered = TemplateRenderer(templates_dir=tmp_path).render_for_row({
        "Manual Decision": "Insight",
        "Overseas": "No",
        "Family Name of the Last Author": "A&B <test>",
        "Journal": "J&A <b>",
        "Title": "T&A <x>.",
        "Research field": "R&D <neuro>",
    }, date(2026, 5, 16))
    assert "A&amp;B &lt;test&gt;" in rendered.rendered_html
    assert "A&amp;" in rendered.rendered_html


def test_missing_html_template_raises_clear_error(tmp_path: Path):
    with pytest.raises(FileNotFoundError, match="Template not found"):
        TemplateRenderer(templates_dir=tmp_path).render_for_row({"Manual Decision": "Insight"}, date(2026, 5, 16))


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("nature neuroscience", "Nature Neuroscience"),
        ("NATURE NEUROSCIENCE", "Nature Neuroscience"),
        ("science (new york, n.y.)", "Science (New York, N.Y.)"),
        ("Science (New York, N.Y.)", "Science (New York, N.Y.)"),
        ("PNAS", "PNAS"),
        ("snyny", "SNYNY"),
        ("u.s.a.", "U.S.A."),
    ],
)
def test_normalize_journal_name(raw: str, expected: str):
    assert normalize_journal_name(raw) == expected


def test_rendered_html_uses_normalized_journal_placeholder(tmp_path: Path):
    tpl = tmp_path / "NB_Template_Insight.html"
    tpl.write_text("<html><body>Jjjjj</body></html>", encoding="utf-8")
    rendered = TemplateRenderer(templates_dir=tmp_path).render_for_row(
        {
            "Manual Decision": "Insight",
            "Overseas": "No",
            "Journal": "science (new york, n.y.)",
        },
        date(2026, 5, 16),
    )
    assert "Science (New York, N.Y.)" in rendered.rendered_html
