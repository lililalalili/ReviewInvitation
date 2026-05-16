from datetime import date

from nb_review_invitation_agent.template_renderer import (
    add_calendar_months,
    build_placeholder_map,
    derive_family_name,
    TemplateRenderer,
    select_template_and_subject,
)


def test_template_selection_cases():
    c1 = select_template_and_subject({"Manual Decision": "Review", "Overseas": "Yes"})
    assert c1.template_name == "NB_Template_Review_Yes.docx"

    c2 = select_template_and_subject({"Manual Decision": "Review", "Overseas": "No"})
    assert c2.template_name == "NB_Template_Review_No.docx"

    c3 = select_template_and_subject({"Manual Decision": "Insight", "Overseas": "No"})
    assert c3.template_name == "NB_Template_Insight.docx"


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

def test_insight_overseas_yes_uses_pppppno():
    today = date(2026, 5, 16)
    m = build_placeholder_map({"Manual Decision": "Insight", "Overseas": "Yes"}, today)
    assert m["Pppppyes"] == ""
    assert "publication charges will be covered by the journal" in m["Pppppno"]


def test_add_calendar_months_end_of_month():
    assert add_calendar_months(date(2026, 8, 31), 6) == date(2027, 2, 28)


def test_derive_family_name_prefers_explicit():
    assert derive_family_name({"Family Name of the Last Author": "Li", "Full Name of the Last Author": "Mei Li"}) == "Li"


def test_rendered_invitation_contains_template_path_and_placeholders():
    renderer = TemplateRenderer()
    rendered = renderer.render_for_row(
        {"Manual Decision": "Review", "Overseas": "No", "Full Name of the Last Author": "A B"},
        date(2026, 5, 16),
    )
    assert rendered.template_path.name == "NB_Template_Review_No.docx"
    assert "Aaaaa" in rendered.placeholders
    assert rendered.body_text == ""
