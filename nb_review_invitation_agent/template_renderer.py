from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import calendar
import re
from pathlib import Path


REVIEW_SUBJECT = "Neuroscience Bulletin Invites You to Submit a Review"
INSIGHT_SUBJECT = "Neuroscience Bulletin Invites You to Submit an Insight"


@dataclass(frozen=True)
class TemplateChoice:
    template_name: str
    subject: str


@dataclass(frozen=True)
class RenderedInvitation:
    template_path: Path
    template_name: str
    subject: str
    placeholders: dict[str, str]
    body_text: str = ""


def add_calendar_months(value: date, months: int) -> date:
    month_index = value.month - 1 + months
    year = value.year + month_index // 12
    month = month_index % 12 + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def format_readable_date(value: date) -> str:
    return value.strftime("%B %d, %Y")


def normalize_journal_name(journal: str) -> str:
    parts = re.split(r"\s+", (journal or "").strip())
    return " ".join(p.capitalize() for p in parts if p)


def trim_terminal_period(title: str) -> str:
    text = (title or "").strip()
    return text[:-1] if text.endswith(".") else text


def derive_family_name(row_values: dict[str, str]) -> str:
    family = row_values.get("Family Name of the Last Author", "").strip()
    if family:
        return family
    full_name = row_values.get("Full Name of the Last Author", "").strip()
    if not full_name:
        return ""
    parts = full_name.split()
    return parts[-1] if parts else full_name


def select_template_and_subject(row_values: dict[str, str]) -> TemplateChoice:
    manual = row_values.get("Manual Decision", "").strip()
    overseas = row_values.get("Overseas", "").strip()

    if manual == "Review":
        if overseas == "Yes":
            return TemplateChoice("NB_Template_Review_Yes.docx", REVIEW_SUBJECT)
        return TemplateChoice("NB_Template_Review_No.docx", REVIEW_SUBJECT)
    if manual == "Insight":
        return TemplateChoice("NB_Template_Insight.docx", INSIGHT_SUBJECT)
    if manual == "No":
        raise ValueError("Manual Decision = No，请确认是否邀请")
    raise ValueError("Manual Decision 为空或不支持，请先设置为 Review / Insight")


def build_placeholder_map(row_values: dict[str, str], today: date) -> dict[str, str]:
    overseas = row_values.get("Overseas", "").strip()
    manual = row_values.get("Manual Decision", "").strip()
    p_yes = ""
    p_no = ""
    if manual == "Review":
        if overseas == "Yes":
            p_yes = "For invited overseas authors, article publication charges will be covered by the journal, and NB will pay 1,000 USD remuneration for an accepted Review article."
        else:
            p_no = "For invited authors, article publication charges will be covered by the journal."
    elif manual == "Insight":
        p_no = "For invited authors, article publication charges will be covered by the journal."
    return {
        "Aaaaa": derive_family_name(row_values),
        "Jjjjj": normalize_journal_name(row_values.get("Journal", "")),
        "Ttttt": trim_terminal_period(row_values.get("Title", "")),
        "Fffff": row_values.get("Research field", "").strip(),
        "Pppppyes": p_yes,
        "Pppppno": p_no,
        "Dddddre": format_readable_date(add_calendar_months(today, 6)),
        "Dddddin": format_readable_date(today.fromordinal(today.toordinal() + 125)),
    }


class TemplateRenderer:
    def __init__(self, templates_dir: str | Path = "templates") -> None:
        self.templates_dir = Path(templates_dir)

    def render_for_row(self, row_values: dict[str, str], today: date) -> RenderedInvitation:
        choice = select_template_and_subject(row_values)
        placeholders = build_placeholder_map(row_values, today)
        template_path = self.templates_dir / choice.template_name
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        return RenderedInvitation(
            template_path=template_path,
            template_name=choice.template_name,
            subject=choice.subject,
            placeholders=placeholders,
        )
