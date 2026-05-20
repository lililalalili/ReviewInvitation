from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import calendar
import re
from html import escape
from pathlib import Path


REVIEW_SUBJECT = "Neuroscience Bulletin Invites You to Submit a Review"
INSIGHT_SUBJECT = "Neuroscience Bulletin Invites You to Submit an Insight"

JOURNAL_ACRONYMS = {"PNAS", "CNS", "SNYNY", "JAMA", "BMJ", "EMBO", "FASEB", "PLOS"}

JOURNAL_NAME_EXCEPTIONS: dict[str, str] = {
    "nature": "Nature",
    "science": "Science",
    "cell": "Cell",
    "neuron": "Neuron",
    "nature neuroscience": "Nature Neuroscience",
    "nature methods": "Nature Methods",
    "nature medicine": "Nature Medicine",
    "nature biotechnology": "Nature Biotechnology",
    "nature genetics": "Nature Genetics",
    "nature aging": "Nature Aging",
    "nature communications": "Nature Communications",
    "communications biology": "Communications Biology",
    "communications medicine": "Communications Medicine",
    "science advances": "Science Advances",
    "science translational medicine": "Science Translational Medicine",
    "cell reports": "Cell Reports",
    "cell reports medicine": "Cell Reports Medicine",
    "current biology": "Current Biology",
    "developmental cell": "Developmental Cell",
    "immunity": "Immunity",
    "cancer cell": "Cancer Cell",
    "molecular cell": "Molecular Cell",
    "nature reviews neuroscience": "Nature Reviews Neuroscience",
    "nature reviews neurology": "Nature Reviews Neurology",
    "trends in neurosciences": "Trends in Neurosciences",
    "trends in cognitive sciences": "Trends in Cognitive Sciences",
    "annual review of neuroscience": "Annual Review of Neuroscience",
    "physiological reviews": "Physiological Reviews",
    "journal of neuroscience": "Journal of Neuroscience",
    "the journal of neuroscience": "The Journal of Neuroscience",
    "journal of neurophysiology": "Journal of Neurophysiology",
    "neuroscience bulletin": "Neuroscience Bulletin",
    "brain": "Brain",
    "brain research": "Brain Research",
    "cerebral cortex": "Cerebral Cortex",
    "neuroimage": "NeuroImage",
    "annals of neurology": "Annals of Neurology",
    "jama neurology": "JAMA Neurology",
    "lancet neurology": "Lancet Neurology",
    "the lancet neurology": "The Lancet Neurology",
    "molecular psychiatry": "Molecular Psychiatry",
    "biological psychiatry": "Biological Psychiatry",
    "translational psychiatry": "Translational Psychiatry",
    "nature mental health": "Nature Mental Health",
    "elife": "eLife",
    "plos biology": "PLOS Biology",
    "plos one": "PLOS ONE",
    "plos computational biology": "PLOS Computational Biology",
    "proceedings of the national academy of sciences": "Proceedings of the National Academy of Sciences",
    "proc natl acad sci u s a": "Proc Natl Acad Sci U S A",
    "pnas": "PNAS",
    "pnas nexus": "PNAS Nexus",
    "science (new york, n.y.)": "Science (New York, N.Y.)",
    "science new york n y": "Science (New York, N.Y.)",
    "snyny": "SNYNY",
    "embo journal": "EMBO Journal",
    "the embo journal": "The EMBO Journal",
    "embo reports": "EMBO Reports",
    "f1000research": "F1000Research",
    "biorxiv": "bioRxiv",
    "medrxiv": "medRxiv",
    "cell stem cell": "Cell Stem Cell",
    "nature cell biology": "Nature Cell Biology",
    "nature structural & molecular biology": "Nature Structural & Molecular Biology",
    "nature biomedical engineering": "Nature Biomedical Engineering",
    "advanced science": "Advanced Science",
    "advanced materials": "Advanced Materials",
    "acs chemical neuroscience": "ACS Chemical Neuroscience",
    "journal of clinical investigation": "Journal of Clinical Investigation",
    "new england journal of medicine": "New England Journal of Medicine",
    "the new england journal of medicine": "The New England Journal of Medicine",
    "lancet": "The Lancet",
    "the lancet": "The Lancet",
    "bmj": "BMJ",
    "jama": "JAMA",
}



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
    rendered_html: str


def add_calendar_months(value: date, months: int) -> date:
    month_index = value.month - 1 + months
    year = value.year + month_index // 12
    month = month_index % 12 + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def format_readable_date(value: date) -> str:
    return value.strftime("%B %d, %Y")


def normalize_journal_name(journal: str) -> str:
    normalized_input = (journal or "").strip()
    if not normalized_input:
        return ""

    canonical = JOURNAL_NAME_EXCEPTIONS.get(normalized_input.lower())
    if canonical:
        return canonical

    def normalize_word(word: str) -> str:
        if re.fullmatch(r"(?:[A-Za-z]\.){2,}", word):
            return word.upper()

        if word.isalpha() and word.upper() in JOURNAL_ACRONYMS:
            return word.upper()

        if word.isalpha():
            return word[0].upper() + word[1:].lower()

        return word

    def normalize_segment(segment: str) -> str:
        return re.sub(r"[A-Za-z.]+", lambda m: normalize_word(m.group(0)), segment)

    parts = re.split(r"(\s+)", normalized_input)
    normalized = "".join(normalize_segment(part) if not part.isspace() else part for part in parts)
    return JOURNAL_NAME_EXCEPTIONS.get(normalized.lower(), normalized)


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
            return TemplateChoice("NB_Template_Review_Yes.html", REVIEW_SUBJECT)
        return TemplateChoice("NB_Template_Review_No.html", REVIEW_SUBJECT)
    if manual == "Insight":
        return TemplateChoice("NB_Template_Insight.html", INSIGHT_SUBJECT)
    if manual == "No":
        raise ValueError("Manual Decision = No，请确认是否邀请")
    raise ValueError("Manual Decision 为空或不支持，请先设置为 Review / Insight")


def build_placeholder_map(row_values: dict[str, str], today: date) -> dict[str, str]:
    manual = row_values.get("Manual Decision", "").strip()
    overseas = row_values.get("Overseas", "").strip()
    standard_apc = "For invited authors, article publication charges will be covered by the journal."
    return {
        "Aaaaa": derive_family_name(row_values),
        "Jjjjj": normalize_journal_name(row_values.get("Journal", "")),
        "Ttttt": trim_terminal_period(row_values.get("Title", "")),
        "Fffff": row_values.get("Research field", "").strip(),
        "Pppppyes": "For invited overseas authors, article publication charges will be covered by the journal, and NB will pay 1,000 USD remuneration for an accepted Review article."
        if manual == "Review" and overseas == "Yes"
        else "",
        "Pppppno": standard_apc if manual == "Insight" or (manual == "Review" and overseas != "Yes") else "",
        "Dddddre": format_readable_date(add_calendar_months(today, 6)),
        "Dddddin": format_readable_date(today.fromordinal(today.toordinal() + 125)),
    }


class TemplateRenderer:
    def __init__(self, templates_dir: str | Path = "templates") -> None:
        self.templates_dir = Path(templates_dir)

    def render_for_row(self, row_values: dict[str, str], today: date) -> RenderedInvitation:
        choice = select_template_and_subject(row_values)
        placeholders = build_placeholder_map(row_values, today)
        template_path = (self.templates_dir / choice.template_name).resolve()
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        rendered_html = self.render_html_for_row(template_path, placeholders)
        return RenderedInvitation(
            template_path=template_path,
            template_name=choice.template_name,
            subject=choice.subject,
            placeholders=placeholders,
            rendered_html=rendered_html,
        )

    def render_html_for_row(self, template_path: Path, placeholders: dict[str, str]) -> str:
        template_html = template_path.read_text(encoding="utf-8")
        rendered_html = self.replace_placeholders_in_html(template_html, placeholders)
        self.validate_rendered_html(rendered_html)
        return rendered_html

    def replace_placeholders_in_html(self, template_html: str, placeholders: dict[str, str]) -> str:
        rendered = template_html
        for key, value in placeholders.items():
            rendered = rendered.replace(key, escape(str(value or ""), quote=True))
        return rendered

    def validate_rendered_html(self, rendered_html: str) -> None:
        if len(rendered_html.strip()) <= 20:
            raise RuntimeError("Rendered email body is empty.")
        remaining = [k for k in build_placeholder_map({}, date.today()).keys() if k in rendered_html]
        if remaining:
            raise RuntimeError(f"Rendered email body still contains placeholders: {', '.join(remaining)}")
