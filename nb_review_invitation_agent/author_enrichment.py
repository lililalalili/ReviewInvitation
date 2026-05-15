from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field

from .search_provider import SearchProvider, SearchResult


SANITIZED_PROMPT_FIELDS = [
    "Full Name of the Last Author",
    "Email of the Last Author",
    "Affiliation of the Last Author",
    "First Author Full name",
    "Affiliation of the First Author",
    "First Author Email",
]


def clean_space(text: str | None) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def extract_email(text: str) -> str:
    m = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text or "")
    return m.group(0) if m else ""


def mask_email(email: str) -> str:
    if not email or "@" not in email:
        return ""
    local, domain = email.split("@", 1)
    if not local:
        return f"***@{domain}"
    return f"{local[0]}***@{domain}"


def _coerce_evidence(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [clean_space(str(item)) for item in value if clean_space(str(item))]
    text = clean_space(str(value))
    return [text] if text else []


def _coerce_confidence(value) -> float:
    if value is None:
        return 0.0
    text = clean_space(str(value))
    if not text:
        return 0.0
    try:
        if text.endswith('%'):
            return float(text[:-1].strip()) / 100.0
        return float(text)
    except (TypeError, ValueError):
        return 0.0


@dataclass(frozen=True)
class EnrichmentResult:
    personal_web_url: str = ""
    research_quote: str = ""
    last_author_email: str = ""
    first_author_email: str = ""
    confidence: float = 0.0
    evidence: list[str] = field(default_factory=list)
    status: str = "Needs Review"
    error: str = ""


class AuthorEnricher:
    def __init__(self, *, search_provider: SearchProvider, llm_client, confidence_threshold: float = 0.75):
        self._search_provider = search_provider
        self._llm_client = llm_client
        self._confidence_threshold = confidence_threshold

    def enrich_row(self, row: dict[str, str]) -> dict[str, str]:
        updated = dict(row)
        try:
            result = self._run(row)
        except Exception as exc:
            updated["Author Enrichment Status"] = "Error"
            updated["Author Enrichment Evidence"] = f"provider_error: {clean_space(str(exc))}"
            return updated

        updated["Author Enrichment Evidence"] = json.dumps(asdict(result), ensure_ascii=False)
        if result.confidence < self._confidence_threshold:
            updated["Author Enrichment Status"] = "Needs Review"
            return updated

        if result.personal_web_url:
            updated["Last Author Web"] = result.personal_web_url
        if result.research_quote:
            updated["Last Author Research"] = result.research_quote

        if not clean_space(updated.get("Email of the Last Author")) and result.last_author_email:
            updated["Email of the Last Author"] = result.last_author_email
        if not clean_space(updated.get("First Author Email")) and result.first_author_email:
            updated["First Author Email"] = result.first_author_email

        updated["Author Enrichment Status"] = "Enriched"
        return updated

    def _run(self, row: dict[str, str]) -> EnrichmentResult:
        queries = self._build_queries(row)
        candidates: list[SearchResult] = []
        for q in queries:
            candidates.extend(self._search_provider.search(q))
        prompt = self._build_user_prompt(row=row, candidates=candidates)
        llm_json = self._llm_client.generate_json(system_prompt=self._system_prompt(), user_prompt=prompt)
        return EnrichmentResult(
            personal_web_url=clean_space(llm_json.get("personal_web_url")),
            research_quote=clean_space(llm_json.get("research_quote")),
            last_author_email=extract_email(clean_space(llm_json.get("last_author_email"))),
            first_author_email=extract_email(clean_space(llm_json.get("first_author_email"))),
            confidence=_coerce_confidence(llm_json.get("confidence")),
            evidence=_coerce_evidence(llm_json.get("evidence")),
            status=clean_space(llm_json.get("status")) or "Needs Review",
            error=clean_space(llm_json.get("error")),
        )

    def _build_queries(self, row: dict[str, str]) -> list[str]:
        queries: list[str] = []

        last_name = clean_space(row.get("Full Name of the Last Author"))
        last_email = clean_space(row.get("Email of the Last Author"))
        last_aff = clean_space(row.get("Affiliation of the Last Author"))

        first_name = clean_space(row.get("First Author Full name"))
        first_aff = clean_space(row.get("Affiliation of the First Author"))
        first_email = clean_space(row.get("First Author Email"))

        if last_name and last_aff:
            queries.append(f"{last_name} {last_aff}")
        elif last_name:
            queries.append(f'"{last_name}" neuroscience')

        if last_email:
            queries.append(last_email)

        if not first_email:
            if first_name and first_aff:
                queries.append(f"{first_name} {first_aff}")
            elif first_name:
                queries.append(f'"{first_name}" neuroscience')

        deduped: list[str] = []
        seen: set[str] = set()
        for q in queries:
            cq = clean_space(q)
            if not cq or cq in seen:
                continue
            seen.add(cq)
            deduped.append(cq)
        return deduped

    def _build_user_prompt(self, *, row: dict[str, str], candidates: list[SearchResult]) -> str:
        candidates_text = "\n".join(
            f"- title={c.title}; url={c.url}; snippet={clean_space(c.snippet)}" for c in candidates[:10]
        )
        sanitized_row = {k: clean_space(row.get(k, "")) for k in SANITIZED_PROMPT_FIELDS}
        return (
            f"Author row: {json.dumps(sanitized_row, ensure_ascii=False)}\n"
            f"Candidates:\n{candidates_text}\n"
            "Return JSON with keys: personal_web_url, research_quote, last_author_email, first_author_email, confidence, evidence, status, error."
        )

    @staticmethod
    def _system_prompt() -> str:
        return (
            "Identify author webpage/research/email from candidates. Prefer official university/lab pages. "
            "If uncertain, keep confidence low and explain evidence. Return valid JSON only."
        )
