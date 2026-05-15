from __future__ import annotations

from datetime import date
from urllib.parse import quote_plus

JOURNALS = ["Nature", "Science", "Cell", "Nature Neuroscience", "Neuron"]


def fmt_slash(d: date) -> str:
    return d.strftime("%Y/%m/%d")


def build_pubmed_query(start_date: date, end_date: date) -> str:
    journal_block = " OR ".join([f'\"{j}\"[Journal]' for j in JOURNALS])
    date_block = f'("{fmt_slash(start_date)}"[Date - Publication] : "{fmt_slash(end_date)}"[Date - Publication])'
    exclude = "NOT (erratum[pt] OR editorial[pt] OR comment[pt] OR review[pt])"
    return f"({journal_block}) AND {date_block} AND {exclude}"


def build_pubmed_search_url(query: str) -> str:
    return f"https://pubmed.ncbi.nlm.nih.gov/?term={quote_plus(query)}"
