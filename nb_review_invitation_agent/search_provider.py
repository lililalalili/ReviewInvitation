from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol


@dataclass(frozen=True)
class SearchResult:
    title: str
    url: str
    snippet: str
    text: str | None = None


class SearchProvider(Protocol):
    def search(self, query: str) -> list[SearchResult]:
        ...


class FakeSearchProvider:
    def __init__(self, results_by_query: dict[str, list[SearchResult]] | None = None):
        self._results_by_query = results_by_query or {}

    def search(self, query: str) -> list[SearchResult]:
        return list(self._results_by_query.get(query, []))
