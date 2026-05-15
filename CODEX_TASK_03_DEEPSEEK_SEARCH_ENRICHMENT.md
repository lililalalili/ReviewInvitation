# Codex Task 03 — DeepSeek + search author enrichment

## Goal
Add author enrichment to find last-author web page, research direction quote, and missing emails.

## Required changes
1. Add `DeepSeekClient` compatible with OpenAI-style chat completions.
2. Add configurable search provider interface:
   - `SearchProvider.search(query) -> list[SearchResult]`
   - Include a fake provider for tests.
3. Add page fetcher/extractor:
   - Respect timeout/retry.
   - Extract title, URL, visible text snippet.
   - Skip binary/PDF unless explicitly supported.
4. Add `AuthorEnricher`:
   - Input: row data with last author name/email/affiliation and first author name/affiliation/email.
   - Output structured JSON with confidence/evidence.
5. Write high-confidence results to workbook:
   - `Last Author Web`
   - `Last Author Research`
   - `Email of the Last Author` if missing
   - `First Author Email` if missing
   - status/evidence columns
6. Low confidence results should set `Author Enrichment Status = Needs Review`.
7. Add audit logs without exposing API keys.

## Prompt requirements for DeepSeek
Ask DeepSeek to:
- Determine whether the candidate page belongs to the same author.
- Prefer official university/lab/institution pages over generic profiles.
- Extract the main research direction as a quote when available.
- Return valid JSON only.

## Acceptance tests
- Fake search + fake DeepSeek fills fields deterministically.
- Low confidence does not overwrite existing email.
- Existing non-empty emails are not overwritten unless an explicit setting allows it.
- Missing API key gives a clear error unless fake/dry-run mode is active.
