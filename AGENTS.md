# AGENTS.md — NB Review Invitation Agent

## Project goal
Build a stable Windows desktop workflow for Neuroscience Bulletin invitation management.
The app reads PubMed-derived author rows, enriches author web/email/research information using external search + DeepSeek, writes results into `NB_Author_2026.xlsm`, opens a GUI for review, and sends invitation emails through Outlook only after explicit human confirmation.

## Non-negotiable requirements
- Preserve `.xlsm` macros and workbook formatting when editing `NB_Author_2026.xlsm`.
- Never write `Date of Invitaion` during PubMed/enrichment import. Write it only after an invitation email is actually sent.
- Never auto-send email without a visible user confirmation dialog.
- Never commit API keys, Outlook credentials, real run logs, or private author data.
- Keep a dry-run / fake-provider mode so CI can pass without network, Outlook, Word, or Excel desktop automation.
- Prefer small, reviewable PRs. Each PR must include tests and a short verification note.

## Runtime assumptions
- Production runtime is Windows with Microsoft Office / Outlook installed.
- Cloud CI/Codex may run on Linux; therefore Outlook COM, Word COM, and real GUI behavior must be mocked or tested with non-COM fallback logic in CI.
- Real API calls should be disabled by default in tests.

## Coding style
- Python 3.11+.
- Favor modular files over a single monolithic script.
- Use dataclasses or pydantic-style typed structures where helpful.
- Use explicit column lookup by header name, not hard-coded column numbers, except in migration tests.
- Provide clear error messages for locked Excel files, missing templates, missing Outlook account, and missing API keys.

## Safety and privacy
- Do not log full author email lists or secret values.
- If logging email addresses is unavoidable for debugging, log only the row number and masked email, e.g. `a***@example.edu`.
- Real workbook files should live in a runtime directory that is gitignored unless the repository owner intentionally adds a sanitized fixture.

## Review guidelines
- Flag any change that can write `Date of Invitaion` before successful email send as P1.
- Flag any path that sends Outlook email without human confirmation as P1.
- Flag any hard-coded secret, API key, or private author data as P1.
- Flag tests that depend on live DeepSeek/search/Outlook as P1 unless explicitly marked integration and skipped by default.
- Flag changes that break `.xlsm` macro preservation (`keep_vba=True`) as P1.
