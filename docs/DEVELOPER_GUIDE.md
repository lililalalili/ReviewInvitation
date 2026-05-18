# Developer Guide

## Package structure
- `nb_review_invitation_agent/`: main package.
- `legacy/`: v14 fallback implementation.
- `tests/`: offline unit tests.
- `scripts/`: release helper scripts.
- `docs/`: operational/release documentation.

## Task history
- Tasks 01-05 established package scaffold, Excel safety behavior, enrichment flow, GUI baseline, and Outlook invitation safety invariants.
- Task 06 focuses on test/release hardening and documentation, without risky production behavior changes.

## Safety invariants (must not regress)
- Never send without explicit confirmation.
- Never write `Date of Invitaion` before successful send.
- Cancelled/failed/skipped sends must not write invitation date.
- Preserve workbook macros (`keep_vba=True`) for `.xlsm`.

## Testing strategy
- CI/Linux: offline tests only (`pytest -q`).
- No dependency on Outlook/Word COM, GUI display, network, DeepSeek/search keys.
- Use `--dry-run --fake-providers --no-gui` to verify safe bootstrap behavior.

## Extend search provider
- Add a new `SearchProvider` implementation in package code.
- Keep fake/offline provider path for tests.
- Add config/env wiring and unit tests without live API calls.

### Brave Search TODO (planning only)
- Add optional Brave Search API implementation later via provider interface.
- Use `BRAVE_SEARCH_API_KEY` env var.
- Do not automate Brave browser app.
- Keep offline tests with fake provider.

## Debugging Outlook/Word COM on Windows
1. Install `.[windows]` extras.
2. Ensure Outlook is installed and account configured.
3. Run a single-row invitation with explicit confirmation.
4. Inspect account selection/template resolution errors.
5. Reproduce with dry-run to isolate non-COM logic.
