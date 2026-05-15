# Codex Task 06 — Final test, packaging, and release artifact

## Goal
Produce a stable downloadable version.

## Required changes
1. Add complete README:
   - Installation
   - Environment variables
   - How to run dry-run
   - How to run real enrichment
   - How to run GUI only
   - How to send current author invitation
   - How to use BatchInvitation safely
   - Troubleshooting locked Excel files and Outlook account issues
2. Add `.bat` launchers:
   - `run_import_and_gui.bat`
   - `run_gui_only.bat`
   - `run_dry_run_tests.bat`
3. Add packaging script:
   - zip source + README + `.bat` + `.env.example`
   - Do not include `.env`, private workbooks, RunLogs, PubMed XML, or SQLite state.
4. Add test report template.
5. Add release checklist.

## Acceptance tests
- `python -m pytest` passes.
- Dry-run command runs without network/Outlook.
- Package zip excludes secrets and runtime data.
- README explains that real Outlook send must be tested on Windows with Office installed.

## Final release checklist
- [ ] Cloud tests pass.
- [ ] Windows smoke test: GUI opens.
- [ ] Windows smoke test: template preview renders.
- [ ] Windows smoke test: fake/dry-run invitation does not send.
- [ ] Windows smoke test: real Outlook draft appears for one test row.
- [ ] Verified no invitation date is written before send.
- [ ] Verified invitation date is written after confirmed send.
