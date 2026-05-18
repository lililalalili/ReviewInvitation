# NB Review Invitation Agent

NB Review Invitation Agent supports Neuroscience Bulletin invitation operations with a safe workflow:
- import/enrich reviewer rows,
- review rows in GUI,
- create/send Outlook invitations only after explicit human confirmation,
- write `Date of Invitaion` only after successful send.

## Installation

### Cross-platform dev/CI
```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
python -m pytest -q
```

### Windows runtime setup
```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev,windows]"
python -m pytest -q
python -m nb_review_invitation_agent.cli --gui
```

`pywin32` is optional and only installed through `.[windows]`.

## Local operational files (do not commit)
Set `REVIEW_INVITATION_BASE_DIR` to a local directory outside git. Put runtime data files there:
- `NB_Author_2026.xlsm`
- `ReviewInvitationLog.xlsx`

For this release, keep template files in the repository/release `templates/` folder:
- `templates/NB_Template_Review_Yes.docx`
- `templates/NB_Template_Review_No.docx`
- `templates/NB_Template_Insight.docx`

Custom template directory support can be added later if needed.

Never commit real workbooks, logs, `.env`, PubMed XML, sqlite files, RunLogs, or private author data.

## CLI
```bash
python -m nb_review_invitation_agent.cli --help
```

Important options:
- `--gui`: launch GUI.
- `--dry-run`: no-op/validation mode.
- `--fake-providers`: offline fake provider flow for CI/tests.
- `--legacy-v14`: run legacy monolithic path.

## Windows launch scripts
- `run_gui_windows.bat`
- `run_dry_run_windows.bat`
- `run_tests_windows.bat`
- `run_import_and_gui.bat`
- `run_gui_only.bat`
- `run_dry_run_tests.bat`

## Dry-run and test commands
```bash
python -m nb_review_invitation_agent.cli --dry-run --fake-providers --no-gui
python -m pytest -q
```

These commands are designed to run without Outlook/Word COM, GUI display, network, or API keys.

## DeepSeek/Search configuration (later/optional)
Configure in `.env` (see `.env.example`):
- `DEEPSEEK_API_KEY`, `DEEPSEEK_BASE_URL`, `DEEPSEEK_MODEL`
- `SEARCH_PROVIDER`, `SEARCH_API_KEY`
- `BRAVE_SEARCH_API_KEY` (reserved for planned Brave Search provider)

## Outlook safety notes
- No email should be sent without explicit confirmation.
- Cancelled/failed/skipped sends must not write `Date of Invitaion`.
- Keep `.xlsm` macro preservation enabled (`keep_vba=True`).
- Validate real sending only on Windows + Office/Outlook installed.

## Packaging and release
- Use `python scripts/make_release_zip.py` to build a release zip.
- The zip intentionally excludes sensitive/runtime files (`.env`, Excel runtime data, RunLogs, PubMed XML, sqlite, caches).
- See docs in:
  - `docs/USER_GUIDE.md`
  - `docs/DEVELOPER_GUIDE.md`
  - `docs/WINDOWS_SMOKE_TEST_CHECKLIST.md`
  - `docs/RELEASE_CHECKLIST.md`
