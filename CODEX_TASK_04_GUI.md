# Codex Task 04 — Tkinter GUI review workflow

## Goal
Add GUI that opens after information collection and supports review/edit/navigation.

## Required changes
1. Build Tkinter GUI in `gui_app.py`.
2. On launch, load workbook and select latest `Batch ID`; display first row in that batch.
3. Batch dropdown lists all batch IDs and switches displayed rows.
4. Display required fields from project requirements.
5. Editable controls:
   - `Overseas`: Yes/No dropdown
   - `Manual Decision`: Review/Insight/No dropdown
   - `Research field`: text area
6. Link buttons:
   - Open `Last Author Web`
   - Open `Pubmed Link`
   - Open Bing search for `Email of the Last Author`
7. Navigation buttons:
   - Previous
   - Next
   - First in batch
   - Last in batch
   - Save all modifications
   - Invite current author
   - BatchInvitation
8. Ensure unsaved edits are not lost on navigation.

## Acceptance tests
- Non-GUI model/controller tests verify batch selection and row navigation.
- Save writes editable values back to fixture workbook.
- Link builder returns correct URLs.
- GUI import does not require Outlook.

## Manual smoke test
- Run app on Windows.
- Confirm it opens latest batch first row.
- Edit Overseas/Manual Decision/Research field and save.
