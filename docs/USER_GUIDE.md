# User Guide

## Required local files location
Put runtime workbook/log files in `REVIEW_INVITATION_BASE_DIR` (outside git):
- `NB_Author_2026.xlsm`
- `ReviewInvitationLog.xlsx`

For this version, keep the three template files in the release `templates/` folder:
- `templates/NB_Template_Review_Yes.docx`
- `templates/NB_Template_Review_No.docx`
- `templates/NB_Template_Insight.docx`

Custom template directory support can be added later if needed.

## Open GUI
- CLI: `python -m nb_review_invitation_agent.cli --gui`
- Windows: `run_gui_windows.bat`

## Basic review flow
1. Open workbook in GUI.
2. Select batch from dropdown.
3. Edit fields `Overseas`, `Manual Decision`, and `Research`.
4. Save updates.
5. Use link buttons to open web/research references.

## Send one invitation
1. Select a single row.
2. Trigger current-row invitation.
3. Confirm in dialog.
4. On successful send, `Date of Invitaion` is written.

If cancelled/failed/skipped, date must remain unchanged.

## BatchInvitation
- Use batch action only after review.
- Confirmation is required per row.
- Already invited rows are skipped.
- `Manual Decision = No` rows are skipped.

## Status meanings
- Pending: ready for review/invitation.
- Skipped: row excluded by rule or user decision.
- Cancelled: user cancelled at confirmation.
- Sent: Outlook send succeeded and date recorded.
- Failed: send or render failed; no date update.

## Outlook account not found (`nsb@ion.ac.cn`)
1. Confirm Outlook profile contains the account.
2. Re-open Outlook and retry.
3. Check `.env` `OUTLOOK_SEND_ACCOUNT` value.
4. Use dry-run/fake-provider path until Outlook is fixed.
