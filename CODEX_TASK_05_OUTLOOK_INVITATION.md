# Codex Task 05 — Word template rendering + Outlook invitation send

## Goal
Implement current-author invitation and batch invitation through Outlook with manual confirmation.

## Required changes
1. Add `template_renderer.py`.
   - Preserve Word formatting for `.docx` templates.
   - Replace placeholders:
     `Aaaaa`, `Jjjjj`, `Ttttt`, `Fffff`, `Pppppyes`, `Pppppno`, `Dddddre`, `Dddddin`.
2. Add `mailer_outlook.py`.
   - Create Outlook MailItem.
   - Set To, optional CC, subject, and body.
   - Set sending account to `nsb@ion.ac.cn` via config.
   - Display email and require confirmation before send.
   - Support fake mailer in tests.
3. Implement workflow guards:
   - `Manual Decision = No` stops and prompts confirmation.
   - Non-empty `Date of Invitaion` stops with `已邀请过`.
   - Missing recipient email stops with clear error.
   - Send failure does not write invitation date.
4. After successful send only, write current date to `Date of Invitaion` and update status.
5. Implement BatchInvitation over current batch only.

## Template mapping
- Review + Overseas Yes: `NB_Template_Review_Yes.docx`; subject `Neuroscience Bulletin Invites You to Submit a Review`.
- Review + Overseas No: `NB_Template_Review_No.docx`; subject `Neuroscience Bulletin Invites You to Submit a Review`.
- Insight: `NB_Template_Insight.docx`; subject `Neuroscience Bulletin Invites You to Submit an Insight`.

## Acceptance tests
- Template selection matrix is correct.
- Date calculation works:
  - Review: today + 6 months
  - Insight: today + 125 days
- Title trailing period is removed.
- Fake mailer sends only after confirmation.
- Invitation date is written only on successful fake send.
- Batch skips already invited rows.
