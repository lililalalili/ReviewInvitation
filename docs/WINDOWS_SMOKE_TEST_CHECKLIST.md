# Windows Smoke Test Checklist

1. Clone/download repo.
2. Install Python 3.11.
3. Create venv.
4. Install package with windows extra.
5. Run pytest.
6. Copy real local files into `REVIEW_INVITATION_BASE_DIR`.
7. Run GUI.
8. Verify workbook opens.
9. Verify batch dropdown works.
10. Verify editable fields save (`Overseas`, `Manual Decision`, `Research field`). Ensure `NB_Author_2026.xlsm` is closed in Excel before Save or sending invitations.
11. Verify link buttons open references.
12. Verify current-row invitation preflights workbook save before opening any Outlook draft, renders body from clean UTF-8 HTML templates under `templates/`, assigns `Outlook.MailItem.HTMLBody`, then creates draft/send flow.
13. Verify cancel in confirmation dialog does not send and does not write date.
14. Verify confirmed successful send writes `Date of Invitaion` only after send success.
15. If placeholders remain, do **not** send and report the remaining placeholders. A draft should not be displayed if template rendering fails.
16. Capture and report exactly which placeholders remain (for example: `Aaaaa`, `Ttttt`).
17. Intended production path: clean HTML template rendering (`templates/*.html`) with placeholder replacement and validation before `mail.Display()`.
18. `.docx` templates are retained in the repository only as fallback/archive; they are not the normal production rendering path.
19. If email body looks malformed, suspiciously short, or any known placeholder remains, do **not** send.
20. Confirm sender is `nsb@ion.ac.cn` and recipient is the test email before sending.
21. Verify rendering failures block before any Outlook draft is displayed.
22. Verify already-invited row is skipped.
23. Verify `Manual Decision = No` row is skipped.
24. Verify BatchInvitation requests confirmation per row.

Recommended commands:
```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev,windows]"
python -m pytest -q
python -m nb_review_invitation_agent.cli --gui
```

25. If the displayed email body looks squeezed together, malformed, or spacing/punctuation is corrupted, do **not** send; capture and report a screenshot.
26. If placeholders remain anywhere in rendered HTML/body, do **not** send.
