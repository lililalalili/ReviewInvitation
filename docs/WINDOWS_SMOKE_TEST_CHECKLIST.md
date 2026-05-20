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
12. Verify current-row invitation creates formatted Outlook draft/send flow.
13. Verify cancel in confirmation dialog does not send and does not write date.
14. Verify confirmed successful send writes `Date of Invitaion` only after send success.
15. If placeholders remain, do **not** send and report the remaining placeholders. A draft should not be displayed if template rendering fails.
16. Capture and report exactly which placeholders remain (for example: `Aaaaa`, `Ttttt`).
17. Intended production path: Word-first replacement with OOXML fallback if placeholders remain.
18. Verify already-invited row is skipped.
19. Verify `Manual Decision = No` row is skipped.
20. Verify BatchInvitation requests confirmation per row.

Recommended commands:
```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev,windows]"
python -m pytest -q
python -m nb_review_invitation_agent.cli --gui
```
