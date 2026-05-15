# Next steps for GitHub + Codex cloud

1. Create a private GitHub repository, e.g. `nb-review-invitation-agent`.
2. Commit the current v14 script, `.bat`, README, this AGENTS.md, `.gitignore`, `.env.example`, and the task files.
3. Do not commit real secrets. Prefer sanitized workbook fixtures for CI.
4. Connect the repository to Codex web.
5. Create a new Codex task using `CODEX_TASK_01_REPO_BASELINE.md`.
6. Review the PR. Ask Codex for code review with `@codex review` if GitHub review integration is enabled.
7. Merge only after tests and the verification note are acceptable.
8. Repeat tasks 02–06 in order.

Important: Codex cloud can implement and test most Python logic, but final Outlook/Word/Excel COM behavior must be smoke-tested on a Windows machine with Office/Outlook installed.
