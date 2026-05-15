# Codex Task 01 — Create repo baseline and test harness

## Goal
Turn the current single-file v14 script into a maintainable Python project without changing production behavior yet.

## Inputs
- Existing script: `review_invitation_agent_windows_ubuntu_ollama_v14.py`
- Existing `.bat`
- Existing README
- Existing templates and workbook should not be committed unless sanitized.

## Required changes
1. Create package structure:
   ```text
   nb_review_invitation_agent/
     __init__.py
     config.py
     workbook.py
     pubmed_pipeline.py
     llm_clients.py
     author_enrichment.py
     gui_app.py
     mailer_outlook.py
     template_renderer.py
     cli.py
   tests/
   ```
2. Move code gradually from v14 into modules.
3. Add `pyproject.toml` or `requirements.txt`.
4. Add `pytest` tests that do not need network or Outlook.
5. Add `--dry-run`, `--no-gui`, and `--fake-providers` CLI flags.
6. Preserve a compatibility entry point so existing `.bat` still works or is updated clearly.

## Acceptance tests
- `python -m pytest` passes.
- `python -m nb_review_invitation_agent.cli --help` works.
- No production API call is made during tests.
- Existing v14 behavior is preserved at high level.

## Verification note to include in PR
- List files moved.
- State whether any behavior changed. For this task, behavior should be unchanged except CLI/bootstrap.
