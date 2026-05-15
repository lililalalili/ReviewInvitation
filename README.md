# NB Review Invitation Agent

Baseline repository refactor that introduces a Python package scaffold while preserving rollback to legacy v14.

## Run tests

```bash
python -m pytest
```

## CLI help

```bash
python -m nb_review_invitation_agent.cli --help
```

## Behavior notes

- `--legacy-v14` runs the existing monolithic `legacy/review_invitation_agent_windows_ubuntu_ollama_v14.py` entry path.
- `--dry-run`, `--no-gui`, and `--fake-providers` are available for CI-safe bootstrap.
- GUI/enrichment/Outlook implementation are intentionally deferred to later tasks.
