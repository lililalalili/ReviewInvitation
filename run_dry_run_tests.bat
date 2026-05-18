@echo off
setlocal
python -m nb_review_invitation_agent.cli --dry-run --fake-providers --no-gui
python -m pytest -q %*
endlocal
