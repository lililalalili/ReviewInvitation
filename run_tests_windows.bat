@echo off
setlocal
python -m pytest -q %*
endlocal
