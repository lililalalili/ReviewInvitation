@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ==========================================
echo Review Invitation Agent
echo Windows runner + Ubuntu-hosted Ollama
echo ==========================================
echo.

set "SCRIPT=review_invitation_agent_windows_ubuntu_ollama_v14.py"

if not exist "%SCRIPT%" (
    echo ERROR: %SCRIPT% was not found in:
    echo %cd%
    echo.
    pause
    exit /b 1
)

set "PY_CMD="
where py >nul 2>nul
if %errorlevel%==0 set "PY_CMD=py -3"
if not defined PY_CMD (
    where python >nul 2>nul
    if %errorlevel%==0 set "PY_CMD=python"
)
if not defined PY_CMD (
    echo ERROR: Python was not found in PATH.
    pause
    exit /b 1
)

if "%OLLAMA_BASE_URL%"=="" (
    echo ERROR: OLLAMA_BASE_URL is not set.
    echo Example:
    echo   setx OLLAMA_BASE_URL "http://192.168.106.105:11434/api"
    echo.
    pause
    exit /b 1
)

if "%OLLAMA_MODEL%"=="" (
    echo ERROR: OLLAMA_MODEL is not set.
    echo Example:
    echo   setx OLLAMA_MODEL "nemotron-3-super:120b"
    echo.
    pause
    exit /b 1
)

if "%PUBMED_EMAIL%"=="" (
    echo ERROR: PUBMED_EMAIL is not set.
    echo Example:
    echo   setx PUBMED_EMAIL "your_email@example.com"
    echo.
    pause
    exit /b 1
)

echo Using Python command: %PY_CMD%
echo Working directory: %cd%
echo Script: %SCRIPT%
echo OLLAMA_BASE_URL: %OLLAMA_BASE_URL%
echo OLLAMA_MODEL: %OLLAMA_MODEL%
echo Start time: %date% %time%
echo.

%PY_CMD% "%SCRIPT%"
set "ERR=%ERRORLEVEL%"

echo.
echo End time: %date% %time%
if not "%ERR%"=="0" (
    echo Script exited with error code: %ERR%
) else (
    echo Script finished successfully.
)
echo.
pause
exit /b %ERR%
