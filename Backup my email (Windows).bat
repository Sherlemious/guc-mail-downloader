@echo off
setlocal
cd /d "%~dp0"
title University mailbox backup

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo Python 3 is not installed.
    echo Install it from https://www.python.org/downloads/  ^(tick "Add python.exe to PATH"^),
    echo then double-click this file again.
    start "" https://www.python.org/downloads/
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo First run: setting things up, this takes a minute...
    %PY% -m venv .venv || (echo Could not create the Python environment. & pause & exit /b 1)
)
if not exist ".venv\installed.txt" (
    ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt || (echo Could not install the requirements. Check your internet connection. & pause & exit /b 1)
    echo ok> ".venv\installed.txt"
)

".venv\Scripts\python.exe" guc_mail_backup.py %*
