@echo off
cd /d "%~dp0"
set PY=python
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
set PYTHONIOENCODING=utf-8
if "%1"=="auto" (
  %PY% gov_support_watch.py >> data\run_log.txt 2>&1
) else (
  %PY% gov_support_watch.py --open
)
