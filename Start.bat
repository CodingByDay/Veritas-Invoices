@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1
if errorlevel 1 (
  python app.py
  if errorlevel 1 pause
) else (
  start "" pythonw app.py
)
