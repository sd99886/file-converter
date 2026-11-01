@echo off
cd /d "%~dp0"
call "%~dp0scripts\ensure_venv.bat"
if errorlevel 1 (
  pause
  exit /b 1
)

if exist ".venv\Scripts\pythonw.exe" (
  start "" ".venv\Scripts\pythonw.exe" "%~dp0gui.py"
) else (
  ".venv\Scripts\python.exe" "%~dp0gui.py"
  if errorlevel 1 pause
)
