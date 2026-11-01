@echo off
cd /d "%~dp0"
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY (
  echo Python not found. Install Python 3.10+ and check Add to PATH.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv
  ".venv\Scripts\python.exe" -m pip install -U pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
".venv\Scripts\python.exe" -m pip install -U pyinstaller
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed --name "FileConverter" --collect-all pdf2docx --collect-all fitz --hidden-import PIL._tkinter_finder gui.py
echo.
echo Done. EXE: dist\FileConverter\FileConverter.exe
explorer dist\FileConverter
pause
