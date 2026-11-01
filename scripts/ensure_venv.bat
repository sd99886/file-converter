@echo off
rem 在项目根目录准备 .venv，缺依赖时才安装 requirements.txt。
cd /d "%~dp0.."

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY where python3 >nul 2>&1 && set "PY=python3"
if not defined PY (
  echo [错误] 未找到 Python。请安装 Python 3.10 或更高版本：
  echo   https://www.python.org/downloads/
  echo 安装时请勾选 "Add python.exe to PATH"。
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo 正在创建虚拟环境...
  %PY% -m venv .venv
  if errorlevel 1 (
    echo [错误] 创建 .venv 失败
    exit /b 1
  )
)

".venv\Scripts\python.exe" -c "import fitz, PIL, pptx, natsort, pdf2docx" >nul 2>&1
if errorlevel 1 (
  echo 正在安装依赖...
  ".venv\Scripts\python.exe" -m pip install -U pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo [错误] 安装依赖失败
    exit /b 1
  )
)

exit /b 0
