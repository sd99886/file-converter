@echo off
cd /d "%~dp0"
call "%~dp0scripts\ensure_venv.bat"
if errorlevel 1 (
  pause
  exit /b 1
)

echo 正在安装打包依赖...
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 (
  echo [错误] 安装打包依赖失败
  pause
  exit /b 1
)

echo 正在打包...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean FileConverter.spec
if errorlevel 1 (
  echo [错误] 打包失败
  pause
  exit /b 1
)

echo.
echo 完成。程序位置：dist\FileConverter\FileConverter.exe
explorer dist\FileConverter
pause
