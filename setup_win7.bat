@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ========================================
echo   Setup Win7 build environment
echo   配置 Windows 7 构建环境
echo ========================================
echo.
echo This creates .venv-win7 with Python 3.8 and pinned dependencies.
echo 将使用 Python 3.8 创建 .venv-win7 并安装兼容依赖。
echo.

where py >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python launcher ^(py^) not found.
  echo Install Python 3.8 from https://www.python.org/downloads/release/python-3810/
  pause
  exit /b 1
)

py -3.8 --version >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python 3.8 is not installed.
  echo Download: https://www.python.org/downloads/release/python-3810/
  pause
  exit /b 1
)

echo Creating virtual environment: .venv-win7
py -3.8 -m venv "%~dp0.venv-win7"
if errorlevel 1 (
  echo [ERROR] Failed to create .venv-win7
  pause
  exit /b 1
)

set "PY=%~dp0.venv-win7\Scripts\python.exe"
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r "%~dp0requirements-win7.txt"

echo.
echo Done. Use run_win7.bat or build_win7.bat next.
echo 完成。接下来可运行 run_win7.bat 或 build_win7.bat。
pause
endlocal
