@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ========================================
echo   Build NamecardGenerator-Win7.exe
echo   打包 Windows 7 兼容版 exe
echo ========================================
echo.

call "%~dp0_env_win7.bat"
if errorlevel 1 (
  echo.
  echo Tip: run setup_win7.bat first to create .venv-win7
  echo 提示：请先运行 setup_win7.bat 创建 .venv-win7
  exit /b 1
)

echo.
echo Installing Win7 app + build dependencies...
"%PY%" -m pip install -r "%~dp0requirements-win7.txt" "pyinstaller==5.13.2"
if errorlevel 1 (
  echo [ERROR] Failed to install build tools.
  pause
  exit /b 1
)

echo.
echo Building NamecardGenerator-Win7.exe ...
"%PY%" -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "NamecardGenerator-Win7" ^
  --icon "assets\app.ico" ^
  --add-data "assets\app.ico;assets" ^
  "%~dp0namecard_generator.py"

if errorlevel 1 (
  echo [ERROR] Build failed.
  pause
  exit /b 1
)

echo.
echo Done. Output: dist\NamecardGenerator-Win7.exe
echo Run on Win7 with: NamecardGenerator-Win7.exe --lang zh
if exist "%~dp0dist" explorer "%~dp0dist"
endlocal
