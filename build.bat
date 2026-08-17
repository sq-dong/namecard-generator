@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ========================================
echo   Build NamecardGenerator.exe
echo ========================================
echo.

call "%~dp0_env.bat"
if errorlevel 1 exit /b 1

echo.
echo Installing app + build dependencies...
"%PY%" -m pip install -r "%~dp0requirements.txt" pyinstaller
if errorlevel 1 (
  echo.
  echo [ERROR] Failed to install build tools via pip.
  echo If you still see "No module named pip", run:
  echo   "%PY%" -m ensurepip --default-pip --upgrade
  echo Then run build.bat again.
  pause
  exit /b 1
)

echo.
echo Building NamecardGenerator.exe ...
"%PY%" -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "NamecardGenerator" ^
  --icon "assets\app.ico" ^
  --add-data "assets\app.ico;assets" ^
  "%~dp0namecard_generator.py"

if errorlevel 1 (
  echo [ERROR] Build failed.
  pause
  exit /b 1
)

echo.
echo Done. Output: dist\NamecardGenerator.exe
if exist "%~dp0dist" explorer "%~dp0dist"
endlocal
