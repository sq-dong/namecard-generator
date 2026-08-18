@echo off
setlocal EnableExtensions
cd /d "%~dp0"

rem Usage: _launch_win7.bat [zh|en]
set "LANG=%~1"
if /I "%LANG%"=="en" (
  set "LANG=en"
) else (
  set "LANG=zh"
)

call "%~dp0_env_win7.bat"
if errorlevel 1 exit /b 1

echo Language: %LANG%
echo Compatibility: Windows 7

"%PY%" -c "import PyQt5, win32com" 1>nul 2>nul
if errorlevel 1 (
  echo Installing Win7 dependencies...
  "%PY%" -m pip install -r "%~dp0requirements-win7.txt"
  if errorlevel 1 (
    echo [ERROR] Failed to install Win7 dependencies.
    pause
    exit /b 1
  )
)

"%PY%" "%~dp0namecard_generator.py" --lang %LANG% --compat win7
if errorlevel 1 pause
endlocal
