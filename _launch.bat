@echo off
setlocal EnableExtensions
cd /d "%~dp0"

rem Usage: _launch.bat [zh|en]
rem Shared runner used by run.bat / run_zh.bat / run_en.bat
set "LANG=%~1"
if /I "%LANG%"=="en" (
  set "LANG=en"
) else (
  set "LANG=zh"
)

call "%~dp0_env.bat"
if errorlevel 1 exit /b 1

echo Language: %LANG%

"%PY%" -c "import PyQt5, win32com" 1>nul 2>nul
if errorlevel 1 (
  echo Installing dependencies...
  "%PY%" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
  )
)

"%PY%" "%~dp0namecard_generator.py" --lang %LANG%
if errorlevel 1 pause
endlocal
