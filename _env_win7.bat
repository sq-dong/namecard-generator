@echo off
rem Locate Python 3.8 for Windows 7 compatibility.
rem Caller: call "%~dp0_env_win7.bat"
rem On success: PY is set. On failure: exit /b 1.

set "PY="

if exist "%~dp0.venv-win7\Scripts\python.exe" set "PY=%~dp0.venv-win7\Scripts\python.exe"

if not defined PY (
  where py >nul 2>nul && for /f "delims=" %%I in ('py -3.8 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%I"
)

if not defined PY (
  where python >nul 2>nul && for /f "delims=" %%I in ('where python') do (
    if not defined PY (
      "%%I" -c "import sys; raise SystemExit(0 if sys.version_info[:2]==(3,8) else 1)" >nul 2>nul
      if not errorlevel 1 set "PY=%%I"
    )
  )
)

if not defined PY (
  echo [ERROR] Python 3.8 not found.
  echo.
  echo Windows 7 requires Python 3.8.x ^(3.9+ does not run on Win7^).
  echo Run setup_win7.bat first, or install Python 3.8 from python.org.
  pause
  exit /b 1
)

echo Using Win7 Python: %PY%
"%PY%" --version
if errorlevel 1 (
  echo [ERROR] Failed to start Python:
  echo   %PY%
  pause
  exit /b 1
)

"%PY%" -c "import sys; raise SystemExit(0 if sys.version_info < (3,9) else 1)"
if errorlevel 1 (
  echo [ERROR] Selected Python is too new for Windows 7 builds.
  echo   %PY%
  echo Use Python 3.8.x ^(run setup_win7.bat to create .venv-win7^).
  pause
  exit /b 1
)

"%PY%" -m pip --version >nul 2>nul
if errorlevel 1 (
  echo pip missing, bootstrapping with ensurepip...
  "%PY%" -m ensurepip --default-pip --upgrade
  "%PY%" -m pip --version >nul 2>nul
  if errorlevel 1 (
    echo [ERROR] Could not restore pip for:
    echo   %PY%
    pause
    exit /b 1
  )
)

exit /b 0
