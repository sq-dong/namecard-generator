@echo off
rem Shared helper: locate a working Python and ensure pip is available.
rem Caller:  call "%~dp0_env.bat"
rem On success: PY is set. On failure: exit /b 1 (and pause).

set "PY="

if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY if exist "D:\myenv\Scripts\python.exe" set "PY=D:\myenv\Scripts\python.exe"
if not defined PY if exist "D:\Anaconda\python.exe" set "PY=D:\Anaconda\python.exe"

if not defined PY (
  where py >nul 2>nul && for /f "delims=" %%I in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%I"
)

if not defined PY (
  where python >nul 2>nul && for /f "delims=" %%I in ('where python') do (
    if not defined PY set "PY=%%I"
  )
)

if not defined PY (
  echo [ERROR] Python not found. Install Python 3.9+ or create a venv.
  pause
  exit /b 1
)

echo Using: %PY%
"%PY%" --version
if errorlevel 1 (
  echo [ERROR] Failed to start Python:
  echo   %PY%
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
    echo.
    echo Try:
    echo   "%PY%" -m ensurepip --default-pip --upgrade
    echo Or use D:\myenv / Anaconda Python instead.
    pause
    exit /b 1
  )
)

exit /b 0
