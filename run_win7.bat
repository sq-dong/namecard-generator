@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ========================================
echo   Namecard Generator - Windows 7
echo   名签生成器 - Windows 7 版
echo ========================================
echo.
echo   1. Chinese UI  / 中文界面
echo   2. English UI  / 英文界面
echo.
choice /C 12 /N /M "Select / 请选择 (1/2): "
if errorlevel 2 (
  call "%~dp0_launch_win7.bat" en
) else (
  call "%~dp0_launch_win7.bat" zh
)
endlocal
