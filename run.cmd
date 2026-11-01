@echo off
setlocal

set "ROOT=%~dp0"
set /p "ISSUE8=Issue date (YYYYMMDD): "

echo(%ISSUE8%| findstr /r "^20[0-9][0-9][0-9][0-9][0-9][0-9]$" >nul
if errorlevel 1 (
  echo Invalid date. Enter 8 digits, for example 20261027.
  pause
  exit /b 1
)

set "WORKDIR=%ROOT%work_history\%ISSUE8%"
if not exist "%WORKDIR%\input\" (
  echo Missing folder: %WORKDIR%\input
  pause
  exit /b 1
)

pushd "%WORKDIR%"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT%_common\windows_setup.ps1"
set "EXITCODE=%ERRORLEVEL%"
popd

if not "%EXITCODE%"=="0" (
  echo.
  echo Automation stopped with an error.
  pause
)

exit /b %EXITCODE%
