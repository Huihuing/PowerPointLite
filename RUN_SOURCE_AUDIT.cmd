@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "QUIET_ARG="
if /i "%~1"=="--quiet" set "QUIET_ARG=-Quiet"

where powershell.exe >nul 2>nul
if errorlevel 1 (
  echo [FAIL] powershell.exe was not found.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CD%\tools\SourceAudit.ps1" %QUIET_ARG%
set "RESULT=%ERRORLEVEL%"

if not "%RESULT%"=="0" (
  echo.
  echo SOURCE AUDIT FAILED
  if not defined PPLT_NO_PAUSE pause
  exit /b %RESULT%
)

if /i not "%~1"=="--quiet" (
  echo.
  echo SOURCE AUDIT PASSED
  if not defined PPLT_NO_PAUSE pause
)

exit /b 0
