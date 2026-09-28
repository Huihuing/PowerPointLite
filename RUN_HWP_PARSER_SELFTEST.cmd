@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

echo Running HWP 5.x FileHeader / record parser synthetic test...
"PowerPointLite.exe" --hwp-parser-selftest

if errorlevel 1 (
  echo HWP PARSER SELF-TEST FAILED
  pause
  exit /b 1
)

echo.
echo HWP parser self-test process finished.
echo.
echo NOTE: This test does not use or redistribute a third-party HWP document.
echo Test real HWP files only when you have the right to use those fixtures.
pause
