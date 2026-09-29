@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo ==========================================
echo Font licensing parser/policy self-test
echo ==========================================
echo.

"PowerPointLite.exe" --font-license-selftest
if errorlevel 1 (
  echo.
  echo FONT LICENSE SELF-TEST FAILED
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo.
echo FONT LICENSE SELF-TEST PASSED
echo Synthetic font metadata only; no third-party font file is bundled.
echo.
if not defined PPLT_NO_PAUSE pause
exit /b 0
