@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

if exist "DOCUMENT_SAFETY_SELFTEST_OUTPUT" rmdir /s /q "DOCUMENT_SAFETY_SELFTEST_OUTPUT"
mkdir "DOCUMENT_SAFETY_SELFTEST_OUTPUT"

echo ==========================================
echo Deny-by-default document safety self-test
echo ==========================================
echo.

"PowerPointLite.exe" --document-safety-selftest "%CD%\DOCUMENT_SAFETY_SELFTEST_OUTPUT"
if errorlevel 1 (
  echo.
  echo DOCUMENT SAFETY SELF-TEST FAILED
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo.
echo DOCUMENT SAFETY SELF-TEST PASSED
echo Unknown foreign package parts were rejected by edit/conversion safety guards.
echo Test documents were generated entirely by this project.
echo.
if not defined PPLT_NO_PAUSE pause
exit /b 0
