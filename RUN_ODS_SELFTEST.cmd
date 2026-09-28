@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "TEST_ODS_OUTPUT.ods" del /q "TEST_ODS_OUTPUT.ods"

echo [RUN] ODS create/edit/round-trip self-test
"PowerPointLite.exe" --ods-selftest "%CD%\TEST_ODS_OUTPUT.ods"
if errorlevel 1 (
  echo [FAIL] ODS self-test
  pause
  exit /b 1
)

echo [OK] ODS structural self-test completed.
echo Real LibreOffice interoperability still requires manual verification.
pause
