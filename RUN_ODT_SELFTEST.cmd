@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "TEST_ODT_OUTPUT.odt" del /q "TEST_ODT_OUTPUT.odt"

echo [RUN] ODT create/edit/round-trip self-test
"PowerPointLite.exe" --odt-selftest "%CD%\TEST_ODT_OUTPUT.odt"
if errorlevel 1 (
  echo [FAIL] ODT self-test
  pause
  exit /b 1
)

echo [OK] ODT structural self-test completed.
echo Real LibreOffice interoperability still requires manual verification.
pause
