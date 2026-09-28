@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "TEST_ODP_OUTPUT.odp" del /q "TEST_ODP_OUTPUT.odp"

echo [RUN] ODP create/edit/round-trip self-test
"PowerPointLite.exe" --odp-selftest "%CD%\TEST_ODP_OUTPUT.odp"
if errorlevel 1 (
  echo [FAIL] ODP self-test
  pause
  exit /b 1
)

echo [OK] ODP structural self-test completed.
echo Real LibreOffice interoperability still requires manual verification.
pause
