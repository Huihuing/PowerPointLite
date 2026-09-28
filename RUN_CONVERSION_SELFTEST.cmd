@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "CONVERSION_SELFTEST_OUTPUT" rmdir /s /q "CONVERSION_SELFTEST_OUTPUT"
mkdir "CONVERSION_SELFTEST_OUTPUT"

echo [RUN] Cross-format conversion self-test
"PowerPointLite.exe" --conversion-selftest "%CD%\CONVERSION_SELFTEST_OUTPUT"
if errorlevel 1 (
  echo [FAIL] Conversion self-test
  pause
  exit /b 1
)

echo [OK] Guarded text/spreadsheet/presentation conversion self-test completed.
echo Real Office/LibreOffice/Hancom interoperability still requires manual verification.
pause
