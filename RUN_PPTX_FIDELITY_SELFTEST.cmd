@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

if exist "PPTX_FIDELITY_SELFTEST_OUTPUT" rmdir /s /q "PPTX_FIDELITY_SELFTEST_OUTPUT"
mkdir "PPTX_FIDELITY_SELFTEST_OUTPUT"

echo [RUN] PPTX chart / SmartArt fidelity self-test
"PowerPointLite.exe" --pptx-fidelity-selftest "%CD%\PPTX_FIDELITY_SELFTEST_OUTPUT"
if errorlevel 1 (
  echo [FAIL] PPTX chart / SmartArt fidelity self-test
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo [OK] PPTX chart / SmartArt fidelity self-test
echo Output: %CD%\PPTX_FIDELITY_SELFTEST_OUTPUT
if not defined PPLT_NO_PAUSE pause
exit /b 0
