@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

if exist "ANIMATION_SELFTEST_OUTPUT" rmdir /s /q "ANIMATION_SELFTEST_OUTPUT"
mkdir "ANIMATION_SELFTEST_OUTPUT"

echo [RUN] PPTX animation timing self-test
"PowerPointLite.exe" --animation-selftest "%CD%\ANIMATION_SELFTEST_OUTPUT"
if errorlevel 1 (
  echo [FAIL] PPTX animation timing self-test
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo [OK] PPTX animation timing self-test passed.
echo Output: %CD%\ANIMATION_SELFTEST_OUTPUT
if not defined PPLT_NO_PAUSE pause
exit /b 0
