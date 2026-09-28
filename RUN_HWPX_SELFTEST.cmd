@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "TEST_HWPX_OUTPUT.hwpx" del /q "TEST_HWPX_OUTPUT.hwpx"

echo Running HWPX create/read/edit round-trip structural test...
"PowerPointLite.exe" --hwpx-selftest "%CD%\TEST_HWPX_OUTPUT.hwpx"

if errorlevel 1 (
  echo HWPX SELF-TEST FAILED
  pause
  exit /b 1
)

echo.
echo HWPX self-test process finished.
echo Output: %CD%\TEST_HWPX_OUTPUT.hwpx
echo.
echo NOTE: This structural test does not prove Hancom compatibility.
echo Open the generated file in a real Hancom environment before marking HWPX stable.
pause
