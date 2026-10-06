@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

if exist "TEST_SVG_RENDER.png" del /q "TEST_SVG_RENDER.png"

echo [RUN] SVG rendering self-test
"PowerPointLite.exe" --svg-selftest "%CD%\TEST_SVG_RENDER.png"
if errorlevel 1 (
  echo [FAIL] SVG rendering self-test
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo [OK] SVG rendering self-test
echo Output: %CD%\TEST_SVG_RENDER.png
if not defined PPLT_NO_PAUSE pause
exit /b 0
