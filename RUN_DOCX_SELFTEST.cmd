@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

set "OUT=%CD%\TEST_DOCX_OUTPUT.docx"
if exist "%OUT%" del /q "%OUT%"

echo Running DOCX create/read/edit round-trip self-test...
"PowerPointLite.exe" --docx-selftest "%OUT%"

if not exist "%OUT%" (
  echo.
  echo SELF-TEST OUTPUT WAS NOT CREATED
  pause
  exit /b 1
)

echo.
echo Output:
echo %OUT%
pause
