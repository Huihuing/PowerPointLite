@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "TEST_XLSX_OUTPUT.xlsx" del /q "TEST_XLSX_OUTPUT.xlsx"

echo Running XLSX create/read/edit round-trip test...
"PowerPointLite.exe" --xlsx-selftest "%CD%\TEST_XLSX_OUTPUT.xlsx"

if errorlevel 1 (
  echo XLSX SELF-TEST FAILED
  pause
  exit /b 1
)

echo.
echo XLSX self-test process finished.
echo Output: %CD%\TEST_XLSX_OUTPUT.xlsx
echo.
echo Open the generated workbook in Excel or LibreOffice for an additional compatibility check.
pause
