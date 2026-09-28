@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "PDF_SELFTEST_OUTPUT" rmdir /s /q "PDF_SELFTEST_OUTPUT"
mkdir "PDF_SELFTEST_OUTPUT"

echo [RUN] Raster PDF export self-test
"PowerPointLite.exe" --pdf-selftest "%CD%\PDF_SELFTEST_OUTPUT"
if errorlevel 1 (
  echo [FAIL] PDF self-test
  pause
  exit /b 1
)

echo [OK] PDF structural self-test completed.
echo.
echo This exporter currently rasterizes pages.
echo It does not embed the original TTF/OTF font files into PDF output.
echo Visual PDF-reader verification is still required.
pause
