@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  if not defined PPLT_NO_PAUSE pause
  exit /b 1
)

echo ==========================================
echo PowerPointLite format self-tests
echo PPTX / DOCX / XLSX / HWPX / HWP
echo ODT / ODS / ODP / PDF
echo ==========================================
echo.

call :run --writer-selftest "TEST_WRITER_OUTPUT.pptx" "PPTX"
if errorlevel 1 goto :failed

call :run --docx-selftest "TEST_DOCX_OUTPUT.docx" "DOCX"
if errorlevel 1 goto :failed

call :run --xlsx-selftest "TEST_XLSX_OUTPUT.xlsx" "XLSX"
if errorlevel 1 goto :failed

call :run --hwpx-selftest "TEST_HWPX_OUTPUT.hwpx" "HWPX"
if errorlevel 1 goto :failed

call :run --odt-selftest "TEST_ODT_OUTPUT.odt" "ODT"
if errorlevel 1 goto :failed

call :run --ods-selftest "TEST_ODS_OUTPUT.ods" "ODS"
if errorlevel 1 goto :failed

call :run --odp-selftest "TEST_ODP_OUTPUT.odp" "ODP"
if errorlevel 1 goto :failed

echo [RUN] HWP parser
"PowerPointLite.exe" --hwp-parser-selftest
if errorlevel 1 (
  echo [FAIL] HWP parser
  goto :failed
)
echo [OK] HWP parser process completed
echo.

if exist "PDF_SELFTEST_OUTPUT" rmdir /s /q "PDF_SELFTEST_OUTPUT"
mkdir "PDF_SELFTEST_OUTPUT"
echo [RUN] PDF raster exporter
"PowerPointLite.exe" --pdf-selftest "%CD%\PDF_SELFTEST_OUTPUT"
if errorlevel 1 (
  echo [FAIL] PDF raster exporter
  goto :failed
)
echo [OK] PDF raster exporter process completed
echo.

echo ==========================================
echo ALL STRUCTURAL SELF-TESTS FINISHED
echo ==========================================
echo.
echo Additional manual checks are still required:
echo - PowerPointLite Viewer regression
echo - Microsoft PowerPoint / LibreOffice for PPTX
echo - Microsoft Word / LibreOffice for DOCX
echo - Microsoft Excel / LibreOffice for XLSX
echo - Hancom for HWPX
echo - Rights-cleared real HWP 5.x samples for HWP read-only verification
echo - LibreOffice for ODT / ODS / ODP
echo - PDF reader visual inspection and print verification
echo.
echo PDF export currently rasterizes pages and does not embed source font files.
echo.
if not defined PPLT_NO_PAUSE pause
exit /b 0

:run
echo [RUN] %~3
if exist "%~2" del /q "%~2"
"PowerPointLite.exe" %~1 "%CD%\%~2"
if errorlevel 1 (
  echo [FAIL] %~3
  exit /b 1
)
echo [OK] %~3 process completed
echo.
exit /b 0

:failed
echo.
echo ONE OR MORE FORMAT SELF-TESTS FAILED.
if not defined PPLT_NO_PAUSE pause
exit /b 1
