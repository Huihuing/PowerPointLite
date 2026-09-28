@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PPLT_NO_PAUSE=1"

echo ==========================================
echo PowerPointLite pre-merge verification
echo ==========================================
echo.

echo [1/2] Windows .NET Framework build
call BUILD_EXE.cmd --quiet
if errorlevel 1 goto :failed

echo.
echo [2/2] Structural format self-tests
call RUN_ALL_FORMAT_SELFTESTS.cmd
if errorlevel 1 goto :failed

echo.
echo ==========================================
echo PRE-MERGE AUTOMATED CHECKS PASSED
echo ==========================================
echo.
echo This does NOT replace manual Viewer regression and real-app interoperability checks.
echo Required manual checks still include PowerPoint/Word/Excel/LibreOffice/Hancom/PDF viewers as applicable.
echo.
pause
exit /b 0

:failed
echo.
echo ==========================================
echo PRE-MERGE CHECK FAILED
echo ==========================================
echo Fix the first build/self-test error above before merging to main.
echo.
pause
exit /b 1
