@echo off
setlocal
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

if exist "TEST_WRITER_OUTPUT.pptx" del /q "TEST_WRITER_OUTPUT.pptx"

start /wait "" "PowerPointLite.exe" --writer-selftest "%CD%\TEST_WRITER_OUTPUT.pptx"

if exist "TEST_WRITER_OUTPUT.pptx" (
  echo.
  echo Writer self-test output created:
  echo %CD%\TEST_WRITER_OUTPUT.pptx
) else (
  echo.
  echo Writer self-test output was not created.
)

pause
