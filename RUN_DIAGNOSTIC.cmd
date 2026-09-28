@echo off
setlocal
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found.
  echo Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

echo Starting PowerPointLite.exe...
start "" "PowerPointLite.exe"

timeout /t 3 /nobreak >nul

if exist "%LOCALAPPDATA%\PptxViewer\logs\last-startup.log" (
  echo.
  echo Last startup log:
  echo ------------------------------------------
  type "%LOCALAPPDATA%\PptxViewer\logs\last-startup.log"
  echo ------------------------------------------
) else (
  echo No startup log found.
)

echo.
echo If the viewer is open, startup succeeded.
pause
