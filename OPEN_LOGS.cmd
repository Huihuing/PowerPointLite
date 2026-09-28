@echo off
set "LOGDIR=%LOCALAPPDATA%\PptxViewer\logs"
if not exist "%LOGDIR%" (
  echo No PPTX Viewer log directory exists yet.
  echo %LOGDIR%
  pause
  exit /b 0
)
start "" explorer.exe "%LOGDIR%"
