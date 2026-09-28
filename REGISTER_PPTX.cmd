@echo off
setlocal
cd /d "%~dp0"
if not exist "PptxViewer.exe" (
  echo PptxViewer.exe was not found. Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)
set "EXE=%CD%\PptxViewer.exe"
reg add "HKCU\Software\Classes\PptxViewer.File" /ve /d "PowerPoint Presentation - PPTX Viewer" /f >nul
reg add "HKCU\Software\Classes\PptxViewer.File\DefaultIcon" /ve /d "\"%EXE%\",0" /f >nul
reg add "HKCU\Software\Classes\PptxViewer.File\shell\open\command" /ve /d "\"%EXE%\" \"%%1\"" /f >nul
for %%E in (.pptx .pptm) do (
  reg add "HKCU\Software\Classes\%%E\OpenWithProgids" /v "PptxViewer.File" /t REG_NONE /d "" /f >nul
)
echo Registered PPTX Viewer in Windows Open With for .pptx and .pptm.
echo You can choose it from: Right click file ^> Open with ^> Choose another app.
pause
