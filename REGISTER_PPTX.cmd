@echo off
setlocal
cd /d "%~dp0"

if not exist "PowerPointLite.exe" (
  echo PowerPointLite.exe was not found. Run BUILD_EXE.cmd first.
  pause
  exit /b 1
)

set "EXE=%CD%\PowerPointLite.exe"

reg add "HKCU\Software\Classes\PowerPointLite.File" /ve /d "PowerPoint Presentation - PowerPointLite" /f >nul
reg add "HKCU\Software\Classes\PowerPointLite.File\DefaultIcon" /ve /d "\"%EXE%\",0" /f >nul
reg add "HKCU\Software\Classes\PowerPointLite.File\shell\open\command" /ve /d "\"%EXE%\" \"%%1\"" /f >nul

for %%E in (.pptx .pptm) do (
  reg add "HKCU\Software\Classes\%%E\OpenWithProgids" /v "PowerPointLite.File" /t REG_NONE /d "" /f >nul
)

echo Registered PowerPointLite in Windows Open With for .pptx and .pptm.
echo Right click a presentation ^> Open with ^> Choose another app.
pause
