@echo off
reg delete "HKCU\Software\Classes\PowerPointLite.File" /f >nul 2>nul
for %%E in (.pptx .pptm) do reg delete "HKCU\Software\Classes\%%E\OpenWithProgids" /v "PowerPointLite.File" /f >nul 2>nul
echo PowerPointLite registration removed.
pause
