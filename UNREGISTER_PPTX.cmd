@echo off
reg delete "HKCU\Software\Classes\PptxViewer.File" /f >nul 2>nul
for %%E in (.pptx .pptm) do reg delete "HKCU\Software\Classes\%%E\OpenWithProgids" /v "PptxViewer.File" /f >nul 2>nul
echo PPTX Viewer registration removed.
pause
