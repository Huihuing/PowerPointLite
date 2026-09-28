@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ==========================================
echo PPTX Viewer 1.2.1 - EXE Builder
echo Internal PPTX reader included
echo .NET 8 SDK is NOT required
echo ==========================================
echo.

set "CSC="

if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" (
  set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
)

if not defined CSC if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" (
  set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

if not defined CSC (
  echo ERROR: Windows .NET Framework C# compiler was not found.
  pause
  exit /b 1
)

if exist "PptxViewer.exe" del /q "PptxViewer.exe"

"%CSC%" ^
  /nologo ^
  /target:winexe ^
  /platform:anycpu ^
  /optimize+ ^
  /out:"PptxViewer.exe" ^
  /reference:System.dll ^
  /reference:System.Core.dll ^
  /reference:System.Drawing.dll ^
  /reference:System.Windows.Forms.dll ^
  /reference:System.Xml.dll ^
  /reference:System.IO.Compression.dll ^
  /reference:System.IO.Compression.FileSystem.dll ^
  /reference:Microsoft.CSharp.dll ^
  "PptxViewer.cs"

if errorlevel 1 (
  echo.
  echo BUILD FAILED
  pause
  exit /b 1
)

echo.
echo ==========================================
echo BUILD SUCCESS
echo ==========================================
echo.
echo Created:
echo %CD%\PptxViewer.exe
echo.
echo You can now open TEST_INTERNAL_READER.pptx
echo even when PowerPoint and LibreOffice are not installed.
echo.
start "" explorer.exe /select,"%CD%\PptxViewer.exe"
pause
