@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ==========================================
echo PowerPointLite 1.3 - EXE Builder
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

if exist "PowerPointLite.exe" del /q "PowerPointLite.exe"

"%CSC%" ^
  /nologo ^
  /target:winexe ^
  /platform:anycpu ^
  /optimize+ ^
  /out:"PowerPointLite.exe" ^
  /reference:System.dll ^
  /reference:System.Core.dll ^
  /reference:System.Drawing.dll ^
  /reference:System.Windows.Forms.dll ^
  /reference:System.Xml.dll ^
  /reference:System.IO.Compression.dll ^
  /reference:System.IO.Compression.FileSystem.dll ^
  /reference:Microsoft.CSharp.dll ^
  "src\*.cs"

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
echo Created portable executable:
echo %CD%\PowerPointLite.exe
echo.
echo This EXE can be launched directly without BUILD_EXE.cmd.
echo It may also be committed to this private repository if desired.
echo.
start "" explorer.exe /select,"%CD%\PowerPointLite.exe"
pause
