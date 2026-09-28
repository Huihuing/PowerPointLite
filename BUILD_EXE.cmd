@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

echo ==========================================
echo PowerPointLite - EXE Builder
echo Recursive source build enabled
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

if not exist "src" (
  echo ERROR: src directory was not found.
  pause
  exit /b 1
)

set "SOURCE_LIST=%TEMP%\PowerPointLite_sources_%RANDOM%_%RANDOM%.rsp"
if exist "%SOURCE_LIST%" del /q "%SOURCE_LIST%"

set /a SOURCE_COUNT=0
for /r "src" %%F in (*.cs) do (
  >>"%SOURCE_LIST%" echo "%%~fF"
  set /a SOURCE_COUNT+=1
)

if !SOURCE_COUNT! LEQ 0 (
  echo ERROR: No C# source files were found under src.
  if exist "%SOURCE_LIST%" del /q "%SOURCE_LIST%"
  pause
  exit /b 1
)

echo Found !SOURCE_COUNT! C# source file(s).
echo.

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
  @"%SOURCE_LIST%"

set "BUILD_RESULT=%ERRORLEVEL%"
if exist "%SOURCE_LIST%" del /q "%SOURCE_LIST%"

if not "%BUILD_RESULT%"=="0" (
  echo.
  echo BUILD FAILED
  pause
  exit /b %BUILD_RESULT%
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
echo.
start "" explorer.exe /select,"%CD%\PowerPointLite.exe"
pause
