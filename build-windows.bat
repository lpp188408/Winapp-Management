@echo off
setlocal
chcp 65001 >nul

if /I "%~1" NEQ "__inner" (
  start "Winapp Management Build" cmd /k ""%~f0" __inner"
  exit /b
)

cd /d "%~dp0"
set "LOG=%~dp0build-log.txt"
set "EXE=%~dp0dist\win-x64\WinappManagement.exe"
set "PROJECT=%~dp0src\WinappManagement\WinappManagement.csproj"
set "PUBLISH=%~dp0scripts\publish-windows.ps1"

echo ============================================================
echo Winapp Management build
echo ============================================================
echo.
echo This window will stay open after build.
echo If build fails, send build-log.txt to Codex.
echo.

echo Winapp Management build log>"%LOG%"
echo Time: %date% %time%>>"%LOG%"
echo Folder: %~dp0>>"%LOG%"
echo.>>"%LOG%"

if not exist "%PROJECT%" (
  echo ERROR: Project file not found.
  echo Please make sure you copied the whole Winapp Management folder.
  echo.
  echo Missing project file: %PROJECT%
  echo Missing project file: %PROJECT%>>"%LOG%"
  goto :end_fail
)

if not exist "%PUBLISH%" (
  echo ERROR: Publish script not found.
  echo Please make sure the scripts folder exists.
  echo.
  echo Missing publish script: %PUBLISH%
  echo Missing publish script: %PUBLISH%>>"%LOG%"
  goto :end_fail
)

where dotnet >nul 2>nul
if errorlevel 1 (
  echo ERROR: .NET SDK was not found.
  echo Please install .NET 8 SDK for Windows, not only Runtime.
  echo Download: https://dotnet.microsoft.com/download/dotnet/8.0
  echo.
  echo dotnet command not found.>>"%LOG%"
  goto :end_fail
)

echo Found .NET:
dotnet --info
dotnet --info>>"%LOG%" 2>&1

echo.
echo Building. The first build may take a few minutes.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PUBLISH%" >>"%LOG%" 2>&1
if errorlevel 1 (
  echo ERROR: Build failed.
  echo Details were saved to:
  echo %LOG%
  echo.
  type "%LOG%"
  goto :end_fail
)

if not exist "%EXE%" (
  echo ERROR: Build command finished, but exe was not found.
  echo Expected:
  echo %EXE%
  echo.
  echo Expected output missing: %EXE%>>"%LOG%"
  type "%LOG%"
  goto :end_fail
)

echo.
echo SUCCESS: Build completed.
echo Your exe is here:
echo %EXE%
echo.
goto :end_ok

:end_fail
echo.
echo Build did not complete.
echo Please send build-log.txt to Codex.
echo.
pause
exit /b 1

:end_ok
pause
exit /b 0
