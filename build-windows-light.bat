@echo off
setlocal
chcp 65001 >nul

if /I "%~1" NEQ "__inner" (
  start "Winapp Management Light Build" cmd /k ""%~f0" __inner"
  exit /b
)

cd /d "%~dp0"
set "LOG=%~dp0build-light-log.txt"
set "EXE=%~dp0dist\win-x64-light\WinappManagement.exe"
set "PROJECT=%~dp0src\WinappManagement\WinappManagement.csproj"
set "PUBLISH=%~dp0scripts\publish-windows-light.ps1"

echo ============================================================
echo Winapp Management framework-dependent build
echo ============================================================
echo.
echo This build creates a smaller exe, but the target PC must have
echo Microsoft .NET 8 Desktop Runtime installed.
echo.

echo Winapp Management light build log>"%LOG%"
echo Time: %date% %time%>>"%LOG%"
echo Folder: %~dp0>>"%LOG%"
echo.>>"%LOG%"

if not exist "%PROJECT%" (
  echo ERROR: Project file not found.
  echo Missing project file: %PROJECT%>>"%LOG%"
  goto :end_fail
)

if not exist "%PUBLISH%" (
  echo ERROR: Publish script not found.
  echo Missing publish script: %PUBLISH%>>"%LOG%"
  goto :end_fail
)

where dotnet >nul 2>nul
if errorlevel 1 (
  echo ERROR: .NET SDK was not found.
  echo Please install .NET 8 SDK for Windows.
  echo dotnet command not found.>>"%LOG%"
  goto :end_fail
)

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
  echo Expected output missing: %EXE%>>"%LOG%"
  type "%LOG%"
  goto :end_fail
)

echo.
echo SUCCESS: Light build completed.
echo Your exe is here:
echo %EXE%
echo.
echo Reminder: the target PC needs Microsoft .NET 8 Desktop Runtime.
echo.
goto :end_ok

:end_fail
echo.
echo Build did not complete.
echo Please send build-light-log.txt to Codex.
echo.
pause
exit /b 1

:end_ok
pause
exit /b 0
