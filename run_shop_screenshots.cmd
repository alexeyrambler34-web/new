@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0screenshots_logs" mkdir "%~dp0screenshots_logs"
set LOG=%~dp0screenshots_logs\cmd_start.log
echo ==============================>> "%LOG%"
echo CMD started: %DATE% %TIME%>> "%LOG%"
echo Folder: %CD%>> "%LOG%"
echo Starting PowerShell...>> "%LOG%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0capture_shop_screenshots.ps1" >> "%LOG%" 2>&1
echo PowerShell exit code: %ERRORLEVEL%>> "%LOG%"
echo.>> "%LOG%"
type "%LOG%"
echo.
pause
endlocal
