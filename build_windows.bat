@echo off
REM Wrapper to run BUILD_WINDOWS.ps1 from cmd
set SCRIPT_DIR=%~dp0
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%BUILD_WINDOWS.ps1" %*
pause
