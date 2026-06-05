@echo off
REM Запускает main.py на Windows. Не модифицирует исходники.
REM Если есть виртуальное окружение venv, используем его, иначе системный python.

SETLOCAL ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

IF EXIST "%~dp0venv\Scripts\python.exe" (
  SET PYEXEC=%~dp0venv\Scripts\python.exe
) ELSE (
  SET PYEXEC=python
)

"%PYEXEC%" main.py %*
EXIT /B %ERRORLEVEL%
