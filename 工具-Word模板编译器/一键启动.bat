@echo off
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYEXE="

where pythonw.exe >nul 2>nul && set "PYEXE=pythonw.exe"
if not defined PYEXE where python.exe >nul 2>nul && set "PYEXE=python.exe"
if not defined PYEXE where py.exe >nul 2>nul && set "PYEXE=py.exe"

if not defined PYEXE (
    echo.
    echo   Python was not found on this computer.
    echo.
    echo   Install Python 3.10 or newer from https://www.python.org/downloads/
    echo   Tick "Add python.exe to PATH" while installing, then run:
    echo.
    echo       pip install python-docx
    echo.
    pause
    exit /b 1
)

"%PYEXE%" "start_tool.py" %*
