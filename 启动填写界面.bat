@echo off
setlocal
set "HERE=%~dp0"
set "GUI=%HERE%skills\json-skill-builder\scripts\skill_builder_gui.py"

if not exist "%GUI%" set "GUI=%USERPROFILE%\.codex\skills\json-skill-builder\scripts\skill_builder_gui.py"

if not exist "%GUI%" (
    echo.
    echo   skill_builder_gui.py not found.
    echo   Expected next to this file:  skills\json-skill-builder\scripts\skill_builder_gui.py
    echo.
    pause
    exit /b 1
)

set "PYTHONUTF8=1"
set "PYEXE="
where pythonw.exe >nul 2>nul && set "PYEXE=pythonw.exe"
if not defined PYEXE where python.exe >nul 2>nul && set "PYEXE=python.exe"
if not defined PYEXE where py.exe >nul 2>nul && set "PYEXE=py.exe"

if not defined PYEXE (
    echo.
    echo   Python was not found on this computer.
    echo   Install Python 3.10 or newer and tick "Add python.exe to PATH",
    echo   then run:  pip install python-docx
    echo.
    pause
    exit /b 1
)

"%PYEXE%" "%GUI%"
