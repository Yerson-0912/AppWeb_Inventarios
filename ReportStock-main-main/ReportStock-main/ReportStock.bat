@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

REM Obtener la ruta del script actual
set "SCRIPT_DIR=%~dp0"

REM Configurar Python desde el entorno virtual
set "PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"

REM Ejecutar main.py
"%PYTHON_EXE%" "%SCRIPT_DIR%main.py"

pause
