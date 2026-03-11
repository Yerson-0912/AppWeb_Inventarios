@echo off
REM Script de instalación rápida para Windows
REM Uso: setup.bat

echo ========================================
echo ReportStock - Instalador rápido
echo ========================================
echo.

echo Paso 1: Crear entorno virtual...
python -m venv venv
call venv\Scripts\activate.bat

echo Paso 2: Instalar dependencias...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ========================================
echo Instalación completada!
echo ========================================
echo.
echo Para ejecutar la aplicación:
echo   1. Activar entorno: venv\Scripts\activate.bat
echo   2. Ejecutar: python main.py
echo.
echo Para generar ejecutable:
echo   1. pip install pyinstaller
echo   2. pyinstaller reportstock.spec
echo.
pause
