#!/bin/bash
# Script de instalación rápida para macOS/Linux
# Uso: bash setup.sh

echo "========================================"
echo "ReportStock - Instalador rápido"
echo "========================================"
echo ""

echo "Paso 1: Crear entorno virtual..."
python3 -m venv venv
source venv/bin/activate

echo "Paso 2: Instalar dependencias..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "========================================"
echo "Instalación completada!"
echo "========================================"
echo ""
echo "Para ejecutar la aplicación:"
echo "  1. Activar entorno: source venv/bin/activate"
echo "  2. Ejecutar: python main.py"
echo ""
echo "Para generar ejecutable:"
echo "  1. pip install pyinstaller"
echo "  2. pyinstaller reportstock.spec"
echo ""
