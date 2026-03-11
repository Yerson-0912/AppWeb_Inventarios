@echo off
cd /d d:\Usuario\Downloads\ReportStock-main-main\ReportStock-main-main\ReportStock-main
call D:/Usuario/Downloads/ReportStock-main-main/.venv/Scripts/activate.bat
python -m PyInstaller --onefile --windowed --distpath=dist --name ReportStock main.py
if exist dist\ReportStock.exe (
    echo Copiando archivo a escritorio...
    copy dist\ReportStock.exe "%USERPROFILE%\Desktop\ReportStock.exe"
    echo Completado! El archivo ReportStock.exe esta en tu escritorio
) else (
    echo Error: No se pudo crear el ejecutable
)
pause
