from cx_Freeze import setup, Executable
import sys

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="ReportStock",
    version="1.0",
    description="Sistema de Reportes de Agotados",
    executables=[Executable("main.py", base=base, icon="icon.ico")],
    options={
        "build_exe": {
            "packages": ["tkinter", "customtkinter", "pandas", "openpyxl", "xlrd", "reportlab"],
            "include_files": ["icon.ico"],
        }
    }
)
