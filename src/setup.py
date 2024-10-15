import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["os"], "includes": ["tkinter"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Gerenciador de Planilha Financeira by Euller Julio V0.1.0.3",
    version="0.1.0.3",
    description="Gerencia Planilhas Financeiras",
    options={"build_exe": build_exe_options},
    executables=[Executable("App.py", base=base)]
)