from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": [
        "os", "ctypes", "subprocess", "sys", "requests", "platform", 
        "winreg", "tempfile", "shutil", "threading", "datetime", 
        "colorama", "urllib.parse", "tqdm", "tkinter", "tkinter.ttk", 
        "tkinter.messagebox"
    ],
    "excludes": [],
    "include_files": [],
}

setup(
    name="ManagerOfficeScriptTool",
    version="2.4",
    description="Herramienta de Detección, Instalación y Desinstalación Automática de Office",
    options={"build_exe": build_exe_options},
    executables=[Executable(
        "ManagerOfficeScriptTool.py", 
        base="console", 
        icon="icono.ico"
    )],
)
