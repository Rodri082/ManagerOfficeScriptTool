from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["os", "subprocess", "requests", "urllib.parse", "sys", "tkinter", "shutil", "threading", "datetime", "colorama"],
    "excludes": [],
    "include_files": [],
    "optimize": 2,
}

exe = Executable(
    script="DeploymentScriptTool.py",
    base=None,
)

setup(
    name="DeploymentScriptTool",
    version="1.0",
    description="Herramienta de configuración de ODT",
    options={"build_exe": build_exe_options},
    executables=[exe],
)
