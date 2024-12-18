from cx_Freeze import setup, Executable
import os
import sys

script_dir = (
    os.path.dirname(os.path.abspath(sys.argv[0]))
    if getattr(sys, 'frozen', False)
    else os.path.dirname(os.path.abspath(__file__))
)

files_dir = os.path.join(script_dir, 'Files')
xml_path = os.path.join(files_dir, "ODT_ConfigXML")

include_files = [
    (os.path.join(xml_path, "OfficeConfig2013.xml"), 'Files/ODT_ConfigXML/OfficeConfig2013.xml'),
    (os.path.join(xml_path, "OfficeConfig2016.xml"), 'Files/ODT_ConfigXML/OfficeConfig2016.xml'),
    (os.path.join(xml_path, "OfficeConfig2019.xml"), 'Files/ODT_ConfigXML/OfficeConfig2019.xml'),
    (os.path.join(xml_path, "OfficeConfig2021.xml"), 'Files/ODT_ConfigXML/OfficeConfig2021.xml'),
]

build_exe_options = {
    "include_files": include_files,
}

setup(
    name="ManagerOfficeScriptTool",
    version="2.2",
    description="Herramienta de Instalación y Desinstalación Automática de Office",
    options={"build_exe": build_exe_options},
    executables=[Executable("ManagerOfficeScriptTool.py", base="console")],
)
