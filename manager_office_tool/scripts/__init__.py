"""
Subpaquete scripts de ManagerOfficeScriptTool.

Expone las clases y funciones principales para instalación y
desinstalación de Office.
"""

from .installer import OfficeInstaller
from .uninstaller import OfficeUninstaller, run_uninstallers

__all__ = [
    "OfficeInstaller",
    "OfficeUninstaller",
    "run_uninstallers",
]
