"""
Subpaquete core de ManagerOfficeScriptTool.

Contiene la lógica principal para la gestión, detección y manipulación de
instalaciones de Office.
Expone las clases y funciones clave para uso externo.
"""

from .odt_fetcher import fetch_odt_download_info
from .odt_manager import ODTManager
from .office_installation import OfficeInstallation
from .office_manager import OfficeManager
from .registry_utils import RegistryReader

__all__ = [
    "ODTManager",
    "fetch_odt_download_info",
    "OfficeInstallation",
    "OfficeManager",
    "RegistryReader",
]
