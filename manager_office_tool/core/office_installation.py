"""
office_installation.py

Define la clase OfficeInstallation, que representa una instalación detectada
de Microsoft Office en el sistema. Incluye lógica para extraer información
relevante desde el comando de desinstalación.
"""

import re


class OfficeInstallation:
    """
    Representa una instalación de Microsoft Office detectada en el sistema.

    Atributos principales:
        name (str): Nombre del producto Office.
        version (str): Versión instalada.
        install_path (str): Ruta de instalación.
        click_to_run (bool): Si es instalación Click-to-Run.
        product (str): ID del producto.
        bitness (str): Arquitectura ("32-Bits" o "64-Bits").
        updates_enabled (bool): Si las actualizaciones están habilitadas.
        update_url (str): URL del canal de actualizaciones.
        client_culture (str): Idioma de la instalación.
        version_to_report (str): Versión informada del cliente.
        media_type (str): Tipo de medio de instalación.
        uninstall_string (str): Comando de desinstalación.
    """

    def __init__(
        self,
        name: str,
        version: str,
        install_path: str,
        click_to_run: bool,
        product: str,
        bitness: str,
        updates_enabled: bool,
        update_url: str,
        client_culture: str,
        media_type: str,
        uninstall_string: str,
    ) -> None:
        self.name = name
        self.version = version
        self.install_path = install_path
        self.click_to_run = click_to_run
        self.uninstall_string = uninstall_string
        self.updates_enabled = updates_enabled
        self.bitness_original = bitness
        self.product_original = product
        self.update_url = update_url
        self.media_type = media_type

        # Extrae el idioma (culture) desde la cadena de desinstalación
        # si está presente
        match_culture = re.search(r"culture=([a-zA-Z\-]+)", uninstall_string)
        self.client_culture = (
            match_culture.group(1) if match_culture else client_culture
        )

        # Extrae la arquitectura (bitness) desde la cadena de desinstalación
        # si está presente
        match_platform = re.search(
            r"platform=(x86|x64)", uninstall_string, re.IGNORECASE
        )
        if match_platform:
            self.bitness = (
                "64-Bits"
                if match_platform.group(1).lower() == "x64"
                else "32-Bits"
            )
        else:
            self.bitness = bitness

        # Usa siempre el product pasado por el constructor
        self.product = product

    def __repr__(self) -> str:
        """
        Devuelve una representación legible de la instalación,
        útil para depuración y visualización rápida.
        """
        return (
            f"<OfficeInstallation: {self.name} ({self.version}) - "
            f"{self.client_culture} - {self.bitness}>"
        )
