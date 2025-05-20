"""
office_manager.py

Administra la detección y visualización de instalaciones de Microsoft Office
en el sistema, consultando el registro de Windows y utilizando la clase
OfficeInstallation para representar cada hallazgo.
"""

import re
from pathlib import Path
from typing import List

from colorama import Fore
from core.office_installation import OfficeInstallation
from core.registry_utils import RegistryReader


class OfficeManager:
    """
    Administra la detección y visualización de instalaciones de
    Microsoft Office.

    Args:
        show_all (bool): Si es True, mostrará todas las instalaciones
                            encontradas.
                         Si es False, solo mostrará las instalaciones de
                            Microsoft Office o 365.
    """

    def __init__(self, show_all: bool = False) -> None:
        self.show_all = show_all
        self.registry = RegistryReader()
        self.installations: List[OfficeInstallation] = []

    def _get_installations(self) -> List[OfficeInstallation]:
        """
        Busca y almacena las instalaciones de Microsoft Office detectadas
        en el sistema.

        Returns:
            List[OfficeInstallation]: Lista de objetos OfficeInstallation
                encontrados.
        """
        office_keys = [
            r"SOFTWARE\Microsoft\Office",
            r"SOFTWARE\Wow6432Node\Microsoft\Office",
        ]
        uninstall_keys = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        ]
        versions = ["ClickToRun", "15.0", "16.0"]
        found_names: set[str] = set()
        installations: List[OfficeInstallation] = []

        # Recorre las claves de registro para encontrar instalaciones de Office
        # y sus respectivas configuraciones.
        for office_key in office_keys:
            for version in versions:
                if version == "ClickToRun":
                    version_key = str(
                        Path(office_key) / version / "Configuration"
                    )
                else:
                    version_key = str(
                        Path(office_key)
                        / version
                        / "ClickToRun"
                        / "Configuration"
                    )

                platform_value = self.registry.get_registry_value(
                    version_key, "Platform"
                )
                if not platform_value:
                    continue

                updates_enabled = (
                    self.registry.get_registry_value(
                        version_key, "UpdatesEnabled"
                    )
                    == "True"
                )
                update_url = self.registry.get_registry_value(
                    version_key, "CDNBaseUrl"
                )

                product_ids_raw = self.registry.get_registry_value(
                    version_key, "ProductReleaseIds"
                )
                product_id_map = {}

                if product_ids_raw:
                    product_list = [
                        p.strip() for p in product_ids_raw.split(",")
                    ]
                    for pid in product_list:
                        media = self.registry.get_registry_value(
                            version_key, f"{pid}.MediaType"
                        )
                        product_id_map[pid] = media if media else ""

                for uninstall_key in uninstall_keys:
                    subkeys = self.registry.get_registry_keys(uninstall_key)
                    for subkey in subkeys:
                        uninstall_key_path = str(Path(uninstall_key) / subkey)
                        display_name = self.registry.get_registry_value(
                            uninstall_key_path, "DisplayName"
                        )
                        if not display_name or display_name in found_names:
                            continue

                        # Filtra las instalaciones de Microsoft Office
                        # según el nombre mostrado en el registro.
                        if not self.show_all and not (
                            "Microsoft Office" in display_name
                            or "Microsoft 365" in display_name
                            or "Microsoft Project" in display_name
                            or "Microsoft Visio" in display_name
                        ):
                            continue

                        found_names.add(display_name)
                        display_version = self.registry.get_registry_value(
                            uninstall_key_path, "DisplayVersion"
                        )
                        install_location = self.registry.get_registry_value(
                            uninstall_key_path, "InstallLocation"
                        )
                        uninstall_string = self.registry.get_registry_value(
                            uninstall_key_path, "UninstallString"
                        )
                        click_to_run = "ClickToRun" in uninstall_string

                        match_product = re.search(
                            r"productstoremove=([A-Za-z]+)", uninstall_string
                        )
                        product_id = (
                            match_product.group(1) if match_product else ""
                        )

                        media_type = product_id_map.get(product_id, "")

                        # Crea el objeto OfficeInstallation
                        # y lo agrega a la lista de instalaciones.
                        installations.append(
                            OfficeInstallation(
                                name=display_name,
                                version=display_version,
                                install_path=install_location,
                                click_to_run=click_to_run,
                                product="",
                                bitness="",
                                updates_enabled=updates_enabled,
                                update_url=update_url,
                                client_culture="",
                                media_type=media_type,
                                uninstall_string=uninstall_string,
                            )
                        )

        self.installations = installations
        return installations

    def display_installations(self) -> None:
        """
        Imprime por consola las instalaciones de Office encontradas con
        formato visual.
        """
        installations = self._get_installations()
        if not installations:
            print(
                Fore.YELLOW
                + ("No se encontraron instalaciones de Microsoft Office.")
            )
            return

        print(Fore.CYAN + "=" * 110)
        print(
            Fore.GREEN
            + (
                "Se encontraron las siguientes instalaciones de "
                "Microsoft Office:"
            )
        )
        print(Fore.CYAN + "=" * 110)

        # Muestra información detallada de cada instalación
        # en un formato legible y estructurado.
        for install in installations:
            office_info = {
                "DisplayName": install.name,
                "Version": install.version,
                "Language": install.client_culture,
                "Architecture": install.bitness,
                "IsClickToRun": install.click_to_run,
                "InstallPath": install.install_path,
                "ProductID": install.product,
                "IsUpdatesEnabled": install.updates_enabled,
                "UpdateChannelUrl": install.update_url,
                "MediaType": install.media_type,
            }
            for key, value in office_info.items():
                print(Fore.LIGHTWHITE_EX + f"{key:<20}: {str(value)}")
            print(Fore.CYAN + "=" * 110)

    def get_installations(self) -> List[OfficeInstallation]:
        """
        Retorna la lista de instalaciones encontradas.

        Returns:
            List[OfficeInstallation]: Lista de instalaciones actuales.
        """
        return self._get_installations()
