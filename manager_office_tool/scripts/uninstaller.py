"""
uninstaller.py

Módulo encargado de desinstalar instalaciones de Microsoft Office usando ODT.
Incluye generación dinámica del XML de desinstalación y manejo robusto
de errores.
"""

import logging
import re
import subprocess
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
from itertools import groupby
from pathlib import Path
from typing import List, Optional

from colorama import Fore, Style
from manager_office_tool.core import ODTManager, OfficeInstallation
from manager_office_tool.utils import safe_log_path


class OfficeUninstaller:
    """
    Encargado de desinstalar una instalación específica de Microsoft Office
    utilizando ODT.

    Args:
        office_uninstall_dir (str): Directorio temporal donde se descargará y
                                    ejecutará ODT.
        installation (OfficeInstallation): Instalación de Office a eliminar.
        odt_manager (ODTManager, opcional): Instancia de ODTManager a
                                            reutilizar.
    """

    def __init__(
        self,
        office_uninstall_dir: str,
        installations: List[OfficeInstallation],
        odt_manager: Optional[ODTManager] = None,
    ) -> None:
        self.office_uninstall_dir = office_uninstall_dir
        self.installations = installations
        self.setup_path: Optional[Path] = None
        self.odt_manager = odt_manager or ODTManager(office_uninstall_dir)

        logging.debug(
            "[*] Inicializando OfficeUninstaller para grupo: "
            f"{self.installations[0].name} | "
            f"Product: {self.installations[0].product} | "
            f"Idiomas: {[i.client_culture for i in self.installations]}"
        )

    def _generar_configuracion_remocion(self) -> str:
        """
        Genera un archivo XML de configuración para desinstalar la instalación
        de Office actual.

        Returns:
            str: Ruta del archivo de configuración generado.
        """
        configuration = ET.Element("Configuration")
        remove = ET.SubElement(configuration, "Remove")

        base_product = self.installations[0].product
        product = ET.SubElement(remove, "Product", ID=base_product)

        for inst in self.installations:
            culture = normalize_culture(inst.client_culture)
            ET.SubElement(product, "Language", ID=culture)

        ET.SubElement(
            configuration, "Display", Level="None", AcceptEULA="TRUE"
        )

        # Paths
        uninstall_dir = Path(self.office_uninstall_dir)
        file_path = Path(self.odt_manager.office_dir) / "configuration.xml"
        sanitized_uninstall_path = safe_log_path(uninstall_dir)
        sanitized_file_path = safe_log_path(file_path)

        try:
            uninstall_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            msg = (
                "[CONSOLE] No se pudo crear el directorio de desinstalación: "
                f"{sanitized_uninstall_path}"
            )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            raise

        try:
            rough_string = ET.tostring(configuration, encoding="utf-8")
            pretty_xml = minidom.parseString(rough_string).toprettyxml(
                indent="  ", encoding="utf-8"
            )

            file_path.write_bytes(pretty_xml)
            logging.debug(
                "[*] Archivo de configuración XML de desinstalación "
                f"generado en: {sanitized_file_path}"
            )
            return str(file_path)
        except Exception as e:
            logging.exception(
                f"Error {e} al generar el archivo XML de desinstalación en "
                f"{sanitized_file_path}"
            )
            raise

    def ejecutar_desinstalacion(self) -> bool:
        """
        Ejecuta la desinstalación de Office usando el setup.exe y el archivo
        XML generado.

        Returns:
            bool: True si la desinstalación fue exitosa, False si falló.
        """
        base_name = get_base_name(self.installations[0].name)
        idiomas = " / ".join(
            normalize_culture(inst.client_culture)
            for inst in self.installations
        )

        logging.info(
            f"{Fore.GREEN}"
            "Iniciando proceso de desinstalación para: "
            f"{base_name} - {idiomas}"
            f"{Style.RESET_ALL}"
        )

        office_uninstall_path = Path(self.odt_manager.office_dir)
        self.setup_path = office_uninstall_path / "setup.exe"
        sanitized_uninstall_path = safe_log_path(office_uninstall_path)

        # Verifica que setup.exe esté presente antes de intentar la
        # desinstalación
        if not self.setup_path.exists():
            msg = (
                "[CONSOLE] setup.exe no encontrado en "
                f"{sanitized_uninstall_path}"
            )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False

        try:
            # Genera el archivo XML de configuración necesario para la
            # desinstalación
            config_path = self._generar_configuracion_remocion()
        except Exception:
            msg = "[CONSOLE] Error al generar el archivo de configuración XML."
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False

        try:
            logging.info(
                f"{Fore.CYAN}"
                "Ejecutando: setup.exe /configure configuration.xml"
                f"{Style.RESET_ALL}"
            )

            result = subprocess.run(
                [str(self.setup_path), "/configure", config_path],
                cwd=str(office_uninstall_path),
                capture_output=True,
                text=True,
                check=True,
            )

            logging.debug(f"setup.exe stdout:\n{result.stdout}")

            return True

        except subprocess.CalledProcessError as e:
            logging.error(
                "setup.exe falló con código "
                f"{e.returncode}.\nStderr:\n{e.stderr}"
            )
        except Exception as e:
            logging.exception(
                f"Error inesperado durante la desinstalación. {e}"
            )

        return False

    def execute(self) -> str:
        """
        Punto de entrada externo para iniciar la desinstalación.

        Returns:
            str: Mensaje de estado coloreado según el resultado.
        """
        try:
            if self.ejecutar_desinstalacion():
                base_name = get_base_name(self.installations[0].name)
                idiomas = " / ".join(
                    inst.client_culture for inst in self.installations
                )
                return (
                    f"{Fore.GREEN}"
                    f"Desinstalado exitosamente: {base_name} - {idiomas}"
                    f"{Style.RESET_ALL}"
                )
            else:
                return f"{Fore.RED}Error al desinstalar: {self.installations[0].name}{Style.RESET_ALL}"  # noqa E501
        except Exception as e:
            logging.exception(
                "Excepción en 'execute()' para "
                f"{self.installations[0].name}: {e}"
            )
            return (
                f"{Fore.RED}"
                f"Falló inesperadamente: {self.installations[0].name}"
                f"{Style.RESET_ALL}"
            )


def get_base_name(name: str) -> str:
    # Esto quita el código de idioma al final en formato " - xx-xx"
    return re.sub(r"\s-\s[a-z]{2}-[a-z]{2}$", "", name)


def normalize_culture(culture: str) -> str:
    """
    Normaliza el código de cultura al formato xx-XX:
    - idioma en minúsculas
    - región en mayúsculas
    """
    culture = culture.strip()
    if len(culture) == 5 and culture[2] == "-":
        return culture[:2].lower() + "-" + culture[3:].upper()
    return culture.lower()


def run_uninstallers(
    installations: List[OfficeInstallation], uninstall_dir: Path
) -> None:
    """
    Ejecuta secuencialmente la desinstalación de múltiples instalaciones de
    Office utilizando ODT.

    Args:
        installations (List[OfficeInstallation]): Lista de instalaciones a
                                                    desinstalar.
        uninstall_dir (Path): Directorio temporal para ODT.
    """
    logging.info(
        f"{Fore.LIGHTCYAN_EX}"
        f"Iniciando desinstalación de {len(installations)} instalación(es)..."
        f"{Style.RESET_ALL}"
    )

    odt_managers: dict[str, ODTManager] = {}

    # Primero ordenar para agrupar correctamente
    installations.sort(key=lambda x: (get_base_name(x.name), x.product))

    for (base_name, product), group in groupby(
        installations, key=lambda x: (get_base_name(x.name), x.product)
    ):
        group_list = list(group)

        logging.info(
            f"{Fore.LIGHTYELLOW_EX}"
            f"Desinstalando grupo: {base_name} - ({len(group_list)} "
            f"idioma(s))"
            f"{Style.RESET_ALL}"
        )
        logging.info(f"{Fore.LIGHTCYAN_EX}{'-' * 80}{Style.RESET_ALL}")

        familia = "2013" if "2013" in base_name else "Modern"
        if familia not in odt_managers:
            odt_dir = uninstall_dir / f"OfficeODT_{familia}"
            manager = ODTManager(str(odt_dir))
            odt_managers[familia] = manager

            logging.info(
                f"{Fore.GREEN}"
                f"Preparando ODT para familia {familia}…"
                f"{Style.RESET_ALL}"
            )
            if not manager.download_and_extract(familia):
                msg = f"[CONSOLE] Error al preparar ODT para familia {familia}"
                logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
                continue

        uninstaller = OfficeUninstaller(
            str(uninstall_dir), group_list, odt_managers[familia]
        )

        try:
            result = uninstaller.execute()
            logging.info(result)
        except Exception as e:
            msg = f"[CONSOLE] Error al desinstalar {base_name}: {e}"
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")

        logging.info(f"{Fore.LIGHTCYAN_EX}{'-' * 80}{Style.RESET_ALL}")
