"""
uninstaller.py

Módulo encargado de desinstalar instalaciones de Microsoft Office usando ODT.
Incluye generación dinámica del XML de desinstalación y manejo robusto
de errores.
"""

import logging
import subprocess
from pathlib import Path
from typing import List

from colorama import Fore
from core.odt_manager import ODTManager
from core.office_installation import OfficeInstallation
from utils import safe_log_path


class OfficeUninstaller:
    """
    Encargado de desinstalar una instalación específica de Microsoft Office
    utilizando el Office Deployment Tool (ODT).

    Args:
        office_uninstall_dir (str): Directorio temporal donde se descargará
            y ejecutará ODT.
        installation (OfficeInstallation): Objeto que representa la
            instalación de Office que será eliminada.
    """

    def __init__(
        self, office_uninstall_dir: str, installation: OfficeInstallation
    ) -> None:
        self.office_uninstall_dir = office_uninstall_dir
        self.installation = installation
        self.setup_path: Path | None = None

        logging.info(
            "Inicializando OfficeUninstaller para: "
            f"{self.installation.name} | "
            f"Product: {self.installation.product} | "
            f"Idioma: {self.installation.client_culture}"
        )

    def _generar_configuracion_remocion(self) -> str:
        """
        Genera un archivo XML de configuración para desinstalar la instalación
        de Office actual.

        Returns:
            str: Ruta del archivo de configuración generado.
        """
        xml_content = f"""\
<Configuration>
  <Remove>
    <Product ID="{self.installation.product}">
      <Language ID="{self.installation.client_culture}" />
    </Product>
  </Remove>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>"""

        uninstall_dir = Path(self.office_uninstall_dir)
        file_path = uninstall_dir / "configuration.xml"
        sanitized_uninstall_path = safe_log_path(uninstall_dir)
        sanitized_file_path = safe_log_path(file_path)
        try:
            uninstall_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            logging.error(
                "No se pudo crear el directorio de desinstalación: "
                f"{sanitized_uninstall_path}"
            )
            raise

        try:
            file_path.write_text(xml_content, encoding="utf-8")
            print(
                Fore.GREEN
                + (
                    "Archivo de configuración XML de desinstalación generado en: "
                    f"{sanitized_file_path}"
                )
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
        odt_manager = ODTManager(self.office_uninstall_dir)
        version_identifier = self.installation.name

        print(
            Fore.GREEN
            + (
                "Iniciando proceso de desinstalación para: "
                f"{version_identifier}"
            )
        )
        if not odt_manager.download_and_extract(version_identifier):
            print(Fore.RED + "Error al descargar y extraer ODT.")
            logging.error("Fallo la descarga o extracción del ODT.")
            return False

        office_uninstall_path = Path(self.office_uninstall_dir)
        self.setup_path = office_uninstall_path / "setup.exe"
        sanitized_uninstall_path = safe_log_path(office_uninstall_path)

        # Verifica que setup.exe esté presente antes de intentar
        # la desinstalación
        if not self.setup_path.exists():
            print(
                Fore.RED
                + (
                    "No se encontró 'setup.exe' en "
                    f"{sanitized_uninstall_path}."
                )
            )
            logging.error(
                f"'setup.exe' no encontrado en {sanitized_uninstall_path}"
            )
            return False

        try:
            # Genera el archivo XML de configuración necesario para la
            # desinstalación
            config_path = self._generar_configuracion_remocion()
        except Exception:
            print(
                Fore.RED
                + ("Error al generar el archivo de configuración XML.")
            )
            return False

        try:
            subprocess.run(
                [str(self.setup_path), "/configure", config_path],
                cwd=str(office_uninstall_path),
                capture_output=True,
                text=True,
                check=True,
            )

            logging.info(f"Office desinstalado: {self.installation.name}")
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
        if self.ejecutar_desinstalacion():
            return (
                f"{Fore.GREEN} Desinstalado exitosamente: "
                f"{self.installation.name}"
            )
        else:
            return f"{Fore.RED} Error al desinstalar: {self.installation.name}"


def run_uninstallers(
    installations: List[OfficeInstallation], uninstall_dir: Path
) -> None:
    """
    Ejecuta la desinstalación de múltiples instalaciones de Office utilizando
    ODT.

    Args:
        installations (List[OfficeInstallation]): Lista de instalaciones a
            desinstalar.
        uninstall_dir (Path): Directorio temporal para ODT.
    """
    # Ejecuta la desinstalación para cada instalación detectada
    for inst in installations:
        try:
            result = OfficeUninstaller(str(uninstall_dir), inst).execute()
            if result:
                print(result)
                logging.info(result)
        except Exception as e:
            logging.error(f"Error al desinstalar {inst}: {e}")
