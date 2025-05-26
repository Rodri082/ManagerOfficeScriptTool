"""
uninstaller.py

Módulo encargado de desinstalar instalaciones de Microsoft Office usando ODT.
Incluye generación dinámica del XML de desinstalación y manejo robusto
de errores.
"""

import logging
import subprocess
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
        installation: OfficeInstallation,
        odt_manager: Optional[ODTManager] = None,
    ) -> None:
        self.office_uninstall_dir = office_uninstall_dir
        self.installation = installation
        self.setup_path: Path | None = None
        self.odt_manager = odt_manager or ODTManager(office_uninstall_dir)

        logging.debug(
            "[*] Inicializando OfficeUninstaller para: "
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
        file_path = Path(self.odt_manager.office_dir) / "configuration.xml"
        sanitized_uninstall_path = safe_log_path(uninstall_dir)
        sanitized_file_path = safe_log_path(file_path)
        try:
            uninstall_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            msg = (
                "No se pudo crear el directorio de desinstalación: "
                f"{sanitized_uninstall_path}"
            )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            raise

        try:
            file_path.write_text(xml_content, encoding="utf-8")
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
        version_identifier = self.installation.name

        logging.info(
            f"{Fore.GREEN}"
            f"Iniciando proceso de desinstalación para: {version_identifier}"
            f"{Style.RESET_ALL}"
        )

        if not self.odt_manager.download_and_extract(version_identifier):
            msg = "Fallo la descarga o extracción del ODT."
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False

        office_uninstall_path = Path(self.odt_manager.office_dir)
        self.setup_path = office_uninstall_path / "setup.exe"
        sanitized_uninstall_path = safe_log_path(office_uninstall_path)

        # Verifica que setup.exe esté presente antes de intentar la
        # desinstalación
        if not self.setup_path.exists():
            msg = f"'setup.exe' no encontrado en {sanitized_uninstall_path}"
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False

        try:
            # Genera el archivo XML de configuración necesario para la
            # desinstalación
            config_path = self._generar_configuracion_remocion()
        except Exception:
            msg = "Error al generar el archivo de configuración XML."
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False

        try:
            logging.info(
                f"{Fore.CYAN}"
                "Ejecutando: setup.exe /configure configuration.xml"
                f"{Style.RESET_ALL}"
            )

            subprocess.run(
                [str(self.setup_path), "/configure", config_path],
                cwd=str(office_uninstall_path),
                capture_output=True,
                text=True,
                check=True,
            )

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
                return (
                    f"{Fore.GREEN}"
                    f"Desinstalado exitosamente: {self.installation.name}"
                    f"{Style.RESET_ALL}"
                )
            else:
                return f"{Fore.RED}Error al desinstalar: {self.installation.name}{Style.RESET_ALL}"  # noqa E501
        except Exception as e:
            logging.exception(
                f"Excepción en 'execute()' para {self.installation.name}: {e}"
            )
            return (
                f"{Fore.RED}"
                f"Falló inesperadamente: {self.installation.name}"
                f"{Style.RESET_ALL}"
            )


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
    total = len(installations)

    logging.info(
        f"{Fore.LIGHTCYAN_EX}"
        f"Iniciando desinstalación de {total} instalación(es)...\n"
        f"{Style.RESET_ALL}"
    )

    odt_managers: dict[str, ODTManager] = {}

    for idx, inst in enumerate(installations, start=1):
        logging.info(
            f"{Fore.LIGHTYELLOW_EX}"
            f"[{idx}/{total}] Desinstalando: {inst.name}"
            f"{Style.RESET_ALL}"
        )
        logging.info(f"{Fore.LIGHTCYAN_EX}{'-' * 110}{Style.RESET_ALL}")
        # Determina la familia según el nombre: 2013 o moderno (2016+)
        familia = "2013" if "2013" in inst.name else "modern"

        # Si es la primera vez que vemos esta familia, creamos y
        # descargamos ODT
        if familia not in odt_managers:
            odt_dir = uninstall_dir / f"OfficeODT_{familia}"
            manager = ODTManager(str(odt_dir))
            odt_managers[familia] = manager

            logging.info(
                f"{Fore.GREEN}"
                f"Preparando ODT para familia '{familia}'…"
                f"{Style.RESET_ALL}"
            )
            if not manager.download_and_extract(familia):
                msg = f"Error al preparar ODT para '{familia}'"
                logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
                return

        # Reutilizamos el manager cacheado para esta instalación
        uninstaller = OfficeUninstaller(
            str(uninstall_dir), inst, odt_managers[familia]
        )

        try:
            # Ejecuta la desinstalación con el manager cacheado
            result = uninstaller.execute()
            logging.info(result)
        except Exception as e:
            msg = f"Error al desinstalar {inst.name}: {e}"
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")

        # Separador visual entre instalaciones
        logging.info(f"{Fore.LIGHTCYAN_EX}{'-' * 110}{Style.RESET_ALL}")
