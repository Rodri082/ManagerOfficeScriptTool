"""
installer.py

Módulo encargado de ejecutar la instalación de Microsoft Office usando
setup.exe y configuration.xml.
Incluye validaciones y manejo de errores para una instalación robusta.
"""

import logging
import subprocess
from pathlib import Path

from colorama import Fore, Style
from manager_office_tool.utils import safe_log_path


class OfficeInstaller:
    """
    Ejecuta el proceso de instalación de Microsoft Office utilizando setup.exe
    y configuration.xml.

    Args:
        office_install_dir (str): Ruta donde se encuentra el setup.exe y
            el archivo de configuración XML.
        selected_version (str): Versión de Office a instalar
            (solo para mostrar en logs).
    """

    def __init__(
        self,
        office_install_dir: str,
        selected_version: str,
        selected_language_id: str,
    ) -> None:
        """
        Inicializa la instancia con la ruta de instalación de Office.

        Args:
            office_install_dir (str): Carpeta que contiene los archivos
                necesarios para instalar Office.
            selected_version (str): Versión de Office a instalar.
        """
        self.office_install_dir = office_install_dir
        self.selected_version = selected_version
        self.selected_language_id = selected_language_id

    def run_setup_commands(self) -> None:
        """
        Ejecuta el comando de instalación de Office utilizando setup.exe y
        configuration.xml.

        Valida la existencia de los archivos requeridos y maneja cualquier
        excepción durante la ejecución del proceso.
        """
        office_dir = Path(self.office_install_dir)
        setup_path = office_dir / "setup.exe"
        config_path = office_dir / "configuration.xml"
        sanitized_setup_path = safe_log_path(setup_path)
        sanitized_config_path = safe_log_path(config_path)

        # Verifica que setup.exe y configuration.xml existan antes de intentar
        # la instalación
        if not setup_path.exists():
            msg = (
                f"[CONSOLE] No se encontró 'setup.exe': {sanitized_setup_path}"
            )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return

        if not config_path.exists():
            msg = (
                "[CONSOLE] No se encontró 'configuration.xml': "
                f"{sanitized_config_path}"
            )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return

        command = [str(setup_path), "/configure", str(config_path)]

        try:
            logging.info(
                f"{Fore.YELLOW}"
                f"Instalando Microsoft {self.selected_version} "
                f"({self.selected_language_id}). "
                "Por favor, no cierre esta ventana..."
                f"{Style.RESET_ALL}"
            )

            subprocess.run(
                command,
                cwd=str(office_dir),
                capture_output=True,
                text=True,
                check=True,
            )

            logging.info(
                f"{Fore.GREEN}Instalación completada.{Style.RESET_ALL}"
            )

        except subprocess.CalledProcessError as e:
            msg = f"[CONSOLE] La instalación falló. {e}"
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            logging.error(
                "setup.exe falló con código %d\nComando: %s\nStderr:\n%s",
                e.returncode,
                f"{sanitized_setup_path} /configure {sanitized_config_path}",
                e.stderr.strip(),
            )

        except PermissionError:
            msg = (
                "[CONSOLE] Permiso denegado al ejecutar la instalación. "
                "Ejecuta como administrador."
            )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")

        except OSError as e:
            # Maneja errores del sistema operativo
            # (por ejemplo, problemas de acceso a archivos)
            if e.errno == 2:
                msg = (
                    "[CONSOLE] No se encontró el archivo o "
                    "directorio especificado. "
                    f"Verifica la ruta: {sanitized_setup_path}"
                )
            elif e.errno == 13:
                msg = (
                    "[CONSOLE] Permiso denegado al acceder a un "
                    "archivo o directorio. "
                    f"Verifica los permisos: {sanitized_setup_path}"
                )
            else:
                msg = (
                    "[CONSOLE] Error del sistema al iniciar la instalación "
                    f"{e}"
                )
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")

        except Exception as e:
            logging.error("Error inesperado durante la instalación: %s", e)
