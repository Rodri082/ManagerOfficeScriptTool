"""
installer.py

Módulo encargado de ejecutar la instalación de Microsoft Office usando
setup.exe y configuration.xml.
Incluye validaciones y manejo de errores para una instalación robusta.
"""

import logging
import subprocess
from pathlib import Path

from colorama import Fore
from utils import safe_log_path


class OfficeInstaller:
    """
    Encargado de ejecutar el proceso de instalación de Microsoft Office
    utilizando el archivo setup.exe y configuration.xml generados previamente.

    Args:
        office_install_dir (str): Ruta donde se encuentra el setup.exe y el
            archivo de configuración XML.
    """

    def __init__(self, office_install_dir: str) -> None:
        """
        Inicializa la instancia con la ruta de instalación de Office.

        Args:
            office_install_dir (str): Carpeta que contiene los archivos
                necesarios para instalar Office.
        """
        self.office_install_dir = office_install_dir

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
            logging.error(
                "No se encontró 'setup.exe': %s", sanitized_setup_path
            )
            print(Fore.RED + "Error: 'setup.exe' no está disponible.")
            return

        if not config_path.exists():
            logging.error(
                "No se encontró 'configuration.xml': %s", sanitized_config_path
            )
            print(Fore.RED + "Error: 'configuration.xml' no está disponible.")
            return

        command = [str(setup_path), "/configure", str(config_path)]

        # Ejecuta el instalador en modo silencioso y captura la salida
        # para el log
        try:
            print(
                Fore.YELLOW
                + (
                    "Instalando Microsoft Office. "
                    "Por favor, no cierre esta ventana..."
                )
            )

            subprocess.run(
                command,
                cwd=str(office_dir),
                capture_output=True,
                text=True,
                check=True,
            )

            print(Fore.GREEN + "Instalación completada exitosamente.")
            logging.info("Instalación de Office finalizada correctamente.")

        except subprocess.CalledProcessError as e:
            # Maneja errores específicos del proceso de instalación
            print(Fore.RED + "La instalación falló.")
            logging.error(
                "setup.exe falló con código %d\nComando: %s\nStderr:\n%s",
                e.returncode,
                f"{sanitized_setup_path} /configure {sanitized_config_path}",
                e.stderr.strip(),
            )

        except PermissionError:
            # Informa si el usuario no tiene permisos suficientes
            logging.error("Permiso denegado al ejecutar setup.exe.")
            print(
                Fore.RED
                + (
                    "Permiso denegado al ejecutar la instalación. "
                    "Ejecuta el script como administrador."
                )
            )

        except OSError as e:
            # Maneja errores del sistema operativo
            # (por ejemplo, problemas de acceso a archivos)
            logging.error(
                f"Error del sistema operativo al ejecutar la instalación: {e}"
            )
            print(Fore.RED + "Error del sistema al iniciar la instalación")

        except Exception as e:
            # Captura cualquier otro error inesperado
            logging.exception(f"Error inesperado durante la instalación. {e}")
            print(
                Fore.RED
                + ("Ocurrió un error inesperado durante la instalación.")
            )
