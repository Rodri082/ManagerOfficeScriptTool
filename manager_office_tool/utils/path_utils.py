"""
path_utils.py

Utilidades para manejo seguro de rutas, carpetas temporales y sanitización de
logs. Incluye funciones para obtener rutas de datos, crear carpetas temporales,
asegurar subcarpetas, y limpiar rutas sensibles antes de mostrarlas en logs.
"""

import logging
import shutil
import sys
import tempfile
from pathlib import Path
from typing import List, Tuple, Union

from colorama import Fore, Style


def get_data_path(filename: str) -> Path:
    """
    Devuelve la ruta absoluta a un archivo de datos (como config.yaml),
    compatible tanto en desarrollo como empaquetado con Nuitka.
    """
    if getattr(sys, "frozen", False):
        # Ejecutando como binario (Nuitka, PyInstaller, etc.)
        base_path = Path(sys.executable).parent
    else:
        # Ejecutando como script normal
        base_path = Path(__file__).resolve().parents[2]  # Raíz del proyecto
    path = base_path / filename

    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {path}")
    return path


def get_temp_dir(prefix="ManagerOfficeScriptTool_") -> Path:
    """
    Crea y retorna un directorio temporal único para la sesión.
    """
    return Path(tempfile.mkdtemp(prefix=prefix))


def ensure_subfolder(parent: Path, name: str) -> Path:
    """
    Garantiza la existencia de una subcarpeta dentro de un directorio dado.

    Args:
        parent (Path): Carpeta padre.
        name (str): Nombre de la subcarpeta.

    Returns:
        Path: Ruta a la subcarpeta creada o existente.
    """
    folder = parent / name
    folder.mkdir(parents=True, exist_ok=True)
    return folder


# Sanitiza rutas para evitar exponer información sensible en los logs
def safe_log_path(path: Union[str, Path]) -> str:
    """
    Convierte rutas sensibles a una forma anonimizada para el log.
    Reemplaza la carpeta del usuario con %USERPROFILE%.
    """
    try:
        path_obj = Path(path)
        return str(path_obj).replace(str(Path.home()), "%USERPROFILE%")
    except Exception:
        return str(path)


def safe_log_registry_key(reg_path: str) -> str:
    """
    Sanitiza rutas de registro para ocultar posibles datos sensibles.
    """
    return reg_path.replace(
        "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\", "HKLM\\...\\"
    )


def clean_folders(folders: List[Path]) -> Tuple[List[str], List[str]]:
    """
    Intenta eliminar las carpetas indicadas.
    Devuelve dos listas: eliminadas y errores.
    Además, registra logs para cada acción relevante.

    Args:
        folders (List[Path]): Lista de carpetas a eliminar.

    Returns:
        Tuple[List[str], List[str]]: (eliminadas, errores)
    """
    eliminadas, errores = [], []
    for folder_path in folders:
        folder_path = Path(folder_path)
        sanitized_path = safe_log_path(folder_path)
        try:
            if folder_path.is_dir():
                shutil.rmtree(folder_path)
                eliminadas.append(str(folder_path))
                logging.info(
                    f"{Fore.GREEN}"
                    f"Carpeta eliminada: {sanitized_path}"
                    f"{Style.RESET_ALL}"
                )
            else:
                logging.debug(
                    f"La ruta no es una carpeta o no existe: {sanitized_path}"
                )
        except PermissionError:
            msg = f"Permiso denegado al eliminar la carpeta: {sanitized_path}"
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            errores.append(msg)
        except FileNotFoundError:
            msg = f"La carpeta ya no existe: {sanitized_path}"
            logging.warning(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")
        except OSError as e:
            msg = f"Error eliminando {sanitized_path}: {e}"
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            errores.append(msg)
    return eliminadas, errores
