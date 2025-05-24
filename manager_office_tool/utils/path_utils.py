import logging
import shutil
import sys
import tempfile
from pathlib import Path
from typing import List, Tuple, Union


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
    return Path(tempfile.mkdtemp(prefix=prefix))


def ensure_subfolder(parent: Path, name: str) -> Path:
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
    """
    eliminadas, errores = [], []
    for folder_path in folders:
        folder_path = Path(folder_path)
        sanitized_path = safe_log_path(folder_path)
        try:
            if folder_path.is_dir():
                shutil.rmtree(folder_path)
                eliminadas.append(str(folder_path))
                logging.info(f"Carpeta eliminada: {sanitized_path}")
            else:
                logging.debug(
                    f"La ruta no es una carpeta o no existe: {sanitized_path}"
                )
        except PermissionError:
            error_msg = (
                f"Permiso denegado al eliminar la carpeta: {sanitized_path}"
            )
            logging.error(error_msg)
            errores.append(error_msg)
        except FileNotFoundError:
            logging.warning(f"La carpeta ya no existe: {sanitized_path}")
        except OSError as e:
            error_msg = f"Error eliminando {sanitized_path}: {e}"
            logging.error(error_msg)
            errores.append(error_msg)
    return eliminadas, errores
