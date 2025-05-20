"""
utils.py

Funciones y utilidades generales para el manejo de rutas, logs, diálogos y
carpetas temporales en ManagerOfficeScriptTool.
"""

import logging
import shutil
import sys
import tempfile
import tkinter as tk
import tkinter.messagebox as messagebox
from pathlib import Path
from typing import Union

from colorama import Fore


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
        base_path = Path(__file__).parent
    return base_path / filename


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


def init_logging(logs_path: str):
    """
    Inicializa el sistema de logging creando el directorio y configurando el
    archivo de log.
    Usando pathlib para un manejo de rutas más pythonic.
    """
    path = Path(logs_path)
    path.mkdir(parents=True, exist_ok=True)

    log_file = path / "application.log"

    logging.basicConfig(
        filename=str(log_file),
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )


def ask_yes_no(message: str) -> bool:
    """
    Pregunta al usuario en consola y retorna True si la respuesta es
    afirmativa.

    Args:
        message (str): Mensaje a mostrar al usuario.

    Returns:
        bool: True si la respuesta es afirmativa, False en caso contrario.
    """
    while True:
        respuesta = input(f"{message} (S/N): ").strip().lower()
        if respuesta in ("s", "sí", "si"):
            return True
        elif respuesta in ("n", "no"):
            return False
        else:
            print(
                Fore.YELLOW
                + ("Respuesta no válida. Por favor ingresa 'S' o 'N'.")
            )


def center_window(win, width=300, height=200):
    win.update_idletasks()
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (width // 2)
    y = (hs // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")


def clean_temp_folders(folders: list[Path]) -> None:
    """
    Elimina las carpetas temporales generadas por el script si el usuario
    lo confirma.
    Aplica sanitización a las rutas antes de mostrar o registrar mensajes.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if not messagebox.askyesno(
        "Eliminar archivos temporales",
        "¿Deseas eliminar las carpetas temporales creadas por el script?",
    ):
        messagebox.showinfo(
            "Cancelado", "No se eliminaron las carpetas temporales."
        )
        root.destroy()
        return

    eliminadas, errores = [], []
    for folder_path in folders:
        folder_path = Path(folder_path)
        sanitized_folder = safe_log_path(folder_path)
        try:
            if folder_path.is_dir():
                shutil.rmtree(folder_path)
                eliminadas.append(sanitized_folder)
                logging.info(f"Carpeta eliminada: {sanitized_folder}")
        except PermissionError:
            logging.error(
                f"Permiso denegado al eliminar la carpeta: {sanitized_folder}"
            )
            errores.append(f"Permiso denegado: {sanitized_folder}")
        except FileNotFoundError:
            logging.warning(f"La carpeta ya no existe: {sanitized_folder}")
        except OSError as e:
            logging.exception(f"Error eliminando {sanitized_folder}: {e}")
            errores.append(f"{sanitized_folder}: {e}")

    if errores:
        messagebox.showwarning(
            "Error",
            "Ocurrieron errores al eliminar algunas carpetas:\n\n"
            + "\n".join(errores),
        )
    elif eliminadas:
        messagebox.showinfo(
            "Archivos eliminados",
            "Las carpetas temporales han sido eliminadas correctamente.",
        )
    else:
        messagebox.showinfo(
            "Nada que eliminar",
            "No se encontraron carpetas temporales para eliminar.",
        )

    root.destroy()
