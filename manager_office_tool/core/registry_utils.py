"""
registry_utils.py

Módulo para leer claves y valores del registro de Windows de forma
segura y eficiente. Incluye manejo de errores y caché para optimizar el acceso.
"""

import logging
import platform
import winreg
from typing import Dict, List, Tuple

from manager_office_tool.utils import safe_log_registry_key


class RegistryReader:
    """
    Clase para leer claves y valores del registro de Windows de forma
    eficiente.

    Utiliza un caché interno (_cache) para evitar lecturas repetidas de la
    misma clave.
    """

    def __init__(self) -> None:
        self._cache: Dict[Tuple[str, str], str] = {}

    def get_registry_keys(self, key: str) -> List[str]:
        """
        Devuelve las subclaves de una clave del registro.

        Args:
            key (str): Ruta de la clave.

        Returns:
            List[str]: Subclaves encontradas o lista vacía si hay error.
        """
        root_key = winreg.HKEY_LOCAL_MACHINE
        access_flag = winreg.KEY_READ | (
            winreg.KEY_WOW64_64KEY if platform.machine().endswith("64") else 0
        )

        subkeys: List[str] = []
        sanitized_key = safe_log_registry_key(key)
        try:
            with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
                index = 0
                while True:
                    try:
                        subkey = winreg.EnumKey(key_handle, index)
                        subkeys.append(subkey)
                        index += 1
                    except OSError:
                        break
        except FileNotFoundError:
            logging.warning(
                f"Clave del registro no encontrada: '{sanitized_key}'"
            )
        except PermissionError:
            logging.error(
                f"Permiso denegado al acceder a la clave: '{sanitized_key}'"
            )
        except OSError as e:
            logging.error(f"Error OS al abrir clave '{sanitized_key}': {e}")
        except Exception as e:
            logging.exception(
                "Excepción inesperada al acceder a la clave "
                f"'{sanitized_key}': {e}"
            )

        return subkeys

    def get_registry_value(self, key: str, value_name: str) -> str:
        """
        Obtiene un valor de una clave del registro (caché integrada).

        Args:
            key (str): Ruta de la clave.
            value_name (str): Nombre del valor.

        Returns:
            str: Valor encontrado o cadena vacía.
        """
        cache_key = (key, value_name)
        if cache_key in self._cache:
            return self._cache[cache_key]

        root_key = winreg.HKEY_LOCAL_MACHINE
        access_flag = winreg.KEY_READ | (
            winreg.KEY_WOW64_64KEY if platform.machine().endswith("64") else 0
        )

        sanitized_key = safe_log_registry_key(key)

        try:
            with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
                try:
                    value, _ = winreg.QueryValueEx(key_handle, value_name)
                    self._cache[cache_key] = value
                    return value
                except FileNotFoundError:
                    logging.warning(
                        f"Valor '{value_name}' no encontrado en clave: "
                        f"'{sanitized_key}'"
                    )
                except PermissionError:
                    logging.error(
                        f"Permiso denegado para leer '{value_name}' en clave: "
                        f"'{sanitized_key}'"
                    )
                except OSError as e:
                    logging.error(
                        f"Error OS al leer '{value_name}' en clave "
                        f"'{sanitized_key}': {e}"
                    )
        except FileNotFoundError:
            logging.warning(
                f"Clave no encontrada al buscar valor '{value_name}': "
                f"'{sanitized_key}'"
            )
        except PermissionError:
            logging.error(
                f"Permiso denegado al abrir clave '{sanitized_key}' para leer "
                f"'{value_name}'"
            )
        except Exception as e:
            logging.exception(
                f"Excepción inesperada al leer '{value_name}' en clave "
                f"'{sanitized_key}': {e}"
            )

        return ""
