"""
console_utils.py

Utilidades para interacción básica con el usuario en consola.
Incluye funciones para preguntas de sí/no con validación robusta.
"""

import logging

from colorama import Fore, Style


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
        respuesta = input(f"INFO - {message}").strip().lower()
        if respuesta in ("s", "sí", "si"):
            return True
        elif respuesta in ("n", "no"):
            return False
        else:
            msg = "[CONSOLE] Respuesta no válida. Por favor ingresa 'S' o 'N'."
            logging.warning(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")
