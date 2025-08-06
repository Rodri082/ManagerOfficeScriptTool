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
        print(f"INFO - {message} ", end="", flush=True)
        respuesta = input().strip().lower()
        if respuesta in ("s", "sí", "si"):
            return True
        elif respuesta in ("n", "no"):
            return False
        else:
            msg = (
                "[CONSOLE] Respuesta no válida. Por favor, ingresa 'S' o 'N'."
            )
            logging.warning(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")


def ask_menu_option(valid_options: set[str], prompt: str) -> str:
    """
    Pregunta al usuario por una opción válida del menú dentro de valid_options.
    Repite la pregunta hasta que se ingrese una opción válida o se cancele.

    Args:
        valid_options (set[str]): Opciones válidas.
        prompt (str): Mensaje a mostrar al usuario.

    Returns:
        str: Opción válida o "cancel" si el usuario cancela.
    """
    options_str = "/".join(sorted(valid_options))
    while True:
        user_input = (
            input(f"{prompt} ({options_str}, c para cancelar): ")
            .strip()
            .lower()
        )
        if user_input == "c":
            return "cancel"
        if user_input in valid_options:
            return user_input
        logging.info(
            f"{Fore.YELLOW}"
            "Opción inválida. Intenta nuevamente."
            f"{Style.RESET_ALL}"
        )


def ask_single_valid_index(max_index: int) -> int | None:
    """
    Pregunta al usuario por un índice válido entre 1 y max_index.
    Permite cancelar con 'c' o entrada vacía.

    Args:
        max_index (int): Número máximo permitido.

    Returns:
        int | None: Índice válido o None si se cancela.
    """
    valid_indices = set(range(1, max_index + 1))
    prompt = "Ingrese el número de la versión a desinstalar"
    while True:
        user_input = (
            input(
                f"{Fore.LIGHTCYAN_EX}"
                f"{prompt} (1-{max_index}, c para cancelar): "
                f"{Style.RESET_ALL}"
            )
            .strip()
            .lower()
        )
        if user_input == "c":
            return None
        if user_input.isdigit():
            val = int(user_input)
            if val in valid_indices:
                return val
        logging.info(
            f"{Fore.YELLOW}"
            "Selección inválida. Intenta nuevamente."
            f"{Style.RESET_ALL}"
        )


def ask_multiple_valid_indices(max_index: int) -> list[int] | None:
    """
    Pregunta al usuario por múltiples índices válidos separados por comas.
    Permite cancelar con 'c' o entrada vacía.

    Args:
        max_index (int): Número máximo permitido.

    Returns:
        list[int] | None: Lista de índices válidos o None si se cancela.
    """
    valid_options = set(range(1, max_index + 1))
    prompt = [
        "Ingrese los números de las versiones a desinstalar "
        "(por ejemplo, 1,3,5)"
    ]
    while True:
        user_input = (
            input(
                f"{Fore.LIGHTCYAN_EX}"
                f"{prompt}, c para cancelar: "
                f"{Style.RESET_ALL}"
            )
            .strip()
            .lower()
        )
        if user_input == "c":
            return None
        parts = [p.strip() for p in user_input.split(",") if p.strip()]
        if all(p.isdigit() and int(p) in valid_options for p in parts):
            nums = [int(p) for p in parts]
            if len(nums) == len(set(nums)):  # sin duplicados
                return nums
        logging.info(
            f"{Fore.YELLOW}"
            "Selección inválida. Intenta nuevamente."
            f"{Style.RESET_ALL}"
        )
