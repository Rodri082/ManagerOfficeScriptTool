from colorama import Fore


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
