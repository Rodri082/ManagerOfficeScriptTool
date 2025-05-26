"""
logging_utils.py

Utilidades para inicializar y configurar el logging de la aplicación.
Incluye un filtro personalizado para mostrar mensajes INFO y
advertencias/errores marcados en consola, y guarda todos los logs
en archivo.
"""

import logging
from pathlib import Path


class InfoAndConsoleMarkFilter(logging.Filter):
    """
    Filtro para el handler de consola:
    - Permite todos los mensajes INFO.
    - Para WARNING y superiores, solo los que contienen el
        marcador '[CONSOLE]'.
    """

    def __init__(self, marker: str = "[CONSOLE]"):
        super().__init__()
        self.marker = marker

    def filter(self, record: logging.LogRecord) -> bool:
        if record.levelno == logging.INFO:
            return True
        if record.levelno >= logging.WARNING:
            return self.marker in record.getMessage()
        return False


class ConsoleFormatter(logging.Formatter):
    """
    Formatter que elimina el marcador '[CONSOLE]' solo para la consola.
    """

    def format(self, record: logging.LogRecord) -> str:
        msg = record.getMessage()
        if "[CONSOLE]" in msg:
            msg = msg.replace("[CONSOLE]", "").strip()
            record.msg = msg
            record.args = ()
        return super().format(record)


def init_logging(logs_path: str) -> None:
    """
    Inicializa el sistema de logging:
    - Crea el directorio de logs si no existe.
    - Guarda todos los mensajes en 'application.log'.
    - Muestra solo INFO y advertencias/errores marcados en consola.

    Args:
        logs_path (str): Ruta donde se almacenarán los logs.
    """
    path = Path(logs_path)
    path.mkdir(parents=True, exist_ok=True)

    log_file = path / "application.log"

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s"
    )
    file_handler.setFormatter(file_formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.addFilter(InfoAndConsoleMarkFilter())
    console_formatter = ConsoleFormatter("%(levelname)s - %(message)s")
    console_handler.setFormatter(console_formatter)

    logger.handlers.clear()
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.propagate = False
