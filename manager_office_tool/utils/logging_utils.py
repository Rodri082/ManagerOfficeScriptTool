import logging
from pathlib import Path


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
