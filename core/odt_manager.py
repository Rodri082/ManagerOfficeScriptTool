"""
odt_manager.py

Módulo encargado de gestionar la descarga y extracción
del Office Deployment Tool (ODT) utilizando Scrapy para obtener
la URL oficial y Requests para la descarga.
Incluye manejo de reintentos, descargas reanudables y extracción silenciosa.
"""

import json
import logging
import multiprocessing
import os
import random
import re
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import requests
from colorama import Fore
from requests.adapters import HTTPAdapter
from scrapy import Request, Spider
from scrapy.crawler import CrawlerProcess
from scrapy.utils.log import configure_logging
from tqdm import tqdm
from urllib3.util.retry import Retry
from utils import safe_log_path


def run_scrapy_spider_process(temp_file_path, download_id):
    """
    Ejecuta un spider de Scrapy en un proceso separado para obtener la URL
    de descarga del ODT.

    Args:
        temp_file_path (str): Ruta al archivo temporal donde se guardará el
            resultado en JSON.
        download_id (str): ID de descarga de Microsoft para la versión de
            ODT deseada.

    El resultado es un archivo JSON con las claves: url, name, size.
    """
    configure_logging({"LOG_ENABLED": False})
    logging.getLogger("scrapy").propagate = False
    logging.getLogger("scrapy").setLevel(logging.CRITICAL)
    logging.getLogger("twisted").setLevel(logging.CRITICAL)
    logging.getLogger("scrapy.utils.log").setLevel(logging.CRITICAL)

    class ODTSpider(Spider):
        name = "odt_spider"
        start_urls = [
            f"https://www.microsoft.com/en-us/download/details.aspx?id={download_id}"  # noqa: E501
        ]
        custom_settings = {"LOG_LEVEL": "ERROR"}

        def parse(self, response):
            for link in response.css("a::attr(href)").getall():
                if link.endswith(".exe") and "download.microsoft.com" in link:
                    yield Request(link, callback=self.parse_head)

        def parse_head(self, response):
            size = int(response.headers.get("Content-Length", 0))
            cd = response.headers.get("Content-Disposition", b"").decode()
            match = re.search(r'filename="?([^";]+)', cd)
            name = (
                match.group(1)
                if match
                else Path(urlparse(response.url).path).name
            )
            result = {"url": response.url, "name": name, "size": size}
            with open(temp_file_path, "w", encoding="utf-8") as f:
                json.dump(result, f)

    process = CrawlerProcess()
    process.crawl(ODTSpider)
    process.start()


class ODTManager:
    """
    Gestiona la descarga y extracción del Office Deployment Tool (ODT).

    Métodos:
        get_download_url(version_identifier): Obtiene la URL de descarga
            del ODT.
        download_and_extract(version_identifier, max_retries): Descarga y
            extrae el ODT.
    """

    # IDs oficiales de descarga de ODT desde Microsoft Download Center
    ODT_DOWNLOAD_ID_OFFICE_2013 = "36778"
    ODT_DOWNLOAD_ID_OFFICE_2016_AND_LATER = "49117"

    def __init__(self, office_install_dir: str) -> None:
        if not office_install_dir:
            raise ValueError(
                "El directorio de instalación no puede estar vacío"
            )
        self.office_dir = office_install_dir
        self.expected_name: Optional[str] = None
        self.expected_size: Optional[int] = None

    def _run_scrapy_spider(self, download_id: str) -> Optional[tuple]:
        """
        Lanza el proceso Scrapy y obtiene la información de descarga del ODT.

        Args:
            download_id (str): ID de descarga de Microsoft.

        Returns:
            tuple: (url, nombre de archivo, tamaño) o None si falla.
        """
        with tempfile.NamedTemporaryFile(
            mode="w+", suffix=".json", delete=False
        ) as tmp:
            temp_path = tmp.name

        p = multiprocessing.Process(
            target=run_scrapy_spider_process, args=(temp_path, download_id)
        )
        p.start()
        p.join()

        try:
            with open(temp_path, "r", encoding="utf-8") as f:
                result = json.load(f)
                return result["url"], result["name"], result["size"]
        except Exception as e:
            logging.error(f"Fallo al leer resultado del spider: {e}")
            return None

    def get_download_url(self, version_identifier: str) -> Optional[str]:
        version_id = (
            self.ODT_DOWNLOAD_ID_OFFICE_2013
            if "2013" in version_identifier
            else self.ODT_DOWNLOAD_ID_OFFICE_2016_AND_LATER
        )

        try:
            download_info = self._run_scrapy_spider(version_id)
            if download_info:
                self.expected_name = download_info[1]
                self.expected_size = download_info[2]
                return download_info[0]
        except Exception:
            logging.exception(
                "Error inesperado al obtener la URL de descarga."
            )
        return None

    def _is_valid_download(self, path: Path) -> bool:
        """
        Verifica si el archivo descargado es válido según nombre y
        tamaño esperados.

        Args:
            path (Path): Ruta al archivo descargado.

        Returns:
            bool: True si el archivo es válido, False en caso contrario.
        """
        return (
            path.exists()
            and self.expected_name == path.name
            and path.stat().st_size == self.expected_size
        )

    def download_and_extract(
        self, version_identifier: str, max_retries: int = 5
    ) -> bool:
        """
        Descarga y extrae el ODT para la versión indicada.

        Args:
            version_identifier (str): Identificador de la versión de Office.
            max_retries (int): Número máximo de reintentos de descarga.

        Returns:
            bool: True si la descarga y extracción fueron exitosas, False en
                caso contrario.
        """
        if not self.office_dir:
            logging.error("Directorio de instalación no definido")
            return False

        url = self.get_download_url(version_identifier)
        if not url:
            logging.error("No se pudo obtener la URL de descarga.")
            print(
                Fore.RED
                + (
                    "No se pudo obtener la URL de descarga. "
                    "Verifica tu conexión a internet o intenta más tarde."
                )
            )
            return False

        office_dir = Path(self.office_dir)
        office_dir.mkdir(parents=True, exist_ok=True)
        exe_file_name = Path(urlparse(url).path).name
        exe_path = office_dir / exe_file_name
        sanitized_path = safe_log_path(exe_path)
        sanitized_dir = safe_log_path(self.office_dir)

        session = requests.Session()
        # Configura la política de reintentos para la sesión HTTP
        retries = Retry(
            total=max_retries,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            respect_retry_after_header=True,
        )
        session.mount("https://", HTTPAdapter(max_retries=retries))

        # Usamos un archivo temporal para soportar descargas
        # interrumpidas y reanudables
        with tempfile.NamedTemporaryFile(dir=office_dir, delete=False) as tmp:
            tmp_path = Path(tmp.name)
        sanitized_tmp_path = safe_log_path(tmp_path)

        try:
            for attempt in range(1, max_retries + 1):
                # Si el archivo ya existe y es válido,
                # no es necesario descargarlo de nuevo
                if self._is_valid_download(exe_path):
                    logging.info(
                        f"Binario ya descargado y válido: {sanitized_path}"
                    )
                    break

                headers = {}
                downloaded = (
                    tmp_path.stat().st_size if tmp_path.exists() else 0
                )
                if downloaded:
                    head = session.head(url)
                    if head.headers.get("Accept-Ranges") == "bytes":
                        headers["Range"] = f"bytes={downloaded}-"

                logging.info(f"[{attempt}] Descargando {url}")
                try:
                    response = session.get(
                        url, stream=True, timeout=(3, 30), headers=headers
                    )
                    response.raise_for_status()
                    total_size = (
                        int(response.headers.get("Content-Length", 0))
                        + downloaded
                    )

                    with open(tmp_path, "ab") as f, tqdm(
                        total=total_size,
                        initial=downloaded,
                        unit="B",
                        unit_scale=True,
                        desc=(
                            f"Intento {attempt} - Descargando {exe_file_name}"
                        ),
                    ) as bar:
                        for chunk in response.iter_content(8192):
                            if chunk:
                                f.write(chunk)
                                f.flush()
                                os.fsync(f.fileno())
                                bar.update(len(chunk))

                    tmp_path.replace(exe_path)
                    print(
                        Fore.GREEN
                        + (
                            "Archivo descargado exitosamente en: "
                            f"{sanitized_path}"
                        )
                    )

                    if self._is_valid_download(exe_path):
                        break
                except requests.RequestException as e:
                    logging.warning(
                        f"Error en descarga (intento {attempt}): {e}"
                    )
                    print(
                        Fore.RED
                        + (
                            "Error de red durante la descarga. "
                            "Verifica tu conexión a internet "
                            "y permisos de red."
                        )
                    )
                    if attempt < max_retries:
                        delay = min(2**attempt + random.uniform(0, 1), 10)
                        logging.info(
                            f"Esperando {delay:.2f} segundos antes del"
                            "próximo intento..."
                        )
                        time.sleep(delay)
                        continue
                    else:
                        logging.error(
                            "No se completó la descarga tras varios intentos"
                        )
                        return False

            if self._is_valid_download(exe_path):
                try:
                    command = [
                        str(exe_path),
                        "/quiet",
                        f"/extract:{self.office_dir}",
                    ]
                    subprocess.run(
                        command,
                        cwd=str(self.office_dir),
                        capture_output=True,
                        text=True,
                        check=True,
                    )
                    logging.info(
                        f"ODT extraído correctamente en {sanitized_dir}"
                    )
                    return True
                except subprocess.CalledProcessError as e:
                    logging.error(f"Error al ejecutar el ODT:\n{e.stderr}")
                except PermissionError:
                    logging.error(
                        "Permiso denegado al intentar escribir en el disco."
                    )
                except OSError as e:
                    logging.error(f"Error del sistema de archivos: {e}")
                except Exception:
                    logging.exception(
                        "Error inesperado durante la extracción del ODT."
                    )
        finally:
            try:
                tmp_path.unlink()
            except OSError:
                logging.warning(
                    "No se pudo eliminar el archivo temporal: "
                    f"{sanitized_tmp_path}"
                )

        return False
