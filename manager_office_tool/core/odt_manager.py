"""
odt_manager.py

Gestión de descarga y extracción del Office Deployment Tool (ODT) desde la web
oficial de Microsoft.
Incluye manejo de reintentos, descargas reanudables y extracción silenciosa.
"""

import logging
import os
import random
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import requests
from colorama import Fore, Style
from manager_office_tool.utils import safe_log_path
from requests.adapters import HTTPAdapter
from tqdm import tqdm
from urllib3.util.retry import Retry

from .odt_fetcher import fetch_odt_download_info


class ODTManager:
    """
    Gestiona la descarga y extracción del Office Deployment Tool (ODT).

    Seguridad:
    - El directorio de instalación debe ser generado por la lógica interna
        del programa.
    - No se recomienda aceptar rutas arbitrarias de entrada de usuario para
        evitar riesgos de path-injection.
    """

    ODT_DOWNLOAD_ID_OFFICE_2013 = "36778"
    ODT_DOWNLOAD_ID_OFFICE_2016_AND_LATER = "49117"

    def __init__(self, office_install_dir: str) -> None:
        if not office_install_dir:
            raise ValueError(
                "El directorio de instalación no puede estar vacío"
            )
        self.office_dir = Path(office_install_dir)
        self.expected_name: Optional[str] = None
        self.expected_size: Optional[int] = None

    def _run_odt_fetcher(
        self, download_id: str
    ) -> Optional[tuple[str, str, int]]:
        """
        Ejecuta fetch_odt_download_info y retorna
        (url, nombre, tamaño) del ODT.
        """
        try:
            result = fetch_odt_download_info(download_id)
            if result:
                url = str(result["url"])
                name = str(result["name"])
                size_raw = result["size"]
                if isinstance(size_raw, int):
                    size = size_raw
                elif isinstance(size_raw, str) and size_raw.isdigit():
                    size = int(size_raw)
                else:
                    size = 0
                return url, name, size
        except Exception as e:
            msg = f"[CONSOLE] Error al obtener info de descarga : {e}"
            logging.exception(f"{Fore.RED}{msg}{Style.RESET_ALL}")
        return None

    def get_download_url(self, version_identifier: str) -> Optional[str]:
        """
        Obtiene la URL de descarga del ODT según la versión de Office.
        """
        logging.debug(
            f"[*] ODTManager: obteniendo URL para '{version_identifier}'…"
        )
        version_id = (
            self.ODT_DOWNLOAD_ID_OFFICE_2013
            if "2013" in version_identifier
            else self.ODT_DOWNLOAD_ID_OFFICE_2016_AND_LATER
        )
        try:
            download_info = self._run_odt_fetcher(version_id)
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
        """
        if not path.exists():
            return False
        valid = (
            self.expected_name == path.name
            and path.stat().st_size == self.expected_size
        )
        if not valid:
            logging.debug(
                f"Archivo no válido: esperado '{self.expected_name}' "
                f"({self.expected_size} bytes), recibido '{path.name}' "
                f"({path.stat().st_size} bytes)"
            )
        return valid

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
            msg = "[CONSOLE] Directorio de instalación no definido."
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False
        logging.debug(
            "[*] ODTManager: download_and_extract iniciada para "
            f"{version_identifier}"
        )
        url = self.get_download_url(version_identifier)
        if not url:
            msg = "[CONSOLE] No se pudo obtener la URL de descarga."
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return False

        office_dir = Path(self.office_dir)
        office_dir.mkdir(parents=True, exist_ok=True)
        exe_file_name = Path(urlparse(url).path).name
        exe_path = office_dir / exe_file_name
        sanitized_path = safe_log_path(exe_path)
        sanitized_dir = safe_log_path(self.office_dir)

        session = requests.Session()
        retries = Retry(
            total=max_retries,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            respect_retry_after_header=True,
        )
        session.mount("https://", HTTPAdapter(max_retries=retries))

        with tempfile.NamedTemporaryFile(dir=office_dir, delete=False) as tmp:
            tmp_path = Path(tmp.name)
        sanitized_tmp_path = safe_log_path(tmp_path)

        try:
            for attempt in range(1, max_retries + 1):
                if self._is_valid_download(exe_path):
                    logging.debug(
                        f"Binario ya descargado y válido: {sanitized_path}"
                    )
                    break
                headers = {}
                downloaded = (
                    tmp_path.stat().st_size if tmp_path.exists() else 0
                )
                if downloaded:
                    try:
                        head = session.head(url, timeout=(3, 10))
                        accept_ranges = head.headers.get(
                            "Accept-Ranges", ""
                        ).lower()
                        if accept_ranges == "bytes":
                            headers["Range"] = f"bytes={downloaded}-"
                        else:
                            logging.warning(
                                f"{Fore.YELLOW}"
                                "El servidor no permite reanudar descargas "
                                "(sin Accept-Ranges). "
                                "Eliminando archivo temporal para comenzar "
                                "desde cero."
                                f"{Style.RESET_ALL}"
                            )
                            try:
                                tmp_path.unlink()
                                downloaded = 0
                            except OSError:
                                logging.warning(
                                    f"{Fore.YELLOW}"
                                    "No se pudo eliminar el archivo temporal: "
                                    f"{safe_log_path(tmp_path)}"
                                    f"{Style.RESET_ALL}"
                                )
                    except requests.RequestException as e:
                        logging.warning(
                            f"{Fore.YELLOW}"
                            "Error al verificar soporte de reanudación: "
                            f"{e}{Style.RESET_ALL}"
                        )

                logging.info(
                    f"{Fore.GREEN}"
                    f"Descargando {exe_file_name}"
                    f"{Style.RESET_ALL}"
                )
                logging.info(f"{Fore.GREEN}URL: {url}{Style.RESET_ALL}")
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
                        desc=(f"INFO - [{attempt}/{max_retries}]"),
                    ) as bar:
                        for chunk in response.iter_content(8192):
                            if chunk:
                                f.write(chunk)
                                f.flush()
                                os.fsync(f.fileno())
                                bar.update(len(chunk))

                    tmp_path.replace(exe_path)
                    logging.debug(
                        "[*] Archivo descargado exitosamente en: "
                        f"{sanitized_path}"
                    )
                    logging.debug(
                        f"Tamaño esperado: {self.expected_size} bytes, "
                        f"tamaño real: {exe_path.stat().st_size} bytes"
                    )

                    if self._is_valid_download(exe_path):
                        break
                except requests.RequestException as e:
                    logging.warning(
                        f"Error en descarga (intento {attempt}): {e}"
                    )
                    if attempt < max_retries:
                        delay = min(2**attempt + random.uniform(0, 1), 10)
                        logging.info(
                            f"{Fore.YELLOW}"
                            f"Esperando {delay:.2f} "
                            "segundos antes del próximo intento..."
                            f"{Style.RESET_ALL}"
                        )
                        time.sleep(delay)
                        continue
                    else:
                        msg = (
                            "[CONSOLE] Descarga fallida tras "
                            f"{max_retries} intentos."
                        )
                        logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
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
                        f"{Fore.GREEN}"
                        f"ODT extraído: {sanitized_dir}"
                        f"{Style.RESET_ALL}"
                    )
                    return True
                except subprocess.CalledProcessError as e:
                    logging.error(f"Error al ejecutar el ODT:\n{e.stderr}")
                    msg = "[CONSOLE] Error al ejecutar el ODT."
                    logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
                except PermissionError:
                    msg = (
                        "[CONSOLE] Permiso denegado: "
                        "Este proceso requiere ejecutar con "
                        "privilegios de administrador."
                    )
                    logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
                except OSError as e:
                    logging.error(f"Error del sistema de archivos: {e}")
                except Exception:
                    logging.exception(
                        "Error inesperado durante la extracción del ODT."
                    )
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    msg = (
                        "[CONSOLE] No se pudo eliminar el archivo temporal: "
                        f"{sanitized_tmp_path}"
                    )
                    logging.warning(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")

        return False
