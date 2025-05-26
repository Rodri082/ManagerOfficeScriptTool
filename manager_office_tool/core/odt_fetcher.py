"""
odt_fetcher.py

Funciones y clases para obtener información de descarga del
Office Deployment Tool (ODT) desde la web oficial de Microsoft.
Utiliza PySide6 para renderizar páginas dinámicas y requests para
obtener metadatos del archivo ejecutable.

Incluye:
- Renderizado de páginas web para extraer enlaces de descarga.
- Obtención de nombre y tamaño del archivo ODT.
- Cacheo de resultados para eficiencia.
- Solo debe usarse con IDs oficiales de Microsoft para evitar
  riesgos de seguridad.
"""

import logging
import re
from pathlib import Path
from typing import Dict, Optional
from urllib.parse import urlparse

import requests
from PySide6.QtCore import QEventLoop, QUrl
from PySide6.QtWebEngineCore import QWebEnginePage
from PySide6.QtWidgets import QApplication

# Cache local para evitar renderizar la misma página varias veces
_FETCH_CACHE: Dict[str, Dict[str, object]] = {}


class WebPage(QWebEnginePage):
    """
    Página web renderizada con PySide6 para obtener el HTML final
    de páginas dinámicas.
    """

    def __init__(self, url: str) -> None:
        self.app = QApplication.instance() or QApplication([])
        super().__init__()
        self.html: Optional[str] = None
        logging.debug(f"[*] WebPage: cargando URL con QWebEngine: {url}")
        self.loadFinished.connect(self._on_load_finished)
        self.load(QUrl(url))
        self.loop = QEventLoop()
        self.loop.exec_()

    def javaScriptConsoleMessage(
        self, level, message, lineNumber, sourceID, *args
    ) -> None:
        location = f"{sourceID or 'desconocido'}:{lineNumber}"
        if level == 2:
            logging.error(f"[JS ERROR] {location} - {message}")
        elif level == 1:
            logging.warning(f"[JS WARNING] {location} - {message}")
        else:
            logging.info(f"[JS INFO] {message}")

    def _on_load_finished(self, ok: bool) -> None:
        if not ok:
            logging.error("[!] WebPage: fallo al cargar la página.")
            self.loop.quit()
            return
        logging.debug("[*] WebPage: carga completada, extrayendo HTML...")
        self.toHtml(self._store_html)

    def _store_html(self, data: Optional[str]) -> None:
        self.html = data
        logging.debug(
            "[*] WebPage: HTML almacenado, saliendo del bucle de eventos."
        )
        self.loop.quit()


def get_rendered_html(url: str) -> Optional[str]:
    """
    Renderiza una página web usando PySide6 y devuelve el HTML final.
    Solo debe usarse con URLs confiables.
    """
    try:
        page = WebPage(url)
        return page.html
    except Exception as e:
        logging.exception(f"[!] Error al renderizar HTML: {e}")
        return None


def parse_head(url: str) -> Dict[str, object]:
    """
    Realiza una solicitud HEAD a la URL para obtener cabeceras HTTP,
    extrayendo el nombre y tamaño del archivo. Valida que la URL final
    pertenezca a un dominio seguro.
    """
    logging.debug(f"[*] parse_head: realizando HEAD a {url}")
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/115.0 Safari/537.36"
        )
    }
    response = requests.head(
        url, headers=headers, timeout=10, allow_redirects=True
    )
    response.raise_for_status()

    # Validación de dominio seguro tras redirecciones
    final_url = response.url
    parsed = urlparse(final_url)
    if parsed.hostname != "download.microsoft.com":
        logging.error(
            f"[!] Redirección a dominio no permitido: {parsed.hostname}"
        )
        raise ValueError("Redirección a dominio no permitido")

    size_str = response.headers.get("Content-Length", "0")
    size = (
        int(size_str)
        if isinstance(size_str, str) and size_str.isdigit()
        else 0
    )
    cd = response.headers.get("Content-Disposition", "")

    if isinstance(cd, bytes):
        cd = cd.decode("utf-8")

    match = re.search(r'filename="?([^";]+)', cd)
    name = match.group(1) if match else Path(urlparse(url).path).name
    logging.debug(f"[*] parse_head: obtenido name={name}, size={size}")
    return {"url": final_url, "name": name, "size": size}


def fetch_odt_download_info(download_id: str) -> Optional[Dict[str, object]]:
    """
    Obtiene la URL, nombre y tamaño del ODT oficial de Microsoft usando PySide6
    para renderizar la página y extraer la URL del archivo ejecutable.
    Implementa caching para evitar peticiones repetidas.

    Args:
        download_id (str): ID de descarga de Microsoft.

    Returns:
        dict con las claves 'url', 'name' y 'size', o None si falla.
    """
    if download_id in _FETCH_CACHE:
        logging.debug(
            f"[*] fetch_odt_download_info: usando cache para id={download_id}"
        )
        return _FETCH_CACHE[download_id]

    summary_url = f"https://www.microsoft.com/en-us/download/details.aspx?id={download_id}"  # noqa: E501
    logging.debug(
        f"[*] fetch_odt_download_info: cargando página {summary_url}"
    )
    html = get_rendered_html(summary_url)
    if not html:
        logging.error(
            f"[!] fetch_odt_download_info: HTML vacío para id={download_id}"
        )
        return None

    # Extrae enlaces .exe del HTML renderizado
    pattern = r"https://download\.microsoft\.com/[^\"]+\.exe"
    matches = re.findall(pattern, html)
    enlaces_unicos = list(dict.fromkeys(matches))  # preserva el orden
    logging.debug(
        "[*] fetch_odt_download_info: "
        f"encontrados {len(enlaces_unicos)} enlaces .exe"
    )

    if not enlaces_unicos:
        logging.error("[!] No se encontró ningún enlace .exe en la página.")
        return None

    enlace = enlaces_unicos[0]
    info = parse_head(enlace)
    if not info or "error" in info:
        logging.error(
            "[!] Error en parse_head: "
            f"{info.get('error') if info else 'Desconocido'}"
        )
        return None

    _FETCH_CACHE[download_id] = info
    logging.debug(
        "[*] fetch_odt_download_info: "
        f"info almacenada para id={download_id} -> {info}"
    )
    return info
