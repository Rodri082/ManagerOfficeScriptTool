"""
odt_fetcher.py

Módulo para obtener información de descarga del Office Deployment Tool (ODT)
desde la página oficial de Microsoft.

Este módulo utiliza la biblioteca `requests` para recuperar el HTML de la
página y obtener metadatos del archivo ejecutable (.exe) de instalación.

Características:
- Descarga de páginas web estáticas y extracción de enlaces .exe válidos.
- Obtención de nombre y tamaño del archivo ODT mediante cabeceras HTTP.
- Caché LRU en memoria para evitar solicitudes repetidas innecesarias.
- Validación estricta del dominio (solo dominios seguros y permitidos).
- Solo debe usarse con IDs oficiales de Microsoft para minimizar riesgos
de seguridad (evita IDs arbitrarios o maliciosos).
"""

import ipaddress
import logging
import re
import string
from collections import OrderedDict
from pathlib import Path
from typing import Dict, Optional
from urllib.parse import urlparse

import requests
from lxml import html
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

ALLOWED_DOMAINS = {"download.microsoft.com"}
ALLOWED_DOMAINS_SUBDOMAINS = {"download.microsoft.com"}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/115.0 Safari/537.36"
    )
}

MAX_HTML_SIZE = 2 * 1024 * 1024  # 2 MB
MAX_CACHE_SIZE = 100  # límite arbitrario para evitar crecimiento indefinido

# Caché LRU para evitar descargar la misma página varias veces, con límite
_FETCH_CACHE: OrderedDict[str, Dict[str, object]] = OrderedDict()


def domain_allowed(hostname: Optional[str]) -> bool:
    """
    Verifica que el hostname sea válido, pertenezca a dominios permitidos,
    y no sea una IP ni tenga caracteres homoglifos.
    Permite subdominios explícitos.
    """
    if not hostname:
        return False

    hostname = hostname.lower()

    # Validar que no sea IP (IPv4 o IPv6)
    try:
        ipaddress.ip_address(hostname)
        logging.debug(f"[!] hostname es IP, no permitido: {hostname}")
        return False
    except ValueError:
        # No es IP, sigue validando hostname
        pass

    # Normalizar dominio con idna para evitar homoglifos y unicode confusos
    try:
        hostname_idna = hostname.encode("idna").decode("ascii")
    except Exception as e:
        logging.error(
            f"[!] Error al codificar hostname a idna: {hostname} - {e}"
        )
        return False

    # Validar que el dominio sea exactamente uno permitido o un
    # subdominio autorizado
    if hostname_idna in ALLOWED_DOMAINS:
        return True

    # Permitir subdominios de ALLOWED_DOMAINS_SUBDOMAINS
    for domain in ALLOWED_DOMAINS_SUBDOMAINS:
        if hostname_idna.endswith("." + domain):
            return True

    logging.debug(f"[!] Dominio no permitido: {hostname_idna}")
    return False


def sanitize_filename(name: str) -> str:
    """
    Sanitiza el nombre del archivo para evitar path traversal y caracteres
    no válidos. Limita longitud a 255 caracteres y asegura que sea basename
    (sin separadores). Si queda vacío, retorna un nombre por defecto seguro.
    """
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    filtered = "".join(c for c in name if c in valid_chars).strip()

    # Evitar nombres con separadores de ruta
    filtered = filtered.replace("/", "").replace("\\", "")

    if not filtered:
        filtered = "file.exe"

    # Limitar longitud a 255 (suficiente para la mayoría de sistemas)
    if len(filtered) > 255:
        filtered = filtered[:255]

    return filtered


def get_requests_session() -> requests.Session:
    """
    Crea una sesión de requests con reintentos para mejorar robustez.
    """
    session = requests.Session()
    retries = Retry(
        total=3,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"],
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def get_page_html(url: str) -> Optional[str]:
    """
    Descarga el HTML de una página confiable (Microsoft) usando requests,
    con límite máximo de tamaño para evitar ataques DoS.
    """
    session = get_requests_session()

    try:
        with session.get(
            url, headers=HEADERS, timeout=15, stream=True, verify=True
        ) as response:
            response.raise_for_status()

            content = []
            total = 0
            for chunk in response.iter_content(chunk_size=8192):
                total += len(chunk)
                if total > MAX_HTML_SIZE:
                    logging.error(
                        "[!] HTML demasiado grande, abortando descarga"
                    )
                    return None
                content.append(chunk)

            # Decodificar con manejo de errores
            encoding = (
                response.encoding
                if response.encoding
                else response.apparent_encoding
            )
            html_content = b"".join(content).decode(encoding, errors="replace")
            return html_content

    except requests.RequestException as e:
        logging.exception(f"[!] Error al obtener HTML: {e}")
        return None


def parse_head(url: str) -> Dict[str, object]:
    """
    Realiza una solicitud HEAD a la URL para obtener cabeceras HTTP,
    extrayendo el nombre y tamaño del archivo. Valida que la URL final
    pertenezca a un dominio seguro y use HTTPS.
    """
    logging.debug(f"[*] parse_head: realizando HEAD a {url}")
    session = get_requests_session()

    try:
        response = session.head(
            url, headers=HEADERS, timeout=10, allow_redirects=True, verify=True
        )
        response.raise_for_status()
    except requests.RequestException as e:
        logging.error(f"[!] Error en HEAD request: {e}")
        return {}

    final_url = response.url
    parsed = urlparse(final_url)

    if parsed.scheme != "https":
        logging.error(f"[!] Esquema no permitido: {parsed.scheme}")
        return {}

    if not domain_allowed(parsed.hostname):
        logging.error(f"[!] Dominio no permitido: {parsed.hostname}")
        return {}

    size_str = response.headers.get("Content-Length", "0")
    size = int(size_str) if size_str.isdigit() else 0

    cd = response.headers.get("Content-Disposition", "")
    if isinstance(cd, bytes):
        try:
            cd = cd.decode("utf-8")
        except Exception:
            cd = ""

    match = re.search(r'filename="?([^";]+)', cd)
    name = (
        sanitize_filename(match.group(1))
        if match
        else sanitize_filename(Path(parsed.path).name)
    )

    logging.debug(f"[*] parse_head: obtenido name={name}, size={size}")
    return {"url": final_url, "name": name, "size": size}


def get_from_cache(key: str) -> Optional[Dict[str, object]]:
    if key in _FETCH_CACHE:
        _FETCH_CACHE.move_to_end(key)  # Marca como más reciente
        return _FETCH_CACHE[key]
    return None


def add_to_cache(key: str, value: Dict[str, object]) -> None:
    _FETCH_CACHE[key] = value
    _FETCH_CACHE.move_to_end(key)  # Marca como más reciente
    if len(_FETCH_CACHE) > MAX_CACHE_SIZE:
        _FETCH_CACHE.popitem(last=False)  # Elimina menos recientemente usado


def fetch_odt_download_info(download_id: str) -> Optional[Dict[str, object]]:
    """
    Obtiene metadatos del instalador del Office Deployment Tool (ODT)
    desde la página oficial de Microsoft a partir de un `download_id`.

    Realiza las siguientes acciones:
    - Descarga la página HTML correspondiente al ID usando `requests`.
    - Extrae el primer enlace válido a `.exe` desde `download.microsoft.com`.
    - Valida el dominio, protocolo HTTPS y cabeceras.
    - Realiza una solicitud HEAD para obtener nombre y tamaño del archivo.
    - Almacena resultados en caché local para evitar peticiones repetidas.

    La caché solo vive durante la ejecución del proceso (no persistente).

    Args:
        download_id (str): ID numérico oficial del ODT de Microsoft.
            Ejemplo: "49117"

    Returns:
        Optional[Dict[str, object]]: Diccionario con las claves:
            - 'url'  (str): Enlace directo al ejecutable .exe.
            - 'name' (str): Nombre limpio del archivo.
            - 'size' (int): Tamaño en bytes del archivo.
        Devuelve None si ocurre un error o si el ID no es válido.
    """

    if not download_id.isdigit():
        logging.error("[!] download_id inválido, debe ser numérico")
        return None

    cached = get_from_cache(download_id)
    if cached:
        logging.debug(
            f"[*] fetch_odt_download_info: usando cache para id={download_id}"
        )
        return cached

    summary_url = f"https://www.microsoft.com/en-us/download/details.aspx?id={download_id}"  # noqa: E501
    logging.debug(
        f"[*] fetch_odt_download_info: cargando página {summary_url}"
    )
    html_content = get_page_html(summary_url)
    if not html_content:
        logging.error(
            f"[!] fetch_odt_download_info: HTML vacío para id={download_id}"
        )
        return None

    try:
        tree = html.fromstring(html_content)
        # Buscar todos los enlaces <a> con href que terminen en .exe
        # y en dominio permitido
        links = tree.xpath("//a[@href]")
        exe_links = []
        for link in links:
            href = link.get("href")
            if href and href.lower().endswith(".exe"):
                parsed = urlparse(href)
                if parsed.scheme in ("http", "https") and domain_allowed(
                    parsed.hostname
                ):
                    exe_links.append(href)

        # Eliminar duplicados manteniendo orden
        enlaces_unicos = list(dict.fromkeys(exe_links))
        logging.debug(
            "[*] fetch_odt_download_info: encontrados "
            f"{len(enlaces_unicos)} enlaces .exe"
        )

    except Exception as e:
        logging.error(f"[!] Error al parsear HTML con lxml: {e}")
        return None

    if not enlaces_unicos:
        logging.error("[!] No se encontró ningún enlace .exe en la página.")
        return None

    info = parse_head(enlaces_unicos[0])
    if not info:
        logging.error("[!] Error en parse_head, info vacía")
        return None

    add_to_cache(download_id, info)
    logging.debug(
        "[*] fetch_odt_download_info: info almacenada para id="
        f"{download_id} -> {info}"
    )
    return info
