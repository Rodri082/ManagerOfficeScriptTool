import json
import re
import tempfile
from pathlib import Path
from urllib.parse import urlparse

import scrapy
import scrapy.utils.versions
from scrapy import Request, Spider
from scrapy.crawler import CrawlerProcess


def fetch_odt_download_info(download_id: str) -> dict:
    """
    Ejecuta un spider de Scrapy para obtener la URL, nombre y tamaño
    del ejecutable del ODT desde el sitio oficial de Microsoft.

    Args:
        download_id (str): ID de descarga de Microsoft.

    Returns:
        dict: Contiene las claves 'url', 'name' y 'size', o vacío si falla.
    """

    def _patched_version(package):
        return "unknown"

    scrapy.utils.versions._version = _patched_version

    with tempfile.NamedTemporaryFile(
        mode="w+", suffix=".json", delete=False
    ) as tmp:
        temp_path = Path(tmp.name)

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
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(result, f)

    process = CrawlerProcess({"LOG_LEVEL": "ERROR"})
    process.crawl(ODTSpider)
    process.start()

    try:
        with open(temp_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}
    finally:
        temp_path.unlink(missing_ok=True)
