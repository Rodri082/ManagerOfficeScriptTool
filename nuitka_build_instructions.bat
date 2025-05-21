@echo off
setlocal

REM Detectar dinámicamente la ruta de scrapy/mime.types
FOR /F "usebackq delims=" %%F IN (`py -c "import scrapy; from pathlib import Path; p = Path(scrapy.__file__).parent / 'mime.types'; print(p if p.exists() else '')"`) DO (
    set "MIME_TYPES_PATH=%%F"
)

IF NOT DEFINED MIME_TYPES_PATH (
    echo [ERROR] No se encontró el archivo mime.types de Scrapy.
    exit /b 1
)

echo [OK] Ruta mime.types detectada: %MIME_TYPES_PATH%

REM Compilación con Nuitka
py -m nuitka ^
--standalone ^
--enable-plugin=tk-inter ^
--windows-icon-from-ico=icon.ico ^
--include-data-file=config.yaml=config.yaml ^
--include-data-file="%MIME_TYPES_PATH%"=scrapy/mime.types ^
--include-package=core ^
--include-package=gui ^
--include-package=scripts ^
--include-module=utils ^
--include-package=scrapy ^
--include-package=scrapy.spiderloader ^
--include-package=scrapy.spiders ^
--include-package=scrapy.utils ^
--include-package=scrapy.crawler ^
--include-package=scrapy.selector ^
--include-package=scrapy.link ^
--include-package=scrapy.downloadermiddlewares ^
--include-package=scrapy.pipelines ^
--include-package=scrapy.extensions ^
--include-package=scrapy.core ^
--include-package=scrapy.settings ^
--include-package=scrapy.logformatter ^
--nofollow-import-to=unittest,doctest ^
--company-name="Rodri082" ^
--product-name="ManagerOfficeScriptTool" ^
--file-version=5.0.0.0 ^
--product-version=5.0.0.0 ^
--file-description="Herramienta ManagerOfficeScriptTool" ^
--copyright="Licencia MIT © 2024 Rodri082" ^
--windows-uac-admin ^
--output-dir=build ^
--output-filename=ManagerOfficeScriptTool.exe ^
--msvc=latest ^
--lto=yes ^
main.py

endlocal