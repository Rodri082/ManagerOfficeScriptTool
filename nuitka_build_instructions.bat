@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo Current dir: %cd%

REM Detectar dinámicamente la ruta de scrapy/mime.types
set "MIME_TYPES_PATH="
for /f "usebackq delims=" %%F in (`py -c "import scrapy; from pathlib import Path; p = Path(scrapy.__file__).parent / 'mime.types'; print(p if p.exists() else '')"`) do (
    set "MIME_TYPES_PATH=%%F"
)

if not defined MIME_TYPES_PATH (
    echo [ERROR] No se encontró el archivo mime.types de Scrapy.
    pause
    exit /b 1
)

echo [OK] Ruta mime.types detectada: "%MIME_TYPES_PATH%"

REM Compilación con Nuitka
python -m nuitka ^
    main.py ^
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
    --lto=yes

if errorlevel 1 (
    echo [ERROR] La compilacion fallo.
    pause
)

endlocal
