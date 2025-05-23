@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo Current dir: %cd%
echo.

:: --- Inicio de lógica principal ---

:: Limpiar directorio build si existe
if exist build (
    echo Limpiando directorio de compilacion anterior...
    rmdir /s /q build
)

:: Verificar si el archivo main.py existe
if not exist main.py (
    echo [ERROR] No se encontró el archivo main.py
    pause
    exit /b 1
)

:: Verificar si el compilador MSVC (cl.exe) está disponible
where cl >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontró cl.exe.
    pause
    exit /b 1
)

:: Verificar si Python está disponible
where py >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontró Python.
    pause
    exit /b 1
)

:: Verificar si Nuitka está disponible
where nuitka >nul 2>&1 || (
    py -m nuitka --version >nul 2>&1 || (
        echo [ERROR] No se encontró Nuitka ni como comando ni como módulo.
        pause
        exit /b 1
    )
)

:: Ejecutar Python inline y guardar salida en variable sin usar for /f
set "MIME_TYPES_PATH="
for /f "usebackq delims=" %%F in (`py -c "import scrapy;from pathlib import Path;p=Path(scrapy.__file__).parent/'mime.types';print(p if p.exists() else '')"`) do (
    set "MIME_TYPES_PATH=%%F"
)

if "!MIME_TYPES_PATH!"=="" (
    echo [ERROR] No se encontró el archivo mime.types de Scrapy.
    pause
    exit /b 1
)

echo [OK] Ruta mime.types detectada: "!MIME_TYPES_PATH!"

echo Iniciando compilacion con Nuitka...

python -m nuitka ^
    main.py ^
    --standalone ^
    --enable-plugin=tk-inter ^
    --windows-icon-from-ico=icon.ico ^
    --include-data-file=config.yaml=config.yaml ^
    --include-data-file="!MIME_TYPES_PATH!"=scrapy/mime.types ^
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
    exit /b 1
)

if exist "build\main.dist\ManagerOfficeScriptTool.exe" (
    echo [OK] Compilación completada exitosamente.
    echo Ejecutable generado en: %cd%\build\main.dist\ManagerOfficeScriptTool.exe
) else (
    if exist "build\ManagerOfficeScriptTool.exe" (
        echo [OK] Compilación completada exitosamente.
        echo Ejecutable generado en: %cd%\build\ManagerOfficeScriptTool.exe
    )
)


pause
endlocal
