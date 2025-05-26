@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM ManagerOfficeScriptTool - Script de compilación con Nuitka
REM
REM USO:
REM   Ejecuta este archivo desde el directorio raíz del proyecto.
REM   ¡IMPORTANTE! Ejecuta este script como administrador.
REM   Haz clic derecho sobre este archivo y selecciona "Ejecutar como administrador".
REM
REM REQUISITOS:
REM     - Python 3.10+ instalado y en PATH
REM     - MSVC (cl.exe) disponible en PATH (Visual Studio Build Tools)
REM     - Nuitka instalado (el script intentará instalarlo si no está)
REM     - config.yaml y icon.ico presentes en el directorio raíz
REM
REM Este script compila el proyecto en modo standalone usando Nuitka,
REM limpia compilaciones previas y verifica la creación del ejecutable.
REM ============================================================================

cd /d "%~dp0"

echo Current dir: %cd%
echo.

REM --- Comprobación de permisos de administrador ---
net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Debes ejecutar este script como administrador.
    echo Haz clic derecho sobre este archivo y selecciona "Ejecutar como administrador".
    pause
    exit /b 1
)

:: --- Comprobación de permisos de escritura en el directorio actual ---
echo Verificando permisos de escritura en el directorio de trabajo...
set "PERM_TEST_FILE=%cd%\__perm_test__.tmp"
echo test > "%PERM_TEST_FILE%" 2>nul
if not exist "%PERM_TEST_FILE%" (
    echo [ERROR] No tienes permisos de escritura en este directorio: %cd%
    pause
    exit /b 1
)
del "%PERM_TEST_FILE%"

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

:: Verificar si el archivo config.yaml existe
if not exist config.yaml (
    echo [ERROR] No se encontró config.yaml
    pause
    exit /b 1
)

:: Verificar si el archivo icon.ico existe
if not exist icon.ico (
    echo [ERROR] No se encontró icon.ico
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
where nuitka >nul 2>&1
if errorlevel 1 (
    echo Nuitka no está instalado. Intentando instalarlo...
    py -m pip install nuitka
    if errorlevel 1 (
        echo [ERROR] No se pudo instalar Nuitka. Asegúrate de tener permisos y una conexión a Internet.
        pause
        exit /b 1
    ) else (
        echo [OK] Nuitka instalado correctamente.
    )
)

echo Iniciando compilacion con Nuitka...

py -m nuitka ^
    main.py ^
    --standalone ^
    --enable-plugin=tk-inter ^
    --enable-plugin=pyside6 ^
    --windows-icon-from-ico=icon.ico ^
    --include-package=manager_office_tool ^
    --include-data-files=./config.yaml=config.yaml ^
    --nofollow-import-to=unittest,doctest,types_PyYAML,types_requests ^
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
    --report=build/compilation-report.xml

if errorlevel 1 (
    echo [ERROR] La compilacion fallo. Consulta build\compilation-report.xml para mas detalles.
    pause
    exit /b 1
)

:: Verificar si el ejecutable fue creado
:: Se verifica en dos ubicaciones posibles: main.dist y la raíz de build
set "EXE_PATH="
if exist "build\main.dist\ManagerOfficeScriptTool.exe" (
    set "EXE_PATH=%cd%\build\main.dist\ManagerOfficeScriptTool.exe"
) else if exist "build\ManagerOfficeScriptTool.exe" (
    set "EXE_PATH=%cd%\build\ManagerOfficeScriptTool.exe"
)

if defined EXE_PATH (
    echo [OK] Compilacion completada exitosamente.
    echo.
    echo Ejecutable generado en: !EXE_PATH!
) else (
    echo [ERROR] Compilación finalizada pero no se encontró el ejecutable.
)
:: --- Fin de lógica principal ---

echo.
pause
endlocal
