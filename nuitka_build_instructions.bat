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
REM     - Python 3.13+ instalado y en PATH
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
    echo(
    REM Indica al usuario cómo ejecutar como administrador
    echo Haz clic derecho sobre este archivo y selecciona "Ejecutar como administrador".
    echo(
    pause
    exit /b 1
)

:: --- Comprobación de permisos de escritura en el directorio actual ---
echo Verificando permisos de escritura en el directorio de trabajo...
set "PERM_TEST_FILE=%cd%\__perm_test__.tmp"
echo test > "%PERM_TEST_FILE%" 2>nul
if not exist "%PERM_TEST_FILE%" (
    echo [ERROR] No tienes permisos de escritura en este directorio: %cd%
    echo(
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
    echo(
    pause
    exit /b 1
)

:: Verificar si el archivo config.yaml existe
if not exist config.yaml (
    echo [ERROR] No se encontró config.yaml
    echo(
    pause
    exit /b 1
)

:: Verificar si el archivo icon.ico existe
if not exist icon.ico (
    echo [ERROR] No se encontró icon.ico
    echo(
    pause
    exit /b 1
)

:: Verificar si el compilador MSVC (cl.exe) está disponible
where cl >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontró cl.exe.
    echo(
    pause
    exit /b 1
)

:: Verificar si Python está disponible
where py >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontró Python.
    echo(
    pause
    exit /b 1
)

:: --- Verificar e instalar dependencias desde requirements.txt ---
if exist "requirements.txt" (
    echo Verificando dependencias en requirements.txt...
    set "MISSING=0"
    for /f "usebackq delims=" %%L in ("%cd%\requirements.txt") do (
        set "line=%%L"
        rem Trim leading spaces
        for /f "tokens=* delims= " %%S in ("!line!") do set "line=%%S"
        if "!line!"=="" (
            rem linea vacia -> saltar
        ) else (
            set "first=!line:~0,1!"
            if "!first!"=="#" (
                rem comentario -> saltar
            ) else (
                for /f "tokens=1 delims==<>" %%P in ("!line!") do set "pkg=%%P"
                for /f "tokens=* delims= " %%Q in ("!pkg!") do set "pkg=%%Q"
                py -m pip show "!pkg!" >nul 2>&1
                if errorlevel 1 (
                    echo [MISSING] !pkg!
                    set /a MISSING+=1
                ) else (
                    echo [OK] !pkg!
                )
            )
        )
    )
    if "!MISSING!"=="0" (
        echo Todas las dependencias estan instaladas.
    ) else (
        echo Instalando paquetes faltantes...
        py -m pip install -r requirements.txt
        if errorlevel 1 (
            echo [ERROR] No se pudieron instalar los paquetes. Asegurate de tener permisos y conexion a Internet.
            echo(
            pause
            exit /b 1
        ) else (
            echo [OK] Dependencias instaladas.
        )
    )
) else (
    echo [WARN] No se encontro requirements.txt. Se omitira la verificacion de dependencias.
)

:: Verificar si Nuitka está disponible
where nuitka >nul 2>&1
if errorlevel 1 (
    echo Nuitka no está instalado. Intentando instalarlo...
    py -m pip install nuitka
    if errorlevel 1 (
        echo [ERROR] No se pudo instalar Nuitka. Asegúrate de tener permisos y una conexión a Internet.
        echo(
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
    --include-package=manager_office_tool ^
    --include-data-files=./config.yaml=config.yaml ^
    --include-data-files=./icon.ico=icon.ico ^
    --windows-icon-from-ico=icon.ico ^
    --company-name="Rodri082" ^
    --product-name="ManagerOfficeScriptTool" ^
    --file-version=5.1.2.0 ^
    --product-version=5.1.2.0 ^
    --file-description="Herramienta ManagerOfficeScriptTool" ^
    --copyright="Licencia MIT Copyright 2024 Rodri082" ^
    --windows-uac-admin ^
    --output-dir=build ^
    --output-filename=ManagerOfficeScriptTool.exe ^
    --msvc=latest ^
    --lto=yes ^
    --nofollow-import-to=PIL._webp ^
    --nofollow-import-to=unittest ^
    --nofollow-import-to=pytest ^
    --nofollow-import-to=tests ^
    --report=build/compilation-report.xml

if errorlevel 1 (
    echo [ERROR] La compilacion fallo. Consulta build\compilation-report.xml para mas detalles.
    echo(
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

echo(
pause
endlocal
