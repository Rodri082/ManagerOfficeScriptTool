@echo off
:: Obtener la ruta del archivo .bat actual
set SCRIPT_DIR=%~dp0

:: Comprobar si el script está ejecutándose como administrador
openfiles >nul 2>nul
if %errorlevel% neq 0 (
    echo No tienes permisos de administrador.
    echo Se le solicitaran permisos de administrador. Presione una tecla para continuar...
    pause >nul
    :: Solicitar permisos UAC para ejecutar el script como administrador
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

:: Mostrar mensaje de confirmación de privilegios elevados
echo Permisos de administrador confirmados.

:: Cambiar al directorio donde se encuentra el archivo .bat
cd /d "%SCRIPT_DIR%Files"

:: Verificar si la carpeta y el archivo existen
if not exist "%SCRIPT_DIR%Files" (
    echo La carpeta "Files" no existe. Verifica la ubicación del script.
    pause
    exit /b
)
if not exist "%SCRIPT_DIR%Files\Install.ps1" (
    echo El archivo "Install.ps1" no se encuentra en la carpeta "Files".
    pause
    exit /b
)

:: Ejecutar el script Install.ps1 con PowerShell
powershell -ExecutionPolicy RemoteSigned -File "Install.ps1"

:: Finalizar
pause
exit /b