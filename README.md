# ManagerOfficeScriptTool - Instalación y Desinstalación Automática de Office

Este proyecto tiene como objetivo facilitar la instalación y desinstalación de Microsoft Office en equipos Windows mediante un script Python. El proceso es compatible con versiones de Office 2013, 2016, 2019 y 2021.

## Descripción

Este repositorio contiene un script en Python que automatiza la instalación y desinstalación de Office en equipos Windows. El flujo de trabajo está diseñado para detectar y eliminar versiones previas de Office antes de proceder con la instalación de una nueva versión.

A partir de la última actualización, el script ha sido mejorado para **eliminar la dependencia de los archivos `Install.ps1`, `RunInstallOffice.bat`, y scripts PowerShell de terceros**. Ahora, todo el proceso se maneja directamente desde el script en Python `ManagerOfficeScriptTool.py`, lo que facilita su uso y distribución.

## Características

- **Instalación y desinstalación automatizada**: Todo el proceso se realiza mediante un único archivo ejecutable `ManagerOfficeScriptTool.exe`, sin necesidad de tener Python instalado.
- **Detección de versiones de Office**: El script ahora detecta las versiones de Office instaladas directamente.
- **Desinstalación con SaRACmd**: El script utiliza **SaRACmd.exe** (Support and Recovery Assistant) para desinstalar versiones anteriores de Office.
- **Compatibilidad con versiones de Windows**: Compatible con Windows 7 y versiones posteriores, tanto en arquitecturas de 32 como 64 bits.
- **Interfaz de usuario amigable**: El script interactúa con el usuario mediante la terminal, solicitando la aceptación de los términos de uso antes de proceder con la desinstalación de Office.

## Requisitos

- **Sistemas operativos soportados**: Windows 7 y versiones posteriores (32-bit y 64-bit).
- **Dependencias**: El script está empaquetado como un archivo ejecutable `ManagerOfficeScriptTool.exe`, por lo que no es necesario tener Python instalado.

## Instalación (.exe)

1. **Descargar**: [ManagerOfficeTool.zip](https://github.com/Rodri082/ManagerOfficeScriptTool/releases).
2. **Descomprimir**: Extrae el contenido del archivo ZIP.
3. **Ejecutar**: Simplemente ejecuta el archivo `ManagerOfficeScriptTool.exe` en el directorio extraído. El archivo se encargará de la instalación o desinstalación de Office de acuerdo con las opciones que elija el usuario.

## Uso

1. **Ejecución inicial**: Al ejecutar el archivo `ManagerOfficeScriptTool.exe`, el sistema comprobará si está ejecutándose con permisos de administrador. Si no es así, se solicitarán permisos elevados.
   
2. **Detección de versiones de Office**: El script detectará automáticamente las versiones de Office instaladas en el sistema.

3. **Desinstalación de versiones previas**: 
   - Se preguntará al usuario si desea aceptar los **Términos y Condiciones** de **Microsoft Support and Recovery Assistant** (SaRACmd).
   - Si el usuario acepta, el script procederá a desinstalar la versión de Office utilizando **SaRACmd.exe**.
   - Si el usuario no acepta, se mostrará un mensaje.

4. **Instalación de una nueva versión de Office**: Después de la desinstalación (si es necesario), el script procederá a instalar la nueva versión de Office utilizando la **Office Deployment Tool (ODT)**.

## Archivos descargados automáticamente

Durante la ejecución del script, se descargan los siguientes archivos:

- **ODT (Office Deployment Tool)**: Se descargan los archivos de la herramienta oficial de Microsoft según la versión de Office seleccionada.
- **SaRACmd.exe**: Se descarga **Support and Recovery Assistant Command Line** para la desinstalación de versiones anteriores de Office.

## Créditos

Este proyecto utiliza herramientas y scripts de los siguientes repositorios de Microsoft:

- **[SaRA_CommandLineVersion](https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/administration/assistant-office-uninstall)**  
  **SaRA_CommandLineVersion** es una herramienta de línea de comandos de Microsoft Support and Recovery Assistant (SaRA) que facilita la desinstalación de versiones anteriores de Office de manera automatizada. Este proyecto utiliza **SaRACmd.exe** para gestionar la eliminación de instalaciones previas de Microsoft Office, simplificando el proceso para los usuarios al eliminar la necesidad de realizar la desinstalación manualmente o usar scripts PowerShell adicionales.


- **[Microsoft Office Deployment Tool (ODT)](http://aka.ms/ODT)**  
  ODT es una herramienta de Microsoft utilizada para la instalación y configuración de Office. Este proyecto automatiza su uso para instalar y configurar las versiones de Office.

Gracias a Microsoft por el desarrollo de estas herramientas y por su liberación bajo la **MIT License**.

## Licencia

Este proyecto está licenciado bajo la MIT License - ver el archivo [LICENSE](./LICENSE) para más detalles.
