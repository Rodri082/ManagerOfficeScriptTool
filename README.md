# ManagerOfficeScriptTool - Instalación y Desinstalación Automática de Office
Este proyecto tiene como objetivo facilitar la instalación y desinstalación de Microsoft Office en equipos Windows mediante un script Python. El proceso es compatible con versiones de Office 2013, 2016, 2019, LTSC-2021, LTSC-2024 y 365. (Versiones de 32 y 64 Bits)

## Descripción
Este repositorio contiene un script en Python que automatiza la instalación y desinstalación de Office en equipos Windows. El flujo de trabajo está diseñado para detectar y eliminar versiones previas de Office antes de proceder con la instalación de una nueva versión.

## Características
- **Instalación y desinstalación automatizada**: Todo el proceso se realiza mediante un único archivo ejecutable `ManagerOfficeScriptTool.exe`.
- **Detección de versiones de Office**: El script ahora detecta las versiones de Office instaladas directamente.
- **Desinstalación con OffScrub**: El script utiliza **Scripts OffScrub** para desinstalar versiones anteriores de Office.

## Requisitos
- **Dependencias**: El script está empaquetado como un archivo ejecutable `ManagerOfficeScriptTool.exe`, por lo que no deberia ser necesario tener Python instalado.

## Instalación de ManagerOfficeScriptTool
1. **Descargar**:
   - Puedes elegir entre dos opciones de descarga:
     - El archivo `.exe` empaquetado con **PyInstaller** utilizando Python 3.13.1.
     - El archivo `.zip` empaquetado con **cx_Freeze** utilizando Python 3.12.8.
   - Descarga desde el siguiente enlace: [ManagerOfficeScriptTool releases](https://github.com/Rodri082/ManagerOfficeScriptTool/releases).

2. **Descomprimir** (solo si descargaste el archivo `.zip`):
   - Si optaste por el archivo `.zip`, extrae su contenido en una ubicación de tu elección.

3. **Ejecutar**:
   - Si descargaste el archivo `.exe`, simplemente ejecútalo.
   - Si descargaste el archivo `.zip`, navega al directorio donde extrajiste los archivos y ejecuta `ManagerOfficeScriptTool.exe`.
   - El programa se encargará de la instalación o desinstalación de Office según las opciones que selecciones durante el proceso.

## Uso
1. **Ejecución inicial**: Al ejecutar el archivo `ManagerOfficeScriptTool.exe`, el sistema comprobará si está ejecutándose con permisos de administrador. Si no es así, se solicitará que vuelva a ejecutar con permisos elevados.
2. **Detección de versiones de Office**: El script detectará automáticamente las versiones de Office instaladas en el sistema.
3. **Desinstalación de versiones previas (si es necesario)**:
    - Si se detectan versiones de Office instaladas, el usuario puede optar por desinstalar esas versiones usando scripts [.vbs]().
    - Si el usuario elige no desinstalar las versiones previas, se le advertirá sobre posibles conflictos.
4. **Instalación de una nueva versión de Office**: Después de la desinstalación (si es necesario), el script procederá a instalar la nueva versión de Office utilizando la **Office Deployment Tool (ODT)**.

## Archivos descargados automáticamente
Durante la ejecución del script, se descargan los siguientes archivos:
   - **ODT (Office Deployment Tool)**: El script descargará el archivo ejecutable de ODT desde los enlaces oficiales de Microsoft según la versión de Office seleccionada:
        - Office 2013: [ODT 2013](https://www.microsoft.com/en-us/download/details.aspx?id=36778)
        - Office 2016/2019/LTSC-2021/LTSC-2024/365: [ODT 2016/2019/LTSC-2021/LTSC-2024/365](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
   - **Scripts de desinstalación**: Se descargan desde el repositorio de GitHub [Office-IT-Pro-Deployment-Scripts](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts):
        - [OffScrub03.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub03.vbs)
        - [OffScrub07.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub07.vbs)
        - [OffScrub10.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub10.vbs)
        - [OffScrub_O15msi.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O15msi.vbs)
        - [OffScrub_O16msi.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O16msi.vbs)
        - [OffScrubc2r.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrubc2r.vbs)

## Créditos
Este proyecto utiliza herramientas y scripts de los siguientes repositorios de Microsoft:
- **[Office IT Pro Deployment Scripts](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)**  
  Este repositorio proporciona varios scripts para la desinstalación de instalaciones anteriores de Office.
- **[Office Deployment Tool (ODT)](http://aka.ms/ODT)**  
  ODT es una herramienta de Microsoft utilizada para la instalación y configuración de Office. Este proyecto automatiza su uso para instalar y configurar las versiones de Office.

## Licencia
Este proyecto está licenciado bajo la MIT License - ver el archivo [LICENSE](./LICENSE) para más detalles.
