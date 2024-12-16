# ManagerOfficeTool - Instalación y Desinstalación Automática de Office

Este proyecto tiene como objetivo facilitar la instalación y desinstalación de **Microsoft Office** en equipos Windows mediante una serie de scripts automatizados. El proceso es compatible con versiones de Office 2013, 2016, 2019, y 2021.

## Descripción

Este repositorio contiene una serie de scripts en **PowerShell** y **Python** que automatizan la instalación y desinstalación de Office en equipos Windows. El flujo de trabajo está diseñado para detectar y eliminar versiones previas de Office antes de proceder con la instalación de una nueva versión.

### Características

- **Verificación de permisos de administrador**: Se asegura de que el script se ejecute con privilegios de administrador.
- **Detección de versiones previas de Office**: Utiliza el script [Get-OfficeVersion.ps1](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Management/Get-OfficeVersion/Get-OfficeVersion.ps1) para detectar versiones existentes de Office en el sistema.
- **Desinstalación de versiones previas**: Si se detectan versiones previas, el script pregunta al usuario si desea desinstalarlas utilizando el script [Remove-PreviousOfficeInstalls.ps1](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls).
- **Instalación automática de Office**: Después de limpiar las versiones previas, el script procede a instalar Office usando la herramienta **Office Deployment Tool (ODT)**.
- **Soporte para versiones de Windows**: Compatible con Windows 7 y versiones posteriores, tanto en arquitecturas de 32 como 64 bits.

## Requisitos

- **Sistemas operativos soportados**: Windows 7 y versiones posteriores (32-bit y 64-bit).
- **Dependencias**: No se requiere ninguna dependencia adicional si se descarga el archivo desde el Release, el cual contiene el script Python empaquetado como ejecutable (.exe) para su uso sin necesidad de tener Python instalado.

## Instalación

1. **Descargar**: [ManagerOfficeTool.zip](https://github.com/Rodri082/ManagerOfficeTool/releases)
2. **Descomprimir**: [ManagerOfficeTool.zip]
3. **Ejecutar el archivo [RunInstallOffice.bat](./RunInstallOffice.bat)**:
    - Si no tienes permisos de administrador, el script te solicitará acceso mediante UAC (Control de cuentas de usuario).
    - El script [RunInstallOffice.bat](./RunInstallOffice.bat) verificará si tienes permisos de administrador y procederá a ejecutar el script [Install.ps1](./Files/Install.ps1).

## Uso

1. **Ejecución inicial**: 
    - Al ejecutar el archivo [`RunInstallOffice.bat`](./RunInstallOffice.bat), el sistema comprobará si está ejecutándose con permisos de administrador. Si no es así, se solicitarán permisos elevados mediante UAC.
    - Luego, se ejecutará el script [Install.ps1](./Files/Install.ps1).

2. **Detección de versiones de Office**:
    - Al inicio de la ejecución de [Install.ps1](./Files/Install.ps1), se preguntará al usuario si desea detectar versiones previas de Office instaladas en el sistema.
    - Si el usuario acepta, se descargará y ejecutará el script [Get-OfficeVersion.ps1](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Management/Get-OfficeVersion/Get-OfficeVersion.ps1) desde el repositorio de GitHub de Microsoft.

3. **Desinstalación de versiones previas (si es necesario)**:
    - Si se detectan versiones de Office instaladas, el usuario puede optar por desinstalar esas versiones usando el script [Remove-PreviousOfficeInstalls.ps1](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls) de Microsoft.
    - Si el usuario elige no desinstalar las versiones previas, se le advertirá sobre posibles conflictos.

4. **Instalación de una nueva versión de Office**:
    - Después de la desinstalación (si se aplica), el script ejecutará el archivo Python [DeploymentScriptTool.py](./Files/DeploymentScriptTool.py) para proceder con la instalación de la nueva versión de Office, usando los archivos de configuración y las herramientas necesarias.

## Archivos descargados automáticamente

- Los siguientes archivos se descargan automáticamente durante la ejecución de los scripts:
    - **ODT (Office Deployment Tool)**: El script descargará el archivo ejecutable de ODT desde los enlaces oficiales de Microsoft según la versión de Office seleccionada:
        - Office 2013: [ODT 2013](https://www.microsoft.com/en-us/download/details.aspx?id=36778)
        - Office 2016/2019/2021: [ODT 2016/2019/2021](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
        - **Documentación de ODT**: [Visita la documentación oficial de ODT](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-office-deployment-tool).
    - **Scripts de desinstalación**: Se descargan desde el repositorio de GitHub [Office-IT-Pro-Deployment-Scripts](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts) :
        - [Get-OfficeVersion.ps1](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Management/Get-OfficeVersion/Get-OfficeVersion.ps1)
        - [Remove-PreviousOfficeInstalls.ps1](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls)
        - [OffScrub03.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub03.vbs)
        - [OffScrub07.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub07.vbs)
        - [OffScrub10.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub10.vbs)
        - [OffScrub_O15msi.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O15msi.vbs)
        - [OffScrub_O16msi.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O16msi.vbs)
        - [OffScrubc2r.vbs](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrubc2r.vbs)
        - [Office2013Setup.exe](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Office2013Setup.exe)
        - [Office2016Setup.exe](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Office2016Setup.exe)

## Créditos

Este proyecto utiliza herramientas y scripts de los siguientes repositorios de Microsoft:

- **[Office IT Pro Deployment Scripts](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)**  
  Este repositorio proporciona scripts de Microsoft para la gestión de Office, como la detección de versiones de Office instaladas y la eliminación de instalaciones anteriores. Estos scripts son utilizados en este proyecto para facilitar la instalación y desinstalación de Microsoft Office.

- **[Microsoft Office Deployment Tool (ODT)](https://learn.microsoft.com/en-us/microsoft-365/apps/deploy/overview-office-deployment-tool)**  
  ODT es una herramienta de Microsoft utilizada para la instalación y configuración de Office. Este proyecto automatiza su uso para instalar y configurar las versiones de Office.

Gracias a Microsoft por el desarrollo de estas herramientas y por su liberación bajo la **MIT License**.

## Licencia

Este proyecto está licenciado bajo la MIT License - ver el archivo [LICENSE](./LICENSE) para más detalles.
