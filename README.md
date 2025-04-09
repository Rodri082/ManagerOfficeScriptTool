# ManagerOfficeScriptTool - Instalación y Desinstalación Automática de Office

Este proyecto tiene como objetivo facilitar la instalación y desinstalación de Microsoft Office en equipos Windows mediante un script Python. El proceso es compatible con versiones de Office 2013, 2016, 2019, LTSC-2021, LTSC-2024 y 365 (32 y 64 Bits).

## ¿Qué es ManagerOfficeScriptTool?

`ManagerOfficeScriptTool` es una herramienta automatizada para la instalación y limpieza de versiones de Microsoft Office, empaquetada en un único ejecutable listo para usar. A partir de la versión **4.0**, el proyecto se compila exclusivamente con **PyInstaller**, eliminando el uso de `cx_Freeze`.

> El código fuente es completamente abierto. Cualquier usuario puede auditarlo, revisar su comportamiento y compilar su propia versión utilizando el archivo `.spec` incluido.

---

## Características

- ✅ **Instalación y desinstalación automatizada** mediante un único ejecutable.
- ✅ **Detección de versiones instaladas de Office**.
- ✅ **Desinstalación limpia con SaRA**.
- ✅ **Instalación de nuevas versiones con Office Deployment Tool (ODT)**.
- ✅ **Soporte para Office 2013, 2016, 2019, LTSC-2021, LTSC-2024 y Microsoft 365**.
- ✅ **Modo gráfico (Tkinter)** para configuración interactiva.
- ✅ **Solicita permisos de administrador al ejecutarse** (gracias al manifiesto o `--uac-admin`).

---

## Requisitos

- Windows 10 o superior.
- Acceso a internet (para descargar ODT y herramientas de limpieza).
- No es necesario tener Python instalado: se distribuye como `.exe` autónomo.

---

## Instalación

1. **Descargar la versión más reciente** desde la página de [Releases en GitHub](https://github.com/Rodri082/ManagerOfficeScriptTool/releases).

2. **Ejecutar el archivo** `ManagerOfficeScriptTool.exe` como administrador. El programa detectará versiones instaladas de Office, ofrecerá desinstalarlas, y permitirá instalar una nueva versión desde una GUI.

3. **Listo**. No requiere instalación adicional ni dependencias externas.

> Nota: Las versiones anteriores incluían una opción empaquetada con `cx_Freeze`, pero desde la versión 4.0 esta opción ha sido retirada para unificar el proceso de construcción con PyInstaller.

---

## Uso

1. **Inicio**: Al ejecutar el `.exe`, se solicitará elevación de permisos (UAC).
2. **Detección de Office**: El script escaneará el sistema y mostrará las versiones de Office instaladas.
3. **Desinstalación limpia (opcional)**: Se podrá elegir desinstalar versiones anteriores mediante SaRA.
4. **Instalación de nueva versión**: Se abrirá una interfaz para seleccionar versión, arquitectura, idioma y apps a instalar.
5. **Configuración**: El script generará un `configuration.xml` compatible con ODT y ejecutará la instalación de Office.

---

## Archivos que descarga

- **Office Deployment Tool (ODT)** desde los enlaces oficiales de Microsoft:
  - [ODT 2013](https://www.microsoft.com/en-us/download/details.aspx?id=36778)
  - [ODT 2016/2019/LTSC/365](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

- **SaRA Enterprise ZIP** desde [SaRA_EnterpriseVersion](https://learn.microsoft.com/es-es/microsoft-365/troubleshoot/administration/sara-command-line-version)

---

## Compilación desde el código fuente

El proyecto se empaqueta con **PyInstaller 6.12.0** utilizando **Python 3.13.2**. Se incluye el archivo `ManagerOfficeScriptTool.spec` para permitir compilar localmente:

```bash
pyinstaller ManagerOfficeScriptTool.spec
```

> Es importante ejecutar la consola como administrador si vas a generar el `.exe` final.

También podés compilar manualmente con:

```bash
pyinstaller --onefile --icon=icon.ico --noupx --uac-admin ManagerOfficeScriptTool.py
```

> Requiere `pyinstaller-hooks-contrib` versión 2025.2 o superior.

---

## Transparencia

Este proyecto es 100% **open source**. No realiza conexiones sospechosas, no instala software de terceros, ni recopila información del usuario. Todas las herramientas utilizadas provienen de fuentes oficiales de Microsoft.

Puedes revisar, modificar o compilar el script por tu cuenta. Se alienta a la comunidad a auditar el código para confirmar su integridad.

---

## Créditos

- **[Office IT Pro Deployment Scripts](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)**
- **[Office Deployment Tool (ODT)](http://aka.ms/ODT)**
- **[SaRA - Support and Recovery Assistant](https://learn.microsoft.com/es-es/microsoft-365/troubleshoot/administration/sara-command-line-version)**

---

## Licencia

Este proyecto está licenciado bajo la [Licencia MIT](./LICENSE).

---

