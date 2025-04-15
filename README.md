# ManagerOfficeScriptTool - Instalación y Desinstalación Automática de Office

Este proyecto facilita la instalación y desinstalación de Microsoft Office en equipos Windows mediante un script Python. Es compatible con Office 2013, 2016, 2019, LTSC-2021, LTSC-2024 y 365 (32 y 64 Bits).

---

## ¿Qué es ManagerOfficeScriptTool?

`ManagerOfficeScriptTool` es una herramienta automatizada para detectar, desinstalar e instalar Microsoft Office, todo desde una sola aplicación. A partir de la versión **4.0**, el proyecto fue **refactorizado completamente** en Python utilizando clases, logging profesional, sanitización de rutas sensibles y estructura modular.  
Se distribuye como un ejecutable `.exe` creado con **PyInstaller**, eliminando el uso de `cx_Freeze`.

> El código fuente es completamente abierto. Puedes revisarlo, modificarlo y compilar tu propia versión con el archivo `.spec` incluido.

---

## Características

- ✅ **Refactorización completa en clases**: `OfficeManager`, `ODTManager`, `OfficeUninstaller`, `OfficeInstaller`, `OfficeSelectionWindow`.
- ✅ **Desinstalación limpia utilizando ODT** (ya no se utiliza SaRA).
- ✅ **Detección detallada** de versiones instaladas de Office.
- ✅ **Interfaz gráfica (Tkinter)** mejorada para seleccionar versión, apps, arquitectura e idioma.
- ✅ **Instalación automatizada** mediante configuración XML y `setup.exe`.
- ✅ **Soporte para Office 2013, 2016, 2019, LTSC-2021, LTSC-2024 y Microsoft 365**.
- ✅ **Permisos de administrador solicitados automáticamente** (`--uac-admin`).
- ✅ **Logging completo** en `logs/application.log`, con rutas sanitizadas para privacidad.
- ✅ **Descarga directa del ODT desde Microsoft** usando Selenium y WebDriver.

---

## Requisitos

- Windows 10 o superior.
- Acceso a internet (para descargar ODT).
- No es necesario tener Python instalado: se distribuye como `.exe`.

---

## Instalación

1. **Descarga la última versión** desde [Releases en GitHub](https://github.com/Rodri082/ManagerOfficeScriptTool/releases).
2. **Ejecuta como administrador** `ManagerOfficeScriptTool.exe`.
3. Detectará las versiones instaladas de Office, te ofrecerá desinstalarlas y luego te permitirá instalar una nueva versión desde una interfaz gráfica.

---

## Uso

1. **Inicio**: Solicita permisos de administrador (UAC).
2. **Detección**: Se listan las versiones instaladas de Office.
3. **Desinstalación opcional**: Puedes elegir una o más versiones para desinstalar con ODT.
4. **Instalación nueva**: Se abre la GUI para elegir versión, arquitectura, idioma y apps.
5. **Ejecución**: Se genera un `configuration.xml` y se lanza la instalación usando `setup.exe`.

---

## Archivos que descarga

- **Office Deployment Tool (ODT)** desde fuentes oficiales de Microsoft:
  - [ODT 2013](https://www.microsoft.com/en-us/download/details.aspx?id=36778)
  - [ODT 2016 y posteriores](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

---

## Compilación desde el código fuente

> Consulta [`RELEASE.md`](./RELEASE.md) para ver las dependencias exactas utilizadas.

El proyecto se empaqueta con:

- **PyInstaller 6.12.0**
- **Python 3.13.3**

Puedes compilarlo con el archivo `.spec`:

```bash
pyinstaller ManagerOfficeScriptTool.spec
```

O bien manualmente:

```bash
pyinstaller --onefile --icon=icon.ico --uac-admin ManagerOfficeScriptTool.py
```

---

## Transparencia

Este proyecto es **100% open source**. No instala software de terceros, no envía datos, y solo descarga herramientas directamente desde servidores oficiales de Microsoft.

---

## Créditos

- **[Office Deployment Tool (ODT)](http://aka.ms/ODT)**

---

## Estado del Proyecto

El proyecto está en estado **activo y estable**. Ha evolucionado de un script simple a una herramienta modular, robusta y confiable.

### 🛠 Planes Futuros

- Reemplazo completo de consola por GUI enriquecida (Tkinter + barra de progreso).
- Vista de logs en tiempo real desde la aplicación.
- Soporte para más idiomas.

---

## Licencia

Este proyecto está licenciado bajo la [Licencia MIT](./LICENSE).
