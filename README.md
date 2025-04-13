# ManagerOfficeScriptTool - Instalación y Desinstalación Automática de Office

Este proyecto tiene como objetivo facilitar la instalación y desinstalación de Microsoft Office en equipos Windows mediante un script Python. El proceso es compatible con versiones de Office 2013, 2016, 2019, LTSC-2021, LTSC-2024 y 365 (32 y 64 Bits).

## ¿Qué es ManagerOfficeScriptTool?

`ManagerOfficeScriptTool` es una herramienta automatizada para la instalación y limpieza de versiones de Microsoft Office, empaquetada en un único ejecutable listo para usar. A partir de la versión **4.0**, el proyecto se compila exclusivamente con **PyInstaller**, eliminando el uso de `cx_Freeze`.

> El código fuente es completamente abierto. Cualquier usuario puede auditarlo, revisar su comportamiento y compilar su propia versión utilizando el archivo `.spec` incluido.

---

## Características

> A partir de la versión 4.0, se eliminó el uso de los antiguos scripts `.vbs` de OffScrub para reducir la complejidad y evitar falsos positivos generados por algunos antivirus. Ahora la desinstalación se realiza exclusivamente mediante SaRA.

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

- **SaRA ComandLineVersion ZIP** desde [SaRA_CommandLineVersion](https://aka.ms/SaRA_CommandLineVersion )

---

## Compilación desde el código fuente

> Para una lista completa de bibliotecas utilizadas en el script, consulta la sección correspondiente en el archivo [`RELEASE.md`](./RELEASE.md).

El proyecto se empaqueta con **PyInstaller 6.12.0** utilizando **Python 3.13.2**. Se incluye el archivo `ManagerOfficeScriptTool.spec` para permitir compilar localmente:

```bash
pyinstaller ManagerOfficeScriptTool.spec
```


También podés compilar manualmente con:

```bash
pyinstaller --onefile --icon=icon.ico --noupx --uac-admin ManagerOfficeScriptTool.py
```


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

## Estado del Proyecto

Este proyecto se encuentra en un estado **activo y estable**. Ha pasado de ser un simple script funcional a una herramienta profesional, automatizada y segura para la instalación y desinstalación de Microsoft Office en sistemas Windows.

### 🛠 Planes Futuros (sujetos a evaluación)
- Reemplazo completo de la consola por una **interfaz gráfica enriquecida**.
  - El log de instalación se mostrará en una ventana tipo terminal (`ScrolledText`) con scroll y colores para errores/éxitos.
  - Se incluirá una opción para **guardar el log en un archivo `.txt`**.
  - Se integrará una **barra de progreso indeterminada** mientras se realiza la instalación de Office, mejorando la experiencia visual sin necesidad de mostrar progreso exacto.
- Reemplazo total de la consola por una **interfaz gráfica completa**, en la cual todos los mensajes actualmente mostrados por consola se integren visualmente en la aplicación (Tkinter).
- Inclusión de más idiomas.

Contribuciones, ideas o reportes de errores son bienvenidos.

## Licencia

Este proyecto está licenciado bajo la [Licencia MIT](./LICENSE).

---

