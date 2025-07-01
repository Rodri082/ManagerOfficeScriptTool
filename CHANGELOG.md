# 📋 CHANGELOG

Historial de cambios del proyecto **ManagerOfficeScriptTool**

## [5.1] - 2025-07-01
### 🆕 Mejoras y nuevas funcionalidades

#### Interacción y experiencia de usuario
- **Menú de desinstalación avanzado**: Ahora permite seleccionar múltiples instalaciones de Office para desinstalar en una sola operación, además de las opciones previas (ninguna, todas, una específica). Se utilizan nuevas funciones robustas para validar la selección y evitar errores de input.
- **Funciones de consola ampliadas**: Se agregan `ask_menu_option`, `ask_single_valid_index` y `ask_multiple_valid_indices` para menús interactivos y selección múltiple, mejorando la flexibilidad y seguridad en la interacción por consola.
- **Mensajes y logs enriquecidos**: Todos los mensajes y logs son ahora más claros, detallados y útiles para auditoría y soporte. Se informa explícitamente cuando no se realiza ninguna acción o si el usuario cancela una operación.

#### Instalación y desinstalación
- **Desinstalación en grupo (multi-idioma)**: Ahora es posible desinstalar varias instalaciones del mismo producto (por ejemplo, diferentes idiomas) en una sola operación, agrupando por nombre base y ProductID. El XML de desinstalación se genera dinámicamente con todos los idiomas seleccionados.
- **Instalador con idioma registrado**: El instalador ahora recibe y muestra el idioma seleccionado, mejorando la trazabilidad y los mensajes al usuario durante la instalación.

#### GUI mejorada
- **Selección de Visio y Project**: Se añaden checkboxes en la GUI para incluir Visio y Project en la instalación, determinando dinámicamente el ProductID correcto según la versión de Office.
- **Validaciones exhaustivas**: La GUI valida que el usuario seleccione al menos una aplicación, versión, arquitectura e idioma antes de proceder, mostrando mensajes claros en caso de error.
- **Soporte para eliminar versiones MSI**: Se añade la opción "Eliminar versiones MSI (RemoveMSI)" en la GUI, agregando el nodo `<RemoveMSI/>` al XML de configuración si está seleccionado.
- **Generación avanzada de configuration.xml**: El archivo XML se genera y formatea profesionalmente usando `xml.etree.ElementTree` y `xml.dom.minidom`.

#### Configuración y compatibilidad
- **Ampliación de aplicaciones por versión**: Las listas de aplicaciones disponibles para cada versión de Office han sido revisadas y ampliadas, incluyendo sistemáticamente Teams, Publisher, Access, Lync, OneDrive, Groove, etc., especialmente en Office 365 y LTSC.
- **Idiomas soportados ampliados**: Se añaden numerosos idiomas adicionales, permitiendo una personalización mucho mayor en la instalación de Office.

#### Núcleo y utilidades
- **Caché LRU real en odt_fetcher.py**: Se implementa una caché LRU eficiente usando `OrderedDict` para la descarga de ODT, mejorando el control de memoria y evitando solicitudes duplicadas.
- **Parseo robusto de HTML con lxml**: Ahora se utiliza `lxml.html` para extraer enlaces de descarga, mejorando la fiabilidad y seguridad.
- **Validaciones y sanitización reforzadas**: Se refuerza la validación de dominios, la sanitización de nombres de archivo y la gestión de errores en la descarga y parseo de ODT.
- **Descarga reanudable de ODT**: El gestor de ODT ahora soporta descargas reanudables si el servidor lo permite, optimizando el uso de ancho de banda y tiempo.
- **Limpieza segura de temporales**: Se asegura la eliminación de carpetas temporales solo si existen y fueron utilizadas, informando al usuario del resultado.
- **Nuevo módulo path_utils**: Centraliza todas las utilidades de rutas, carpetas temporales y sanitización de logs, mejorando la organización y mantenibilidad.

#### Robustez y seguridad
- **Manejo de errores y logs mejorados**: Se agregan logs detallados para cada tipo de error en el acceso al registro de Windows y en operaciones críticas, facilitando la depuración y el soporte.
- **Sanitización avanzada de claves y rutas**: Se refuerza la anonimización de rutas y claves de registro en los logs para evitar exponer información sensible.

#### Dependencias y build
- **Nueva dependencia**: Se añade `lxml` en `requirements.txt` para el parseo robusto de HTML/XML.
- **Script de compilación Nuitka mejorado**: El script de build ahora incluye comprobaciones previas, instalación automática de dependencias, limpieza de builds previas, compilación optimizada y verificación del resultado, mostrando mensajes claros en cada paso.

---

**Resumen:**  
Esta versión refuerza la robustez, seguridad y flexibilidad del proyecto, amplía la compatibilidad con más idiomas y aplicaciones, y mejora significativamente la experiencia de usuario tanto en consola como en GUI. Se centralizan utilidades, se optimiza la gestión de descargas y temporales, y se profesionaliza el proceso de build

---
## [5.0] - 2025-05-26
### 🚀 Modularización total y refactor profesional
- El proyecto se reestructura completamente: de un solo archivo (`ManagerOfficeScriptTool.py`) a un **paquete modularizado**.
- Separación clara de responsabilidades en submódulos:
  - `core/`: lógica de negocio (gestión de Office, registro, ODT, instalación detectada).
  - `gui/`: interfaz gráfica moderna con ttkbootstrap.
  - `scripts/`: instalador y desinstalador.
  - `utils/`: utilidades generales (consola, GUI, logging, rutas).
  - `config.yaml`: configuración centralizada.
  - `main.py`: punto de entrada y orquestador del flujo.
- Cada módulo y clase cuenta con docstrings detallados y tipado de argumentos.

### 🛠️ Mejoras de arquitectura, seguridad y mantenibilidad
- Imports reorganizados y relativos a la nueva estructura de carpetas.
- Configuración, versiones, canales y apps movidos a `config.yaml` para fácil mantenimiento.
- Limpieza y manejo robusto de carpetas temporales.
- Logging profesional y seguro: rutas y claves anonimizadas, sin exponer datos sensibles.
- Acceso al registro de Windows de forma segura y eficiente.
- **No se aceptan rutas arbitrarias del usuario**: todas las rutas temporales y de trabajo se generan internamente.
- Cumplimiento estricto de PEP8: líneas ≤ 79 caracteres, uso de paréntesis en prints y logs largos.
- Uso de Black, isort, flake8 y mypy recomendado y compatible.

### 🏷️ Detección y visualización mejorada de instalaciones de Office
- Lógica de detección de instalaciones actualizada: ahora el `ProductID` y el `MediaType` se obtienen siempre desde el valor `ProductReleaseIds` del registro, garantizando compatibilidad con Office 2013, 2016/2019/2021/2024 y 365.
- Eliminada la dependencia de patrones en el `UninstallString` para identificar productos, evitando errores y detecciones incorrectas.
- Visualización más precisa y robusta de los datos de cada instalación en consola y logs.
- Refactor en `OfficeInstallation`: el atributo `product` ahora siempre refleja el valor real del registro, sin sobrescritura accidental.

### 🌐 Descarga de ODT: de Scrapy a requests con Retry de urllib3
- Reemplazado el uso de Scrapy para obtener los datos de descarga del Office Deployment Tool (ODT).
- Ahora se utiliza requests con Retry de urllib3 para obtener de forma robusta el HTML desde la página oficial de Microsoft.
- Se implementó validación estricta de dominios y certificados, sanitización de nombres de archivo y control de tamaño de HTML descargado.
- Se añadió caché interna con límite de tamaño para evitar solicitudes duplicadas.
- Eliminada dependencia de Scrapy.

### 👤 Experiencia de usuario y extensibilidad
- Mensajes de usuario claros y coloridos en consola y GUI.
- Validaciones exhaustivas en la GUI y en consola.
- Modularidad que facilita la extensión (nuevas versiones, apps, idiomas, etc.).
- Preparado para integración de tests y CI/CD.
- Estructura profesional, fácil de mantener, escalar y compartir en GitHub.

### 📦 Listo para producción y colaboración
- README.md actualizado: refleja la nueva estructura, dependencias y créditos.
- Recomendaciones para agregar `requirements.txt`, `.gitignore` y carpeta `tests/` para futuras mejoras.
- Compilación recomendada con Nuitka para ejecutable optimizado y seguro.

---

## [4.2] - 2025-05-06  
### 🔄 Migración de motor de descarga: de Playwright a Scrapy  
- Reemplazo completo de Playwright por **Scrapy** para obtener la URL oficial de descarga del Office Deployment Tool (ODT).  
- Extracción automatizada del nombre del archivo y tamaño desde los headers HTTP (`Content-Disposition` y `Content-Length`).  
- Eliminadas las dependencias pesadas (`playwright`, `asyncio`, `websockets`).  

### ⚙️ Nueva lógica de descarga robusta y reanudable  
- Descarga con:
  - Reintentos automáticos (`Retry` con `HTTPAdapter`).
  - Soporte para **descarga reanudable** mediante `Range`.
  - Escritura atómica con `NamedTemporaryFile`, `flush()` y `fsync()` para evitar archivos corruptos o bloqueados por antivirus.
  - Validación final por nombre y tamaño esperado.
- Progreso visual detallado con `tqdm`.  
- Descarga segura y eficiente, incluso ante errores intermitentes de red.

### 🧼 Refactor y mejoras internas  
- Uso extensivo y consistente de `pathlib.Path`.
- Sanitización avanzada de rutas locales y claves del registro (`safe_log_path`, `safe_log_registry_key`).
- Mejoras en el logging para trazabilidad completa del flujo.
- Nueva validación en `ODTManager` para asegurar el estado del archivo descargado antes de ejecutar el instalador.

### 📦 Compilación con Nuitka  
- Primer release con **compilación oficial mediante Nuitka** en lugar de PyInstaller:
  - Menor falsos positivos en antivirus.
  - Ejecutable más rápido, optimizado y difícil de descompilar.
- Comando de compilación utilizado:
  ```bash
  python -m nuitka ManagerOfficeScriptTool.py ^
    --standalone ^
    --enable-plugin=tk-inter ^
    --windows-icon-from-ico=icon.ico ^
    --company-name="Rodri082" ^
    --product-name="ManagerOfficeScriptTool" ^
    --file-version=4.2.0.0 ^
    --product-version=4.2.0.0 ^
    --file-description="Herramienta ManagerOfficeScriptTool" ^
    --copyright="Licencia MIT © 2024 Rodri082" ^
    --windows-uac-admin ^
    --output-dir=build ^
    --msvc=latest ^
    --lto=yes ^
    --report=build/compilacion.xml

---

## [4.0] - 2025-04-14
### Renovación Total del Script
- Refactorización completa en **programación orientada a objetos**: separación en clases (`OfficeManager`, `ODTManager`, `OfficeUninstaller`, `OfficeInstaller`, `OfficeSelectionWindow`, `RegistryReader`).
- Código modular, mantenible y preparado para futuras mejoras.

### Descarga dinámica del ODT
- Implementación de `selenium` y `webdriver-manager` para extraer la **URL oficial desde Microsoft**, asegurando compatibilidad incluso si el portal cambia.
- Verificación de nombre, tamaño y dominio del archivo descargado.

### Registro seguro y robusto
- Nueva clase `RegistryReader` con **caché de valores**, sanitización de rutas del registro y manejo elegante de errores.
- Logging detallado y seguro en `logs/application.log`.

### Interfaz gráfica mejorada
- Rediseño completo de la GUI: selección intuitiva de versión, idioma, arquitectura y apps.
- Posibilidad de excluir aplicaciones antes de la instalación.

### Desinstalación limpia y moderna
- Eliminación definitiva del uso de **SaRA y OffScrub**.
- Desinstalación ahora gestionada exclusivamente con ODT (`setup.exe` + `configuration.xml`).

### Rutas seguras y manejo de archivos temporales
- Uso completo de `pathlib` para rutas más claras y seguras.
- Limpieza de carpetas temporales solo tras confirmación del usuario.

### Compilación
- Compilado con:
  - Python **3.13.3**
  - PyInstaller **6.12.0**
  - pyinstaller-hooks-contrib **2025.2**
- SHA256 del ejecutable publicado y verificado en VirusTotal:  
  `de137932cdc26c726147bb19ac0472b7b163426f020cc6126bf55d3743448c49`

---

## [3.0] - 2025-04-08
### Cambios destacados
- Eliminado completamente el uso de scripts `.vbs` de OffScrub.
- SaRA se utiliza exclusivamente para la desinstalación de Office.
- Nueva gestión de directorios temporales con `tempfile.mkdtemp()`.
- URLs actualizadas para descargas de ODT y SaRA desde fuentes oficiales de Microsoft.
- Unificación de subprocesos con `subprocess.run()` usando `capture_output=True`.
- Mejora en la visibilidad de las ventanas gráficas (`-topmost` en tkinter).
- Refactorización de funciones y limpieza de código redundante.
- Nueva compilación con PyInstaller (`.spec` incluido en el repositorio).
- SHA256 del ejecutable publicado y verificado en VirusTotal.
- Compatible exclusivamente con Windows 10 o superior.

---

## [2.6] - 2025-03-28
### Correcciones y Mejoras
- ComboBox configurado como solo lectura (`readonly`) para prevenir errores de entrada.
- Actualización de dependencias y migración a Python de 64 bits para mejorar compatibilidad y reducir falsos positivos.
- Empaquetado dual: PyInstaller y cx_Freeze.
- SHA256 del ejecutable con PyInstaller: `2a9e4f820d1f69933ccb032d862e9407302c93cf4c8e188da5b3209eec6b1e65`

---

## [2.2] - 2024-12-22
### Descripción
- Primera versión pública combinando PyInstaller (.exe) y cx_Freeze (.zip).
- Eliminación de scripts externos como `Install.ps1`, `RunInstallOffice.bat` y carpeta `Files`.
- Mejora en detección de instalaciones de Office y generación del archivo `configuration.xml`.
- SHA256 del `.exe`: `d1839b07f5ff3be5b09d81508d3c0be09982f7e08190cf318e80faa77dc6d238`
- SHA256 del `.zip`: `8374e544cd599bd8e97310319f4c6295be55640f202db49aa12dd32420dcf4bf`

---

## [1.0-alpha] - 2024-12-16
### Primer release (pre-release)
- Capacidad de ejecutar como `.exe` mediante cx_Freeze.
- Eliminación de `Install.ps1` y `RunInstallOffice.bat`.
- Primer ejecutable autónomo sin necesidad de Python.
- Mejoras básicas de rendimiento y documentación inicial en README.

---
