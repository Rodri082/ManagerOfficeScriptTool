# 📋 CHANGELOG

Historial de cambios del proyecto **ManagerOfficeScriptTool**

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
