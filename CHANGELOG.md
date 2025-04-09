# 📋 CHANGELOG

Historial de cambios del proyecto **ManagerOfficeScriptTool**

---

## [4.0] - 2025-04-09
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

## [2.6] - 2024-XX-XX
### Correcciones y Mejoras
- ComboBox configurado como solo lectura (`readonly`) para prevenir errores de entrada.
- Actualización de dependencias y migración a Python de 64 bits para mejorar compatibilidad y reducir falsos positivos.
- Empaquetado dual: PyInstaller y cx_Freeze.
- SHA256 del ejecutable con PyInstaller: `2a9e4f820d1f69933ccb032d862e9407302c93cf4c8e188da5b3209eec6b1e65`

---

## [2.2] - 2024-XX-XX
### Descripción
- Primera versión pública combinando PyInstaller (.exe) y cx_Freeze (.zip).
- Eliminación de scripts externos como `Install.ps1`, `RunInstallOffice.bat` y carpeta `Files`.
- Mejora en detección de instalaciones de Office y generación del archivo `configuration.xml`.
- SHA256 del `.exe`: `d1839b07f5ff3be5b09d81508d3c0be09982f7e08190cf318e80faa77dc6d238`
- SHA256 del `.zip`: `8374e544cd599bd8e97310319f4c6295be55640f202db49aa12dd32420dcf4bf`

---

## [1.0-alpha] - 2024-XX-XX
### Primer release (pre-release)
- Capacidad de ejecutar como `.exe` mediante cx_Freeze.
- Eliminación de `Install.ps1` y `RunInstallOffice.bat`.
- Primer ejecutable autónomo sin necesidad de Python.
- Mejoras básicas de rendimiento y documentación inicial en README.

---
