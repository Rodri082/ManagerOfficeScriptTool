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
### Última versión anterior conocida
- Uso combinado de `cx_Freeze` y versiones anteriores de PyInstaller.
- Desinstalación basada en scripts `.vbs` de OffScrub.
- Descargas automáticas de herramientas desde fuentes oficiales.
- Interfaz gráfica básica para configuración de instalación.

---

*Nota: versiones intermedias entre 2.6 y 4.0 no fueron oficialmente publicadas.*