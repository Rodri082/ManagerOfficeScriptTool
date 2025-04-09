
# 🆕 Versión 4.0 - Cambios y Mejoras

## ✅ Novedades principales

- **Desinstalación 100% vía SaRA**
  - Se eliminó por completo el uso de scripts `.vbs` de OffScrub.
  - Esto reduce la complejidad del código y evita falsos positivos de antivirus generados por esos scripts.

- **Gestión centralizada de archivos temporales**
  - Uso de variables globales (`temp_dir`, `logs_folder`, `install_folder`, `uninstall_folder`) con `tempfile.mkdtemp()`.
  - Mejora la limpieza de archivos, evita conflictos y mantiene todo contenido en un entorno seguro.

- **Actualización de fuentes oficiales**
  - Nuevos endpoints para descargar ODT y SaRA directamente desde Microsoft.
  - Se valida que los archivos descargados sean los correctos antes de continuar.

- **Subprocesos unificados y seguros**
  - Todas las herramientas externas usan `subprocess.run()` con `capture_output=True` y `text=True`.
  - Permite capturar y mostrar errores con mayor claridad y registrar logs más útiles.

- **Mejoras gráficas y de experiencia**
  - Uso uniforme de `tkinter` con ventanas siempre visibles (`-topmost`).
  - Mejora en la interacción con el EULA de SaRA y generación automática del archivo `configuration.xml`.

- **Limpieza más robusta**
  - `remove_temp_folder()` ahora borra únicamente el directorio temporal generado por el script.
  - Se eliminaron excepciones duplicadas y código redundante.

- **Refactorización general**
  - Código más organizado y simplificado.
  - Flujo más intuitivo para la instalación y desinstalación de Office.

---

## 🧪 Compilación reproducible

Este ejecutable fue generado con:

- **Python 3.13.2**
- **PyInstaller 6.12.0**
- **pyinstaller-hooks-contrib 2025.2**

Archivo de compilación incluido: [`ManagerOfficeScriptTool.spec`](./ManagerOfficeScriptTool.spec)

```bash
pyinstaller ManagerOfficeScriptTool.spec
```

---

## 🔐 Seguridad y Transparencia

Este proyecto es 100% open source, no realiza conexiones ocultas ni recopila datos. Todas las descargas provienen de fuentes oficiales de Microsoft.

🔗 Ejecutable verificado en VirusTotal:  
[ManagerOfficeScriptTool.exe](https://www.virustotal.com/gui/file/54cccdd26195d448b516800c0e708d43b2ec392ddb440cec827d2fa12fd4edd4)  
`SHA256: 54cccdd26195d448b516800c0e708d43b2ec392ddb440cec827d2fa12fd4edd4`

---

---

## 📦 Bibliotecas utilizadas

Este script está desarrollado en Python y hace uso de las siguientes bibliotecas estándar y externas:

### Bibliotecas estándar de Python:
- `os`, `sys`, `shutil`, `zipfile`, `subprocess`, `threading`, `datetime`, `tempfile`, `platform`, `winreg`
- `typing.List`
- `urllib.parse`

### Bibliotecas externas:
- [`requests`](https://pypi.org/project/requests/)
- [`tqdm`](https://pypi.org/project/tqdm/)
- [`colorama`](https://pypi.org/project/colorama/)
- `tkinter` (interfaz gráfica incluida en Python estándar para Windows)

> Nota: Todas estas bibliotecas están incluidas en el ejecutable generado con PyInstaller, por lo que no es necesario instalarlas por separado para usar el `.exe`.

---
