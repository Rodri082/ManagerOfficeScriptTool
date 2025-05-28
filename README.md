# ManagerOfficeScriptTool

**Gestión avanzada, modular y automatizada de instalaciones de Microsoft Office en Windows**

---

## 🚀 ¿Qué es ManagerOfficeScriptTool?

`ManagerOfficeScriptTool` es una herramienta profesional, **100% open source**, para detectar, desinstalar e instalar Microsoft Office en equipos Windows.  
Desde la versión **5.0**, el proyecto ha sido completamente **refactorizado y modularizado**, separando la lógica en submódulos claros y configurables, con una interfaz gráfica moderna y soporte para múltiples versiones, arquitecturas e idiomas.

---

## 🗂️ Estructura del Proyecto

```
ManagerOfficeScriptTool/
│
├── manager_office_tool/               # Paquete principal del código fuente
│   ├── __init__.py                    # API pública del paquete
│   │
│   ├── core/                          # Lógica de negocio y gestión de Office
│   │   ├── __init__.py
│   │   ├── office_manager.py          # Detección y visualización de instalaciones
│   │   ├── office_installation.py     # Representación de una instalación detectada
│   │   ├── odt_manager.py             # Descarga y extracción de ODT
│   │   ├── odt_fetcher.py             # Obtención de la URL oficial de ODT
│   │   └── registry_utils.py          # Acceso seguro al registro de Windows
│   │
│   ├── gui/                           # Interfaz gráfica (ttkbootstrap)
│   │   ├── __init__.py
│   │   └── gui.py                     # Ventana de selección y configuración
│   │
│   ├── scripts/                       # Instalador y desinstalador de Office
│   │   ├── __init__.py
│   │   ├── installer.py               # Instalación de Office
│   │   └── uninstaller.py             # Desinstalación de Office
│   │
│   └── utils/                         # Utilidades generales y helpers
│       ├── __init__.py
│       ├── logging_utils.py           # Logging profesional y seguro
│       ├── path_utils.py              # Manejo avanzado y seguro de rutas
│       ├── console_utils.py           # Utilidades para consola
│       └── gui_utils.py               # Helpers para la GUI
│
├── main.py                            # Punto de entrada y orquestador del flujo
├── config.yaml                        # Configuración centralizada (versiones, apps, idiomas)
├── icon.ico                           # Icono del ejecutable
│
├── nuitka_build_instructions.bat      # Script de compilación Nuitka (.exe)
├── requirements.txt                   # Dependencias principales del proyecto
├── requirements-dev.txt               # Dependencias de desarrollo y tipado
│
├── README.md                          # Documentación principal
├── CHANGELOG.md                       # Historial de cambios y versiones
└── LICENSE                            # Licencia del proyecto
```

---

## ✨ Características Principales

- **Arquitectura modular y escalable**: Separación clara de lógica, GUI, utilidades y scripts.
- **Configuración centralizada**: Todas las versiones, canales, apps e idiomas en [`config.yaml`](config.yaml).
- **Detección avanzada** de instalaciones de Office (todas las versiones soportadas).
- **Desinstalación limpia y segura** usando Office Deployment Tool (ODT) oficial.
- **Instalación automatizada** con generación dinámica de `configuration.xml`.
- **Interfaz gráfica moderna** (ttkbootstrap) para seleccionar versión, apps, arquitectura e idioma.
- **Descarga directa y robusta** del ODT desde Microsoft, con reintentos y validación.
- **Logging profesional y seguro**: Rutas y claves anonimizadas, sin exponer datos sensibles.
- **Cumplimiento estricto de PEP8** y uso de herramientas como Black, isort, flake8 y mypy.
- **Preparado para integración continua y testing**.
- **No instala software de terceros ni envía datos personales**.

---

## 🖥️ Requisitos

- Windows 10 o superior.
- Python 3.13+ (para desarrollo) o ejecutable standalone.
- Acceso a internet para descargar ODT y actualizaciones.

---

## ⚙️ Instalación y Uso

1. **Clona el repositorio** o descarga la última release.
2. Instala las dependencias:
   ```sh
   pip install -r requirements.txt
   ```
3. (Opcional, recomendado para desarrollo y análisis estático) Instala las dependencias de desarrollo:
   ```sh
   pip install -r requirements-dev.txt
   ```
4. Ejecuta el script principal:
   ```sh
   python main.py
   ```
5. Sigue las instrucciones en consola y/o GUI para detectar, desinstalar e instalar Office.

---

## 🧩 Configuración

Edita [`config.yaml`](config.yaml) para:
- Agregar o quitar versiones soportadas.
- Definir canales y product IDs.
- Personalizar las aplicaciones disponibles por versión.
- Añadir nuevos idiomas.

---

## 🛡️ Seguridad y Transparencia

- **100% open source**.
- Solo descarga herramientas oficiales de Microsoft.
- No instala software de terceros ni envía datos personales.
- Logging seguro y rutas anonimizadas.
- Acceso al registro de Windows de forma segura y eficiente.
- No se aceptan rutas arbitrarias del usuario: todas las rutas temporales y de trabajo se generan internamente.

---

## 📝 Ejemplo de Flujo

1. **Detección**: El script detecta todas las instalaciones de Office.
2. **Desinstalación**: Puedes elegir desinstalar todas, ninguna o una versión específica.
3. **Instalación**: Selecciona versión, arquitectura, idioma y apps desde la GUI.
4. **Ejecución**: Se genera el XML y se lanza la instalación con ODT.
5. **Limpieza**: Elimina carpetas temporales y muestra logs detallados.

---

## 🛠️ Desarrollo y Contribución

- Cumple con PEP8 y buenas prácticas.
- Usa Black, isort, flake8 y mypy para mantener la calidad.
- Estructura lista para agregar tests (`tests/`).
- Pull requests y sugerencias son bienvenidas.
- Modularidad y tipado para facilitar la extensión y el testing.

---

## 📦 Compilación a ejecutable

Se recomienda usar [Nuitka](https://nuitka.net/) para compilar el proyecto a `.exe`:

```sh
cmd /c nuitka_build_instructions.bat
```
o en PowerShell:
```powershell
.\nuitka_build_instructions.bat
```

Consulta el [`CHANGELOG.md`](CHANGELOG.md) para detalles de versiones y mejoras.

---

## 📄 Licencia

Este proyecto está licenciado bajo la [Licencia MIT](LICENSE).

---

## 🙌 Créditos

- [Office Deployment Tool (ODT)](http://aka.ms/ODT)
- [ttkbootstrap](https://ttkbootstrap.readthedocs.io/)
- [colorama](https://pypi.org/project/colorama/)

---

## 📣 Estado del Proyecto

**Activo y estable.**  
Listo para producción, colaboración y futuras mejoras.

---
