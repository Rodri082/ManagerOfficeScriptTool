"""
ManagerOfficeScriptTool

Herramienta modular para la gestión avanzada de instalaciones de
Microsoft Office en Windows.

Este script principal orquesta la detección, desinstalación e instalación de
Office de forma interactiva y robusta, apoyándose en módulos y archivos de
configuración especializados para cada tarea:

- `config.yaml`: Archivo central de configuración. Define versiones soportadas,
  aplicaciones disponibles por versión e idiomas.
- `office_manager.py`: Detecta y muestra instalaciones existentes de Office
  consultando el registro de Windows.
- `office_installation.py`: Representa y procesa la información de
  cada instalación detectada.
- `uninstaller.py`: Desinstala versiones de Office utilizando ODT y generación
  dinámica de XML.
- `installer.py`: Instala Office usando setup.exe y configuration.xml, con
  validaciones y manejo de errores.
- `gui.py`: Proporciona una interfaz gráfica moderna para seleccionar versión,
  idioma, arquitectura y apps.
- `odt_manager.py`: Descarga y extrae ODT desde Microsoft, y genera el XML de
  instalación.
- `registry_utils.py`: Utilidades para leer el registro de Windows de
  forma segura y eficiente.
- `utils.py`: Funciones generales para manejo de rutas, logs, diálogos y
  carpetas temporales.

El flujo principal es:
1. Detectar instalaciones existentes de Office.
2. Ofrecer al usuario la opción de desinstalar versiones encontradas.
3. Permitir la configuración e instalación de una nueva versión de Office
   mediante GUI.
4. Gestionar errores y limpiar archivos temporales al finalizar.

Ejecuta este archivo para iniciar la herramienta.
"""

import logging
import multiprocessing

from colorama import Fore, init
from core.office_manager import OfficeManager
from gui.gui import OfficeSelectionWindow
from scripts.installer import OfficeInstaller
from scripts.uninstaller import run_uninstallers
from utils import (
    ask_yes_no,
    clean_temp_folders,
    init_logging,
    logs_folder,
    office_install_dir,
    office_uninstall_dir,
)


def main() -> None:
    """
    Función principal del script.

    Ejecuta el flujo principal de la herramienta de gestión de Office.
    - Detecta instalaciones existentes de Office en el sistema.
    - Ofrece la opción de desinstalar versiones encontradas con ODT.
    - Permite configurar e instalar una nueva versión de Office.
    - Gestiona errores y limpia archivos temporales al finalizar.
    """
    try:
        # Inicializa logs y colorama, y ejecuta el flujo principal
        init_logging(str(logs_folder))
        init(autoreset=True)

        # Pregunta al usuario si desea detectar versiones de Office instaladas
        if ask_yes_no("¿Desea detectar las versiones de Office instaladas?"):
            manager = OfficeManager(show_all=False)
            installations = manager.get_installations()
            manager.display_installations()

            # Si hay varias instalaciones, permite al usuario elegir qué hacer
            if installations:
                if len(installations) == 1:
                    if ask_yes_no(
                        "¿Desea desinstalar la versión encontrada con ODT?"
                    ):
                        run_uninstallers(installations, office_uninstall_dir)
                    else:
                        print(
                            Fore.RED
                            + (
                                "Advertencia: "
                                "Tener múltiples versiones "
                                "puede causar conflictos."
                            )
                        )
                else:
                    # Muestra menu de opciones para desinstalar
                    print(Fore.LIGHTWHITE_EX + "Seleccione una opción:")
                    print(
                        Fore.LIGHTWHITE_EX
                        + ("1 - No desinstalar ninguna versión")
                    )
                    print(
                        Fore.LIGHTWHITE_EX
                        + ("2 - Desinstalar todas las versiones encontradas")
                    )
                    print(
                        Fore.LIGHTWHITE_EX
                        + (
                            "3 - Elegir una versión específica "
                            "para desinstalar"
                        )
                    )
                    opcion = input(
                        Fore.LIGHTWHITE_EX + "Opción (1/2/3): "
                    ).strip()

                    if opcion == "2":
                        run_uninstallers(installations, office_uninstall_dir)

                    elif opcion == "3":
                        print(
                            Fore.LIGHTWHITE_EX
                            + ("Seleccione la versión que desea desinstalar:")
                        )
                        for idx, install in enumerate(installations, 1):
                            print(
                                Fore.LIGHTCYAN_EX + (f"{idx} - {install.name}")
                            )

                        seleccion = input(
                            Fore.LIGHTWHITE_EX
                            + "Ingrese el número de la versión a desinstalar: "
                        ).strip()
                        if seleccion.isdigit() and 1 <= int(seleccion) <= len(
                            installations
                        ):
                            seleccionada = installations[int(seleccion) - 1]
                            print(
                                Fore.LIGHTWHITE_EX
                                + (
                                    "Versión seleccionada: "
                                    f"{seleccionada.name}"
                                    f"({seleccionada.client_culture})"
                                )
                            )
                            run_uninstallers(
                                [seleccionada], office_uninstall_dir
                            )
                        else:
                            print(
                                Fore.YELLOW
                                + (
                                    "Selección inválida. "
                                    "No se realizará ninguna desinstalación."
                                )
                            )
                    else:
                        print(
                            Fore.YELLOW
                            + ("No se realizará ninguna desinstalación.")
                        )
        else:
            print(
                Fore.RED
                + (
                    "Advertencia: "
                    "Si tiene más de una versión de Office instalada, "
                    "podría ocasionar problemas."
                )
            )

        if ask_yes_no("¿Desea proceder con una nueva instalación de Office?"):
            print("Iniciando configuración de instalación de Office...")
            selection_window = OfficeSelectionWindow()
            cancelled = selection_window.show()

            if cancelled:
                print(Fore.YELLOW + "El usuario canceló la instalación.")
            else:
                installer = OfficeInstaller(str(office_install_dir))
                installer.run_setup_commands()
        else:
            print(Fore.YELLOW + "Proceso cancelado por el usuario.")

    except KeyboardInterrupt:
        # Maneja la interrupción manual del usuario (Ctrl+C)
        logging.warning("Ejecución interrumpida por el usuario (Ctrl+C).")
        print(Fore.YELLOW + "\nProceso cancelado manualmente.")

    except Exception as e:
        # Captura y muestra cualquier error inesperado
        logging.exception(
            "Ocurrió un error inesperado en la ejecución principal."
        )
        print(Fore.RED + f"Error crítico: {e}")
        print(
            Fore.YELLOW
            + (
                "Sugerencias: "
                "Verifica tu conexión a internet, "
                "permisos de administrador y espacio en disco."
            )
        )

    finally:
        # Limpia carpetas temporales al finalizar el script
        clean_temp_folders()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
