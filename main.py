"""
ManagerOfficeScriptTool

Herramienta modular para la gestión avanzada de instalaciones de
Microsoft Office en Windows.

Este script principal orquesta la detección, desinstalación e instalación de
Office de forma interactiva y robusta, apoyándose en módulos y archivos de
configuración especializados para cada tarea.

El flujo principal es:
1. Detectar instalaciones existentes de Office.
2. Ofrecer al usuario la opción de desinstalar versiones encontradas.
3. Permitir la configuración e instalación de una nueva versión de Office
mediante GUI.
4. Gestionar errores y limpiar archivos temporales al finalizar.

Notas de seguridad:
- Las rutas de trabajo y temporales se generan y controlan internamente por el
    propio flujo del programa.
- No se permite que el usuario final defina rutas arbitrarias para
    instalaciones o descargas, mitigando riesgos de path-injection.
- El uso de `msvcrt` para capturar interrupciones manuales (Ctrl+C) permite
    una salida controlada y limpia del script.
- El manejo de excepciones asegura que cualquier error inesperado se registre
    adecuadamente y no interrumpa el flujo sin control.

Ejecuta este archivo para iniciar la herramienta.
"""

import logging
import msvcrt
import os
from pathlib import Path
from typing import Any, Dict

import colorama
from colorama import Fore, Style
from manager_office_tool import (
    OfficeInstaller,
    OfficeManager,
    OfficeSelectionWindow,
    ask_menu_option,
    ask_multiple_valid_indices,
    ask_single_valid_index,
    ask_yes_no,
    clean_folders,
    ensure_subfolder,
    get_temp_dir,
    init_logging,
    run_uninstallers,
)


def prepare_environment() -> Dict[str, Any]:
    """
    Prepara el entorno mínimo necesario y retorna las rutas relevantes.

    Returns:
        dict: Diccionario con las rutas preparadas.
    """
    try:
        temp_dir = get_temp_dir()
        logs_folder = ensure_subfolder(temp_dir, "logs")

        init_logging(str(logs_folder))
        colorama.init()
        os.environ["QT_LOGGING_RULES"] = "qt.qpa.window=false"

        return {
            "temp_dir": temp_dir,
            "logs_folder": logs_folder,
        }
    except Exception as e:
        msg = f"[CONSOLE] Error preparando entorno: {e}"
        logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
        return {}


def main() -> None:
    """
    Función principal del script.

    Ejecuta el flujo principal de la herramienta de gestión de Office.
    - Detecta instalaciones existentes de Office en el sistema.
    - Ofrece la opción de desinstalar versiones encontradas con ODT.
    - Permite configurar e instalar una nueva versión de Office.
    - Gestiona errores y limpia archivos temporales al finalizar.
    - Maneja interrupciones manuales (Ctrl+C) de forma controlada.
    - Utiliza una interfaz gráfica para la selección de versiones y opciones.
    - Registra toda la actividad en un archivo de log para auditoría y
    depuración.

    Seguridad:
    - Las rutas de trabajo y temporales se generan internamente y no provienen
        de entrada de usuario.
    - No se permite que el usuario final defina rutas arbitrarias para
        instalaciones o descargas, mitigando riesgos de path-injection.
    """
    office_install_dir: Path | None = None
    office_uninstall_dir: Path | None = None

    env = prepare_environment()
    try:
        temp_dir = env["temp_dir"]
        logging.info(
            f"{Fore.GREEN}Entorno preparado correctamente.{Style.RESET_ALL}"
        )

        # Fuerza el vaciado inmediato de buffers de consola y archivo
        for handler in logging.getLogger().handlers:
            handler.flush()

    except KeyError:
        print(
            "No se pudo preparar el entorno temporal. Presione cualquier "
            "tecla para salir..."
        )
        msvcrt.getch()
        return

    try:
        # Pregunta al usuario si desea detectar versiones de Office instaladas
        if ask_yes_no(
            f"{Fore.LIGHTCYAN_EX}"
            "¿Desea detectar las versiones de Office instaladas? (S/N):"
            f"{Style.RESET_ALL}"
        ):
            manager = OfficeManager(show_all=False)
            installations = manager.get_installations()
            manager.display_installations()

            # Si hay varias instalaciones, permite al usuario elegir qué hacer
            if installations:
                if len(installations) == 1:
                    if ask_yes_no(
                        f"{Fore.LIGHTCYAN_EX}"
                        "¿Desea desinstalar "
                        "la versión encontrada con ODT? (S/N):"
                        f"{Style.RESET_ALL}"
                    ):
                        office_uninstall_dir = ensure_subfolder(
                            temp_dir, "UninstallOfficeFiles"
                        )
                        run_uninstallers(installations, office_uninstall_dir)
                    else:
                        logging.info(
                            f"{Fore.RED}"
                            "Advertencia: Tener múltiples versiones "
                            "puede causar conflictos."
                            f"{Style.RESET_ALL}"
                        )
                else:
                    # Muestra menu de opciones para desinstalar
                    logging.info(
                        f"{Fore.LIGHTCYAN_EX}"
                        "Seleccione una opción:"
                        f"{Style.RESET_ALL}"
                    )
                    logging.info(
                        f"{Fore.LIGHTWHITE_EX}"
                        "1 - No desinstalar ninguna versión"
                        f"{Style.RESET_ALL}"
                    )
                    logging.info(
                        f"{Fore.LIGHTWHITE_EX}"
                        "2 - Desinstalar todas las versiones encontradas"
                        f"{Style.RESET_ALL}"
                    )
                    logging.info(
                        f"{Fore.LIGHTWHITE_EX}"
                        "3 - Elegir una versión específica para desinstalar"
                        f"{Style.RESET_ALL}"
                    )
                    logging.info(
                        f"{Fore.LIGHTWHITE_EX}"
                        "4 - Elegir múltiples versiones para desinstalar"
                        f"{Style.RESET_ALL}"
                    )

                    opcion = ask_menu_option({"1", "2", "3", "4"}, "Opción")

                    if opcion == "1":
                        logging.info(
                            f"{Fore.YELLOW}"
                            "No se realizará ninguna desinstalación."
                            f"{Style.RESET_ALL}"
                        )

                    elif opcion == "2":
                        office_uninstall_dir = ensure_subfolder(
                            temp_dir, "UninstallOfficeFiles"
                        )
                        run_uninstallers(installations, office_uninstall_dir)

                    elif opcion == "3":
                        logging.info(
                            f"{Fore.LIGHTCYAN_EX}"
                            "Seleccione la versión que desea desinstalar:"
                            f"{Style.RESET_ALL}"
                        )
                        for idx, install in enumerate(installations, 1):
                            logging.info(
                                f"{Fore.LIGHTCYAN_EX}"
                                f"{idx} - {install.name}"
                                f"{Style.RESET_ALL}"
                            )

                        seleccion_valida = ask_single_valid_index(
                            len(installations)
                        )

                        if seleccion_valida is None:
                            logging.info(
                                f"{Fore.YELLOW}"
                                "No se realizará ninguna desinstalación."
                                f"{Style.RESET_ALL}"
                            )
                        else:
                            seleccionada = installations[seleccion_valida - 1]
                            logging.info(
                                f"{Fore.LIGHTWHITE_EX}"
                                "Versión seleccionada: "
                                f"{seleccionada.name} "
                                f"({seleccionada.client_culture})"
                                f"{Style.RESET_ALL}"
                            )
                            office_uninstall_dir = ensure_subfolder(
                                temp_dir, "UninstallOfficeFiles"
                            )
                            run_uninstallers(
                                [seleccionada], office_uninstall_dir
                            )
                    elif opcion == "4":
                        logging.info(
                            f"{Fore.LIGHTCYAN_EX}"
                            "Seleccione las versiones que desea desinstalar "
                            "(números separados por coma):"
                            f"{Style.RESET_ALL}"
                        )
                        for idx, install in enumerate(installations, 1):
                            logging.info(
                                f"{Fore.LIGHTCYAN_EX}"
                                f"{idx} - {install.name} "
                                f"({install.client_culture})"
                                f"{Style.RESET_ALL}"
                            )

                        indices_validos = ask_multiple_valid_indices(
                            len(installations)
                        )

                        if not indices_validos:
                            logging.info(
                                f"{Fore.YELLOW}"
                                "No se realizará ninguna desinstalación."
                                f"{Style.RESET_ALL}"
                            )
                        else:
                            seleccionadas = [
                                installations[i - 1] for i in indices_validos
                            ]
                            logging.info(
                                f"{Fore.LIGHTWHITE_EX}"
                                "Versiones seleccionadas para desinstalar:"
                                f"{Style.RESET_ALL}"
                            )
                            for s in seleccionadas:
                                logging.info(
                                    f"{Fore.LIGHTWHITE_EX}"
                                    f"- {s.name} ({s.client_culture})"
                                    f"{Style.RESET_ALL}"
                                )

                            office_uninstall_dir = ensure_subfolder(
                                temp_dir, "UninstallOfficeFiles"
                            )
                            run_uninstallers(
                                seleccionadas, office_uninstall_dir
                            )

                    else:
                        logging.info(
                            f"{Fore.YELLOW}"
                            "No se realizará ninguna desinstalación."
                            f"{Style.RESET_ALL}"
                        )
        else:
            logging.info(
                f"{Fore.RED}"
                "Advertencia: Si tiene más de una versión de Office "
                "instalada, podría ocasionar problemas."
                f"{Style.RESET_ALL}"
            )

        if ask_yes_no(
            f"{Fore.LIGHTCYAN_EX}"
            "¿Desea proceder con una nueva instalación de Office? (S/N): "
            f"{Style.RESET_ALL}"
        ):
            logging.info(
                f"{Fore.GREEN}"
                "Iniciando configuración de instalación de Office..."
                f"{Style.RESET_ALL}"
            )
            office_install_dir = ensure_subfolder(
                temp_dir, "InstallOfficeFiles"
            )
            selection_window = OfficeSelectionWindow(office_install_dir)
            install_subdir_path, selected_version, selected_language_id = (
                selection_window.show()
            )

            if not install_subdir_path:
                logging.info(
                    f"{Fore.YELLOW}"
                    "El usuario canceló la instalación."
                    f"{Style.RESET_ALL}"
                )
            elif selected_version is None:
                logging.error(
                    f"{Fore.RED}"
                    "No se pudo determinar la versión seleccionada de Office."
                    f"{Style.RESET_ALL}"
                )
            else:
                installer = OfficeInstaller(
                    str(install_subdir_path),
                    selected_version,
                    str(selected_language_id),
                )
                installer.run_setup_commands()
        else:
            logging.info(
                f"{Fore.YELLOW}"
                "Proceso cancelado por el usuario."
                f"{Style.RESET_ALL}"
            )

    except KeyboardInterrupt:
        # Maneja la interrupción manual del usuario (Ctrl+C)
        msg = "[CONSOLE] Ejecución interrumpida por el usuario (Ctrl+C)."
        logging.warning(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")

    except Exception as e:
        # Captura y muestra cualquier error inesperado
        msg = (
            "[CONSOLE] Ocurrió un error inesperado en la ejecución "
            f"principal. {e}"
        )
        logging.exception(f"{Fore.RED}{msg}{Style.RESET_ALL}")

    finally:
        # Solo limpia si las carpetas existen
        folders_to_clean: list[Path] = []
        if office_install_dir is not None:
            folders_to_clean.append(office_install_dir)
        if office_uninstall_dir is not None:
            folders_to_clean.append(office_uninstall_dir)
        if folders_to_clean:
            if ask_yes_no(
                f"{Fore.LIGHTCYAN_EX}"
                "¿Deseas eliminar las carpetas temporales creadas por el script? (S/N): "
                f"{Style.RESET_ALL}"
            ):
                clean_folders(folders_to_clean)
        print(
            f"{Fore.LIGHTWHITE_EX}"
            "Presione cualquier tecla para salir..."
            f"{Style.RESET_ALL}"
        )
        msvcrt.getch()


if __name__ == "__main__":
    main()
