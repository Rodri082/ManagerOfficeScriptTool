"""
gui.py

Módulo encargado de la interfaz gráfica para la selección y configuración
de la instalación de Microsoft Office.
Permite al usuario elegir versión, arquitectura, idioma y aplicaciones, y
genera el archivo configuration.xml para Office Deployment Tool (ODT).
Utiliza ttkbootstrap para una GUI moderna y mensajes claros.
"""

import logging
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Tuple

import ttkbootstrap as tb
import yaml
from colorama import Fore, Style
from manager_office_tool.core import ODTManager
from manager_office_tool.utils import (
    center_window,
    get_data_path,
    safe_log_path,
)
from ttkbootstrap.dialogs import Messagebox


class OfficeSelectionWindow:
    """
    Ventana gráfica para seleccionar opciones de instalación de
    Microsoft Office.

    Permite al usuario escoger la versión, arquitectura, idioma y aplicaciones
    a instalar, generando luego un archivo de configuración XML compatible con
    Office Deployment Tool (ODT).
    """

    def __init__(self, office_install_dir: Path):
        """
        Inicializa la interfaz gráfica del selector de versiones de Office.
        """
        self.office_install_dir = office_install_dir

        with open(get_data_path("config.yaml"), encoding="utf-8") as f:
            config = yaml.safe_load(f)

        self.all_apps = config["office_apps"]
        self.versiones = config["office_versions"]
        self.languages = config["languages"]

        self.root = tb.Window(themename="darkly")
        self.app_vars: dict[str, tb.BooleanVar] = {}
        self.app_checkbuttons: list[tb.Checkbutton] = []
        self.cancelled: bool = False
        self.install_subdir_path: Path | None = None

    def on_closing(self) -> None:
        """
        Manejador para el evento de cierre de la ventana.
        Cancela la instalación y limpia los temporales.
        """
        if getattr(self, "_already_closed", False):
            return

        self._already_closed = True
        self.cancelled = True
        self.root.destroy()

    def update_apps(self, event=None) -> None:
        """
        Actualiza la lista de aplicaciones mostradas según la versión
        seleccionada.
        """
        selected_version = self.combo_version.get()
        available_apps = self.all_apps.get(selected_version, [])

        # Limpia los checkbuttons anteriores antes de mostrar los nuevos
        for widget in self.app_checkbuttons:
            widget.grid_forget()
        self.app_vars.clear()
        self.app_checkbuttons.clear()

        for i, app in enumerate(available_apps):
            var = tb.BooleanVar()
            self.app_vars[app] = var
            cb = tb.Checkbutton(self.frame_apps, text=app, variable=var)
            cb.grid(row=i, column=0, sticky="w", pady=3)
            self.app_checkbuttons.append(cb)

        self.root.update_idletasks()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()
        current_w = self.root.winfo_width()
        current_h = self.root.winfo_height()
        if w != current_w or h != current_h:
            self.root.geometry(f"{w}x{h}")

    def generate_configuration(
        self,
        selected_version: str,
        bits: str,
        selected_language_name: str,
        remove_msi: bool,
        selected_apps: list[str],
    ) -> str | None:
        """
        Genera el archivo configuration.xml con las opciones seleccionadas por
        el usuario.

        Returns:
            str | None: Ruta al archivo generado o None si hubo error.
        """
        self.root.destroy()
        familia = "2013" if "2013" in selected_version else "modern"
        install_subdir = Path(self.office_install_dir) / f"OfficeODT_{familia}"
        install_subdir.mkdir(parents=True, exist_ok=True)

        odt_manager = ODTManager(str(install_subdir))
        logging.info(
            f"{Fore.CYAN}"
            "Usando carpeta de instalación: "
            f"{safe_log_path(install_subdir)}"
            f"{Style.RESET_ALL}"
        )

        if not odt_manager.download_and_extract(selected_version):
            msg = "[CONSOLE] Error. No se pudo descargar y extraer ODT."
            logging.error(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            return None

        if selected_version not in self.versiones:
            Messagebox.show_error(
                "Versión seleccionada no válida.",
                title="Error",
                parent=self.root,
            )
            return None

        language_id = self.languages.get(selected_language_name)
        if not language_id:
            Messagebox.show_error(
                "Idioma seleccionado no válido.",
                title="Error",
                parent=self.root,
            )
            return None

        office_product_id = self.versiones[selected_version]["product_id"]

        # Determina los IDs de producto para Visio y Project según la versión
        # de Office seleccionada
        if "O365" in office_product_id:
            visio_id = "VisioProRetail"
            project_id = "ProjectProRetail"
        elif office_product_id.startswith("Standard"):
            suffix = office_product_id.removeprefix("Standard")
            visio_id = f"VisioStd{suffix}"
            project_id = f"ProjectStd{suffix}"
        elif office_product_id.startswith("ProPlus"):
            suffix = office_product_id.removeprefix("ProPlus")
            visio_id = f"VisioPro{suffix}"
            project_id = f"ProjectPro{suffix}"
        else:
            visio_id = "VisioProRetail"
            project_id = "ProjectProRetail"

        PRODUCT_IDS = {
            "Office": office_product_id,
            "Visio": visio_id,
            "Project": project_id,
        }

        # Excluye las aplicaciones seleccionadas de la lista de disponibles
        available_apps = set(self.all_apps.get(selected_version, []))
        selected_apps_set = set(selected_apps)
        excluded_apps = available_apps - selected_apps_set

        # Determina los productos a instalar según la selección del usuario
        products = ["Office"]
        if "Visio" in selected_apps:
            products.append("Visio")
        if "Project" in selected_apps:
            products.append("Project")

        configuration = ET.Element("Configuration")
        add = ET.SubElement(
            configuration,
            "Add",
            {
                "OfficeClientEdition": bits,
                "Channel": self.versiones[selected_version]["channel"],
            },
        )

        for product in products:
            product_elem = ET.SubElement(
                add, "Product", {"ID": PRODUCT_IDS[product]}
            )
            ET.SubElement(product_elem, "Language", {"ID": language_id})
            if product == "Office":
                for app in sorted(excluded_apps):
                    ET.SubElement(product_elem, "ExcludeApp", {"ID": app})

        ET.SubElement(
            configuration,
            "Property",
            {"Name": "FORCEAPPSHUTDOWN", "Value": "TRUE"},
        )
        ET.SubElement(
            configuration, "Property", {"Name": "AUTOACTIVATE", "Value": "1"}
        )
        ET.SubElement(configuration, "Updates", {"Enabled": "TRUE"})

        if remove_msi:
            ET.SubElement(configuration, "RemoveMSI")

        ET.SubElement(
            configuration, "Display", {"Level": "Full", "AcceptEULA": "TRUE"}
        )

        config_file_path = install_subdir / "configuration.xml"
        try:
            rough_string = ET.tostring(configuration, encoding="utf-8")
            reparsed = minidom.parseString(rough_string)
            pretty_xml = reparsed.toprettyxml(indent="  ", encoding="utf-8")

            config_file_path.write_bytes(pretty_xml)

            logging.debug(
                "[*] Archivo de configuración generado exitosamente en: "
                f"{safe_log_path(config_file_path)}"
            )
            self.install_subdir_path = install_subdir
            self.selected_version = selected_version
            self.selected_language_id = selected_language_name

            return str(install_subdir)
        except Exception as e:
            msg = "[CONSOLE] Error al escribir configuration.xml"
            logging.exception(f"{Fore.RED}{msg}{Style.RESET_ALL}")
            Messagebox.show_error(
                f"No se pudo guardar el archivo:\n{e}\n\n"
                "Verifica que tienes permisos de escritura "
                "en la carpeta de destino.",
                title="Error",
                parent=self.root,
            )
            return None

    def install_office(self) -> None:
        """
        Valida los campos seleccionados por el usuario y genera el archivo
        de configuración.
        """
        selected_version = self.combo_version.get()
        bits = self.arch_var.get()
        selected_language_name = self.combo_language.get()
        remove_msi = self.remove_msi_var.get()
        selected_apps = [
            app for app, var in self.app_vars.items() if var.get()
        ]

        if self.visio_var.get():
            selected_apps.append("Visio")
        if self.project_var.get():
            selected_apps.append("Project")

        # Verifica que el usuario haya seleccionado al menos una aplicación
        if not selected_apps:
            Messagebox.show_warning(
                "Por favor selecciona al menos una aplicación a instalar.",
                title="Advertencia",
                parent=self.root,
            )
            return
        if not selected_version or not bits or not selected_language_name:
            Messagebox.show_warning(
                "Por favor selecciona una versión, arquitectura y lenguaje.",
                title="Advertencia",
                parent=self.root,
            )
            return
        if selected_version not in self.versiones:
            Messagebox.show_error(
                "No se encontró la configuración de la versión seleccionada.",
                title="Error",
                parent=self.root,
            )
            return

        self.generate_configuration(
            selected_version,
            bits,
            selected_language_name,
            remove_msi,
            selected_apps,
        )

    def show(self) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        Muestra la ventana de selección de Office con todos los
        componentes gráficos.

        Returns:
            tuple str | None, str | None: Ruta de instalación y versión
            seleccionada, o (None, None) si se cancela.
        """
        self.root.title("Instalador de Microsoft Office")
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", lambda: self.on_closing())
        self.root.style.theme_use("darkly")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        frame = tb.Frame(self.root, padding=24)
        frame.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)

        tb.Label(frame, text="Versión de Office:").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        self.combo_version = tb.Combobox(
            frame,
            width=40,
            state="readonly",
            values=list(self.all_apps.keys()),
        )
        self.combo_version.set(list(self.all_apps.keys())[0])
        self.combo_version.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.combo_version.bind("<<ComboboxSelected>>", self.update_apps)

        addons_frame = tb.Frame(frame)
        addons_frame.grid(row=2, column=0, sticky="w", pady=(0, 12))
        self.visio_var = tb.BooleanVar(value=False)
        self.project_var = tb.BooleanVar(value=False)

        options_frame = tb.Frame(frame)
        options_frame.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        tb.Label(options_frame, text="Arquitectura:").grid(
            row=0, column=0, sticky="w"
        )
        self.arch_var = tb.StringVar(value="64")
        tb.Radiobutton(
            options_frame, text="x86", variable=self.arch_var, value="32"
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        tb.Radiobutton(
            options_frame, text="x64", variable=self.arch_var, value="64"
        ).grid(row=0, column=1, sticky="w", padx=(60, 24))
        tb.Label(options_frame, text="Idioma:").grid(
            row=0, column=2, sticky="w"
        )
        self.combo_language = tb.Combobox(
            options_frame,
            width=20,
            state="readonly",
            values=sorted(self.languages.keys()),
            height=15,
        )
        self.combo_language.set(list(self.languages.keys())[0])
        self.combo_language.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self.remove_msi_var = tb.BooleanVar()
        tb.Checkbutton(
            frame,
            text="Eliminar versiones MSI (RemoveMSI)",
            variable=self.remove_msi_var,
        ).grid(row=4, column=0, sticky="w", pady=(0, 12))

        tb.Label(frame, text="Aplicaciones a instalar:").grid(
            row=5, column=0, sticky="w"
        )
        self.frame_apps = tb.Frame(frame)
        self.frame_apps.grid(row=6, column=0, sticky="ew", pady=(0, 16))
        self.update_apps()

        # Añade los checkboxes de Visio y Project aquí, dentro de frame_apps
        self.visio_var = tb.BooleanVar(value=False)
        self.project_var = tb.BooleanVar(value=False)
        self.visio_check = tb.Checkbutton(
            self.frame_apps, text="Visio", variable=self.visio_var
        )
        self.visio_check.grid(
            row=100, column=0, sticky="w", pady=3
        )  # Usa un row alto para que quede al final
        self.project_check = tb.Checkbutton(
            self.frame_apps, text="Project", variable=self.project_var
        )
        self.project_check.grid(row=101, column=0, sticky="w", pady=3)

        button_frame = tb.Frame(frame)
        button_frame.grid(row=7, column=0, pady=(8, 0))

        tb.Button(
            button_frame,
            text="Instalar",
            command=self.install_office,
            width=12,
        ).pack(side="left", padx=(0, 10))

        tb.Button(
            button_frame,
            text="Cancelar",
            command=self.on_closing,
            width=12,
            bootstyle="secondary",  # type: ignore
        ).pack(side="left")

        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        center_window(self.root, w, h)
        self.root.minsize(
            self.root.winfo_reqwidth(), self.root.winfo_reqheight()
        )
        self.root.resizable(False, False)
        self.root.mainloop()
        return (
            (
                str(self.install_subdir_path),
                getattr(self, "selected_version", None),
                getattr(self, "selected_language_id", None),
            )
            if not self.cancelled
            else (None, None, None)
        )
