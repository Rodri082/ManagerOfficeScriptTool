"""
gui.py

Módulo encargado de la interfaz gráfica para la selección y configuración
de la instalación de Microsoft Office.
Permite al usuario elegir versión, arquitectura, idioma y aplicaciones, y
genera el archivo configuration.xml para Office Deployment Tool (ODT).
Utiliza ttkbootstrap para una GUI moderna y mensajes claros.
"""

import logging
from pathlib import Path

import ttkbootstrap as tb
import yaml
from colorama import Fore
from core.odt_manager import ODTManager
from ttkbootstrap.dialogs import Messagebox
from utils import center_window, get_data_path, safe_log_path


class OfficeSelectionWindow:
    """
    Ventana gráfica para seleccionar opciones de instalación de Microsoft
    Office.

    Permite al usuario escoger la versión, arquitectura, idioma, y
    aplicaciones a instalar, generando luego un archivo de configuración
    XML compatible con Office Deployment Tool (ODT).
    """

    def __init__(self, office_install_dir: Path):
        """
        Inicializa la interfaz gráfica del selector de versiones de Office.

        Atributos principales:
            all_apps (dict): Aplicaciones disponibles por versión.
            versiones (dict): Configuración de cada versión de Office.
            languages (dict): Idiomas disponibles.
            root (tb.Window): Ventana principal de la GUI.
            app_vars (dict): Variables de estado para los checkboxes de apps.
            app_checkbuttons (list): Lista de checkbuttons de apps.
            cancelled (bool): Indica si el usuario canceló la instalación.
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

    def on_closing(self) -> None:
        """
        Manejador para el evento de cierre de la ventana.
        Cancela la instalación y limpia los temporales.
        """
        if getattr(self, "_already_closed", False):
            return
        self._already_closed = True

        Messagebox.show_warning(
            "Instalación de Office cancelada...",
            title="Advertencia",
            parent=self.root,
        )
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
        Genera el archivo configuration.xml con las opciones seleccionadas
        por el usuario.

        Args:
            selected_version (str): Versión de Office seleccionada.
            bits (str): Arquitectura seleccionada ("32" o "64").
            selected_language_name (str): Nombre del idioma seleccionado.
            remove_msi (bool): Si se debe eliminar versiones antiguas.
            selected_apps (list[str]): Lista de apps seleccionadas.

        Returns:
            str | None: Ruta al archivo generado o None si hubo error.
        """

        odt_manager = ODTManager(str(self.office_install_dir))
        if not odt_manager.download_and_extract(selected_version):
            Messagebox.show_error(
                "No se pudo descargar y extraer ODT.",
                title="Error",
                parent=self.root,
            )
            return None

        self.root.destroy()

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

        available_apps = set(self.all_apps.get(selected_version, []))
        selected_apps_set = set(selected_apps)
        # Excluye las aplicaciones no seleccionadas usando
        # el bloque <ExcludeApp>
        excluded_apps = available_apps - selected_apps_set
        exclude_apps_block = "\n".join(
            f'      <ExcludeApp ID="{app}" />' for app in sorted(excluded_apps)
        )

        # Determina los productos a instalar según la selección del usuario
        products = ["Office"]
        if "Visio" in selected_apps:
            products.append("Visio")
        if "Project" in selected_apps:
            products.append("Project")

        product_blocks = []
        for product in products:
            product_id = PRODUCT_IDS[product]
            if product == "Office" and exclude_apps_block:
                product_block = (
                    f'    <Product ID="{product_id}">\n'
                    f'      <Language ID="{language_id}" />\n'
                    f"{exclude_apps_block}\n"
                    f"    </Product>"
                )
            else:
                product_block = (
                    f'    <Product ID="{product_id}">\n'
                    f'      <Language ID="{language_id}" />\n'
                    f"    </Product>"
                )

            product_blocks.append(product_block)

        remove_msi_tag = "\n  <RemoveMSI />" if remove_msi else ""

        # Genera el bloque XML de configuración con los productos y
        # apps seleccionados
        config = f"""\
<Configuration>
  <Add OfficeClientEdition="{bits}" Channel="{self.versiones[selected_version]['channel']}">
{"\n".join(product_blocks)}
  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />{remove_msi_tag}
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>"""  # noqa: E501

        config_file_path = Path(self.office_install_dir) / "configuration.xml"
        try:
            config_file_path.write_text(config, encoding="utf-8")
            print(
                Fore.GREEN
                + (
                    "Archivo de configuración generado exitosamente en: "
                    f"{safe_log_path(config_file_path)}"
                )
            )
            return str(config_file_path)
        except Exception as e:
            logging.exception("Error al escribir configuration.xml")
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
        bits = self.combo_arch.get()
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
                "No se encontró la configuración para "
                "la versión seleccionada.",
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

    def show(self) -> bool:
        """
        Muestra la ventana de selección de Office con todos los componentes
        gráficos.
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
        self.combo_version.set("Office Standard 2013")
        self.combo_version.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.combo_version.bind("<<ComboboxSelected>>", self.update_apps)

        addons_frame = tb.Frame(frame)
        addons_frame.grid(row=2, column=0, sticky="w", pady=(0, 12))
        self.visio_var = tb.BooleanVar(value=False)
        self.project_var = tb.BooleanVar(value=False)
        self.visio_check = tb.Checkbutton(
            addons_frame, text="Instalar Visio", variable=self.visio_var
        )
        self.visio_check.pack(side="left", padx=(0, 16))
        self.project_check = tb.Checkbutton(
            addons_frame, text="Instalar Project", variable=self.project_var
        )
        self.project_check.pack(side="left")

        options_frame = tb.Frame(frame)
        options_frame.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        tb.Label(options_frame, text="Arquitectura:").grid(
            row=0, column=0, sticky="w"
        )
        self.combo_arch = tb.Combobox(
            options_frame, width=8, state="readonly", values=["32", "64"]
        )
        self.combo_arch.set("64")
        self.combo_arch.grid(row=0, column=1, sticky="w", padx=(8, 24))
        tb.Label(options_frame, text="Idioma:").grid(
            row=0, column=2, sticky="w"
        )
        self.combo_language = tb.Combobox(
            options_frame,
            width=20,
            state="readonly",
            values=sorted(self.languages.keys()),
        )
        self.combo_language.set("Spanish")
        self.combo_language.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self.remove_msi_var = tb.BooleanVar()
        tb.Checkbutton(
            frame,
            text="Eliminar versiones antiguas (RemoveMSI)",
            variable=self.remove_msi_var,
        ).grid(row=4, column=0, sticky="w", pady=(0, 12))

        tb.Label(frame, text="Aplicaciones a instalar:").grid(
            row=5, column=0, sticky="w"
        )
        self.frame_apps = tb.Frame(frame)
        self.frame_apps.grid(row=6, column=0, sticky="ew", pady=(0, 16))
        self.update_apps()

        tb.Button(
            frame,
            text="Instalar Office",
            command=self.install_office,
            width=24,
        ).grid(row=7, column=0, pady=(8, 0))

        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        center_window(self.root, w, h)
        self.root.minsize(
            self.root.winfo_reqwidth(), self.root.winfo_reqheight()
        )
        self.root.resizable(False, False)
        self.root.mainloop()
        return self.cancelled
