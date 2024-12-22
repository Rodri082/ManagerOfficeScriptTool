import os
import ctypes
import subprocess
import sys
import requests
import platform
import winreg
import tempfile
from typing import List
from urllib.parse import urlparse
from tqdm import tqdm
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import shutil
import threading
import datetime
from colorama import init, Fore

script_dir = (
    os.path.dirname(os.path.abspath(sys.argv[0]))
    if getattr(sys, 'frozen', False)
    else os.path.dirname(os.path.abspath(__file__))
)

files_dir = os.path.join(script_dir, 'Files')

ODT_2013 = "https://download.microsoft.com/download/6/2/3/6230F7A2-D8A9-478B-AC5C-57091B632FCF/officedeploymenttool_x86_5031-1000.exe"
ODT_2016_2019_2021 = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe"

version_scripts = {
    "2003": "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub03.vbs",
    "2007": "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub07.vbs",
    "2010": "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub10.vbs",
    "2013": "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O15msi.vbs",
    "2016": "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O16msi.vbs",
    "ClickToRun": "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/refs/heads/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrubc2r.vbs",
}

registry_cache = {}

class OfficeInstallation:
    def __init__(self, name: str, version: str, install_path: str, click_to_run: bool, product: str, bitness: str,
                 updates_enabled: bool, update_url: str, client_culture: str, version_to_report: str,
                 media_type: str, uninstall_string: str):
        self.name = name
        self.version = version
        self.install_path = install_path
        self.click_to_run = click_to_run
        self.product = product
        self.bitness = bitness
        self.updates_enabled = updates_enabled
        self.update_url = update_url
        self.client_culture = client_culture
        self.version_to_report = version_to_report
        self.media_type = media_type
        self.uninstall_string = uninstall_string

def get_registry_keys(key: str) -> List[str]:
    try:
        root_key = winreg.HKEY_LOCAL_MACHINE
        access_flag = winreg.KEY_READ | (winreg.KEY_WOW64_64KEY if platform.machine().endswith('64') else 0)

        with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
            subkeys = []
            index = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key_handle, index)
                    subkeys.append(subkey_name)
                    index += 1
                except OSError:
                    break
            return subkeys
    except FileNotFoundError:
        log_error(f"Error: No se encontró la clave {key}")
        return []
    except Exception as e:
        log_error(f"Error leyendo la clave del registro {key}: {e}")
        return []

def get_registry_value(key: str, value_name: str) -> str:
    cache_key = (key, value_name)
    if cache_key in registry_cache:
        return registry_cache[cache_key]
    
    try:
        root_key = winreg.HKEY_LOCAL_MACHINE
        access_flag = winreg.KEY_READ | (winreg.KEY_WOW64_64KEY if platform.machine().endswith('64') else 0)

        with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
            value, _ = winreg.QueryValueEx(key_handle, value_name)
            registry_cache[cache_key] = value
            return value
    except OSError as e:
        log_error(f"Error al acceder a la clave {key}, valor {value_name}: {e}")
        return ""

def get_office_installations(show_all_installed_products: bool = False) -> List[OfficeInstallation]:
    office_keys = [
        r"SOFTWARE\Microsoft\Office",
        r"SOFTWARE\Wow6432Node\Microsoft\Office"
    ]

    uninstall_keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    versions = ["ClickToRun", "15.0", "16.0"]
    installations = []
    found_names = set()

    for office_key in office_keys:
        for version in versions:
            if version == "ClickToRun":
                version_key = os.path.join(office_key, version, "Configuration")
            else:
                version_key = os.path.join(office_key, version, "ClickToRun", "Configuration")
            
            platform = get_registry_value(version_key, "Platform")
            if not platform:
                continue
            
            bitness = "64-Bits" if platform.strip() == "x64" else "32-Bits"
            client_culture = get_registry_value(version_key, "ClientCulture")
            updates_enabled = get_registry_value(version_key, "UpdatesEnabled") == "True"
            update_url = get_registry_value(version_key, "CDNBaseUrl")
            version_to_report = get_registry_value(version_key, "VersionToReport")
            product = get_registry_value(version_key, "ProductReleaseIds")
            media_type = get_registry_value(version_key, f"{product}.MediaType")

            for uninstall_key in uninstall_keys:
                for uninstall_subkey in get_registry_keys(uninstall_key):
                    uninstall_key_path = os.path.join(uninstall_key, uninstall_subkey)
                    display_name = get_registry_value(uninstall_key_path, "DisplayName")
                    
                    if display_name and display_name not in found_names:
                        found_names.add(display_name)
                        
                        if "Microsoft Office" not in display_name and "Microsoft 365" not in display_name:
                            continue
                        
                        display_version = get_registry_value(uninstall_key_path, "DisplayVersion")
                        install_location = get_registry_value(uninstall_key_path, "InstallLocation")
                        uninstall_string = get_registry_value(uninstall_key_path, "UninstallString")
                        click_to_run = True if "ClickToRun" in uninstall_string else False

                        installations.append(
                            OfficeInstallation(
                                name=display_name,
                                version=display_version,
                                install_path=install_location,
                                click_to_run=click_to_run,
                                product=product,
                                bitness=bitness,
                                updates_enabled=updates_enabled,
                                update_url=update_url,
                                client_culture=client_culture,
                                version_to_report=version_to_report,
                                media_type=media_type,
                                uninstall_string=uninstall_string,
                            )
                        )
    
    if not show_all_installed_products:
        installations = [inst for inst in installations if "Microsoft Office" in inst.name or "Microsoft 365" in inst.name]
    
    return installations

def display_office_installations(installations: List[OfficeInstallation]):
    print("-------------------------------------------------------------------------------------------------------------")
    print(Fore.RED + "Se encontraron las siguientes instalaciones de Microsoft Office.")
    print("-------------------------------------------------------------------------------------------------------------")
    for install in installations:
        office_info = {
            "DisplayName": install.name,
            "Version": install.version,
            "Language": install.client_culture,
            "Architecture": install.bitness,
            "IsClickToRun": install.click_to_run,
            "InstallPath": install.install_path,
            "ProductID": install.product,
            "IsUpdatesEnabled": install.updates_enabled,
            "UpdateChannelUrl": install.update_url,
            "MediaType": install.media_type
        }

        for key, value in office_info.items():
            print(f"{key}: {'-' * (35 - len(key))} " + str(value))
    print("-------------------------------------------------------------------------------------------------------------")

def detect_office_version(install_folder):
    installations = get_office_installations(show_all_installed_products=False)

    if not installations:
        print(Fore.GREEN + "No se encontraron instalaciones de Microsoft Office.")
        ask_proceed_installation(install_folder)
        return []

    display_office_installations(installations)
    ask_uninstall_and_proceed(install_folder, installations)

    return installations

def get_computer_architecture() -> str:
    return platform.architecture()[0]

def download_script(url: str, destination: str) -> bool:
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        with open(destination, 'wb') as file:
            file.write(response.content)
        return True
    except Exception as e:
        log_error(f"Error descargando el script desde {url}: {e}")
        return False

def uninstall_office_version(installations):
    script_url = None
    if installations.click_to_run:
        script_url = version_scripts["ClickToRun"]
    else:
        if "2003" in installations.version:
            script_url = version_scripts["2003"]
        elif "2007" in installations.version:
            script_url = version_scripts["2007"]
        elif "2010" in installations.version:
            script_url = version_scripts["2010"]
        elif "2013" in installations.version:
            script_url = version_scripts["2013"]
        elif "2016" in installations.version:
            script_url = version_scripts["2016"]

    if not script_url:
        log_error(f"No se pudo determinar el script adecuado para la versión: {installations.version}")
        return

    script_name = script_url.split("/")[-1]
    script_path = os.path.join(tempfile.gettempdir(), script_name)

    if not download_script(script_url, script_path):
        print(Fore.RED + f"No se pudo descargar el script para la versión {installations.version}.")
        return

    try:
        command = ["cscript", "//nologo", "//X", script_path]
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        for line in process.stdout:
            print(line.strip())
        process.wait()

        if process.returncode == 0:
            print(Fore.GREEN + f"Desinstalación completada para {installations.name} ({installations.version}).")
        else:
            error_message = process.stderr.read()
            log_error(f"Error al ejecutar el script de desinstalación: {error_message}")

    except Exception as e:
        log_error(f"Error al ejecutar el script de desinstalación: {e}")
    finally:
        if os.path.exists(script_path):
            os.remove(script_path)

def create_office_selection_window(install_folder):
    local_exe_folder = os.path.join(files_dir, "InstallOfficeFiles")
    os.makedirs(local_exe_folder, exist_ok=True)
    install_folder = local_exe_folder

    def on_closing():
        if messagebox.askyesno("Eliminar archivos temporales", "¿Deseas eliminar los archivos temporales creados por el script?"):
            try:
                remove_temp_folder(install_folder)
                messagebox.showinfo("Archivos eliminados", "Los archivos temporales han sido eliminados.")
                sys.exit(0)
            except Exception as e:
                messagebox.showwarning("Error", f"No se pudieron eliminar los archivos temporales: {e}")
                sys.exit(1)
        root.destroy()

    all_apps = {
        "Office Standard 2013": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "Groove"],
        "Office Professional Plus 2013": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "Access", "InfoPath", "Lync", "Groove"],
        "Office Standard 2016": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive", "Groove"],
        "Office Professional Plus 2016": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive", "Access", "Lync", "Groove"],
        "Office Standard 2019": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive", "Groove"],
        "Office Professional Plus 2019": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive", "Access", "Lync", "Groove"],
        "Office LTSC Standard 2021": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive"],
        "Office LTSC Professional Plus 2021": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "OneDrive", "Access", "Teams", "Lync"],
        "Office LTSC Standard 2024": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "OneDrive"],
        "Office LTSC Professional Plus 2024": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "OneDrive", "Access", "Lync"],
        "Microsoft 365 HomePremium - Current": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Publisher", "Access", "OneDrive"],
        "Microsoft 365 Apps for business - Current": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "Publisher", "Access", "Lync", "OneDrive", "Groove"],
        "Microsoft 365 Apps for business - MonthlyEnterprise": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "Publisher", "Access", "Lync", "OneDrive", "Groove"],
        "Microsoft 365 Apps for business - SemiAnnual": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "Publisher", "Access", "Lync", "OneDrive", "Groove"],
        "Microsoft 365 Apps for enterprise - Current": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "Publisher", "Access", "Lync", "OneDrive", "Groove"],
        "Microsoft 365 Apps for enterprise - MonthlyEnterprise": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "Publisher", "Access", "Lync", "OneDrive", "Groove"],
        "Microsoft 365 Apps for enterprise - SemiAnnual": ["Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "Publisher", "Access", "Lync", "OneDrive", "Groove"],
    }

    def generate_configuration(selected_version, product_id, channel, bits, language, remove_msi, selected_apps):
        remove_msi_tag = "\n    <RemoveMSI />" if remove_msi else ""
        available_apps = all_apps.get(selected_version, [])
        apps_to_exclude = [app for app in available_apps if app not in selected_apps]
        
        exclude_apps = ""
        if apps_to_exclude:
            exclude_apps = "\n".join([f'            <ExcludeApp ID="{app}" />' for app in apps_to_exclude])
        
        config = f"""<Configuration>
    <Add OfficeClientEdition="{bits}" Channel="{channel}">
        <Product ID="{product_id}">
            <Language ID="{language}" />\n{exclude_apps}
        </Product>
    </Add>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="AUTOACTIVATE" Value="1" />
    <Updates Enabled="TRUE" />{remove_msi_tag}
    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
    """
        file_name = os.path.join(install_folder, "configuration.xml")
        with open(file_name, 'w') as file:
            file.write(config)
            root.destroy()
        return file_name

    def update_apps(event=None):
        selected_version = combo_version.get()
        available_apps = all_apps.get(selected_version, [])
        
        for widget in app_checkbuttons:
            widget.grid_forget()
        
        app_vars.clear()
        app_checkbuttons.clear()
        for i, app in enumerate(available_apps):
            var = tk.BooleanVar()
            app_vars[app] = var
            cb = ttk.Checkbutton(frame_apps, text=app, variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, pady=3)
            app_checkbuttons.append(cb)

    def install_office():
        selected_version = combo_version.get()
        bits = combo_bits.get()
        language = combo_language.get()
        remove_msi = remove_msi_var.get()

        selected_apps = [app for app, var in app_vars.items() if var.get()]

        if not selected_apps:
            messagebox.showwarning("Advertencia", "Por favor selecciona al menos una aplicación a instalar.")
            return

        if not selected_version or not bits or not language:
            messagebox.showwarning("Advertencia", "Por favor selecciona una versión, arquitectura y lenguaje.")
            return

        versiones = {
            "Office Standard 2013": ("StandardRetail", "Current"),
            "Office Professional Plus 2013": ("ProplusRetail", "Current"),
            "Office Standard 2016":("StandardRetail", "Current"),
            "Office Professional Plus 2016": ("ProplusRetail", "Current"),
            "Office Standard 2019": ("Standard2019Volume", "PerpetualVL2019"),
            "Office Professional Plus 2019": ("ProPlus2019Volume", "PerpetualVL2019"),
            "Office LTSC Standard 2021": ("Standard2021Volume", "PerpetualVL2021"),
            "Office LTSC Professional Plus 2021": ("ProPlus2021Volume", "PerpetualVL2021"),
            "Office LTSC Standard 2024": ("Standard2024Volume", "PerpetualVL2024"),
            "Office LTSC Professional Plus 2024": ("ProPlus2024Volume", "PerpetualVL2024"),
            "Microsoft 365 HomePremium - Current":("O365HomePremRetail", "Current"),
            "Microsoft 365 Apps for business - Current": ("O365BusinessRetail", "Current"),
            "Microsoft 365 Apps for business - MonthlyEnterprise": ("O365BusinessRetail", "MonthlyEnterprise"),
            "Microsoft 365 Apps for business - SemiAnnual": ("O365BusinessRetail", "SemiAnnual"),
            "Microsoft 365 Apps for enterprise - Current": ("O365ProPlusRetail", "Current"),
            "Microsoft 365 Apps for enterprise - MonthlyEnterprise": ("O365ProPlusRetail", "MonthlyEnterprise"),
            "Microsoft 365 Apps for enterprise - SemiAnnual": ("O365ProPlusRetail", "SemiAnnual"),
        }

        product_id, channel = versiones.get(selected_version, (None, None))

        if not product_id or not channel:
            messagebox.showerror("Error", "No se pudo encontrar la configuración para la versión seleccionada.")
            return

        generate_configuration(selected_version, product_id, channel, bits, language, remove_msi, selected_apps)
        download_and_extract_ODT(selected_version, install_folder)
    root = tk.Tk()
    root.title("Selector de Versiones de Microsoft Office")
    root.geometry("435x600")
    root.config(bg="#f4f4f4")

    frame = ttk.Frame(root, padding="20")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

    ttk.Label(frame, text="Selecciona la versión de Office:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
    combo_version = ttk.Combobox(frame, width=50, font=("Arial", 10))
    combo_version["values"] = [
        "Office Standard 2013", "Office Professional Plus 2013", "Office Standard 2016", "Office Professional Plus 2016",
        "Office Standard 2019", "Office Professional Plus 2019", "Office LTSC Standard 2021", "Office LTSC Professional Plus 2021",
        "Office LTSC Standard 2024", "Office LTSC Professional Plus 2024", "Microsoft 365 HomePremium - Current",
        "Microsoft 365 Apps for business - Current", "Microsoft 365 Apps for business - MonthlyEnterprise", 
        "Microsoft 365 Apps for business - SemiAnnual", "Microsoft 365 Apps for enterprise - Current", 
        "Microsoft 365 Apps for enterprise - MonthlyEnterprise", "Microsoft 365 Apps for enterprise - SemiAnnual",
    ]
    combo_version.grid(row=1, column=0, sticky=tk.W, pady=5)
    combo_version.bind("<<ComboboxSelected>>", update_apps)
    combo_version.set("Office Standard 2013")

    ttk.Label(frame, text="Selecciona la arquitectura:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
    combo_bits = ttk.Combobox(frame, width=10, font=("Arial", 10))
    combo_bits["values"] = ["32", "64"]
    combo_bits.grid(row=3, column=0, sticky=tk.W, pady=5)
    combo_bits.set("64")

    ttk.Label(frame, text="Selecciona el idioma:", font=("Arial", 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
    combo_language = ttk.Combobox(frame, width=20, font=("Arial", 10))
    combo_language["values"] = ["es-es", "en-us"]
    combo_language.grid(row=5, column=0, sticky=tk.W, pady=5)
    combo_language.set("es-es")

    remove_msi_var = tk.BooleanVar()
    remove_msi_checkbox = ttk.Checkbutton(frame, text="Agregar RemoveMSI", variable=remove_msi_var)
    remove_msi_checkbox.grid(row=6, column=0, sticky=tk.W, pady=5)

    frame_apps = ttk.Frame(frame)
    frame_apps.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)

    app_vars = {}
    app_checkbuttons = []

    frame.grid_rowconfigure(16, weight=1)
    update_apps()

    btn_instalar = ttk.Button(frame, text="Generar Configuración", command=install_office)
    btn_instalar.grid(row=17, column=0, pady=10)

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.attributes('-topmost', True)

    root.mainloop()

def download_and_extract_ODT(selected_version, install_folder):
    url = ODT_2013 if "2013" in selected_version else ODT_2016_2019_2021

    if not url:
        log_error("La URL proporcionada es inválida.")
        return

    try:
        temp_download_folder = tempfile.mkdtemp()
        exe_file_name = os.path.basename(urlparse(url).path)
        exe_file_path = os.path.join(temp_download_folder, exe_file_name)

        print(f"Descargando ODT desde: {url}")
        response = requests.get(url, stream=True)
        response.raise_for_status()

        total_size = int(response.headers.get('Content-Length', 0))
        with open(exe_file_path, "wb") as file, tqdm(
            total=total_size, unit="B", unit_scale=True, desc="Descargando"
        ) as bar:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
                bar.update(len(chunk))

        print(f"Archivo descargado en: {exe_file_path}")
    except Exception as e:
        log_error(f"Error al descargar ODT: {e}")
        return

    try:
        print(f"Extrayendo archivos en: {temp_download_folder}")
        subprocess.run(
            [exe_file_path, "/quiet", f"/extract:{temp_download_folder}"],
            check=True
        )
        print("Extracción completada.")
    except subprocess.CalledProcessError as e:
        log_error(f"Error al extraer ODT: {e}")
        return
    except Exception as e:
        log_error(f"Error desconocido durante la extracción: {e}")
        return

    try:
        setup_file = os.path.join(temp_download_folder, "setup.exe")
        if os.path.exists(setup_file):
            shutil.copy(setup_file, install_folder)
            print(f"'setup.exe' copiado con éxito a: {install_folder}")
        else:
            log_error(f"'setup.exe' no encontrado en {temp_download_folder}.")
    except Exception as e:
        log_error(f"Error al copiar 'setup.exe': {e}")

    finally:
        shutil.rmtree(temp_download_folder, ignore_errors=True)

def run_setup_commands(install_folder):
    command = ["cmd", "/c", "setup.exe", "/configure", "configuration.xml"]
    working_directory = install_folder

    try:
        print(Fore.RED + f"Por favor, no cierre esta ventana. Se Instalará la Versión de Microsoft Office Seleccionada...")
        subprocess.run(command, cwd=working_directory, stdout=sys.stdout, stderr=sys.stderr, check=True)
        print("Instalación completada")
    except FileNotFoundError as e:
        log_error(f"El archivo setup.exe no se encuentra en la ubicación esperada: {e}")
    except Exception as e:
        log_error(f"Error al ejecutar comando: {e}")
        sys.exit(1)

def log_error(error_message):
    try:
        logs_folder = os.path.join(files_dir, "logs")
        os.makedirs(logs_folder, exist_ok=True)

        log_path = os.path.join(logs_folder, "log_error.txt")

        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"{current_time} - {error_message}\n")
    except Exception as e:
        print(f"Error al registrar el log_error: {e}")

def check_log_retention(log_path):
    try:
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as log_file:
                lines = log_file.readlines()

            if lines:
                last_line = lines[-1]
                last_log_time_str = last_line.split(" - ")[0]
                last_log_time = datetime.datetime.strptime(last_log_time_str, "%Y-%m-%d %H:%M:%S")

                current_time = datetime.datetime.now()
                retention_period = datetime.timedelta(days=30)

                if current_time - last_log_time >= retention_period:
                    os.remove(log_path)

    except Exception as e:
        print(f"Error al comprobar o eliminar el archivo de log: {e}")

def check_install_folder():
    install_folder = os.path.join(files_dir, "InstallOfficeFiles")

    try:
        if os.path.exists(install_folder):
            shutil.rmtree(install_folder)

    except OSError as e:
        log_error(f"Error al crear el directorio {install_folder}: {e}")
        raise

    return install_folder

def remove_temp_folder(install_folder):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    if not messagebox.askyesno("Eliminar archivos temporales", "¿Deseas eliminar los archivos temporales creados por el script?"):
        messagebox.showinfo("Cancelado", "No se eliminaron los archivos temporales.")
        return

    try:
        for folder in [install_folder]:
            if os.path.isdir(folder):
                shutil.rmtree(folder)
        
        messagebox.showinfo("Archivos eliminados", "Los archivos temporales han sido eliminados correctamente.")
        sys.exit(0)
    except Exception as e:
        messagebox.showwarning("Error", f"No se pudieron eliminar los archivos temporales: {e}")
        sys.exit(1)

    sys.exit(0)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception as e:
        log_error(f"Error al verificar privilegios de administrador: {e}")
        return False

def manage_installation():
    check_log_retention(os.path.join(files_dir, "logs", "log_error.txt"))

    install_folder = check_install_folder()

    detect_office = input("¿Desea detectar las versiones de Office instaladas? (S/N): ").strip().lower()
    if detect_office == 's':
        office_names = detect_office_version(install_folder)
        uninstall_decision = ask_uninstall_and_proceed(install_folder, office_names)
    else:
        print(Fore.RED + "Advertencia: Tener más de una versión de Office instalada puede ocasionar problemas.")
        uninstall_decision = ask_proceed_installation(install_folder)

    if uninstall_decision:
        print("Instalación completada.")
    else:
        print("Proceso cancelado.")
        remove_temp_folder(install_folder)

def ask_uninstall_and_proceed(install_folder,installations):
    uninstall_office = input("¿Desea desinstalar todas las versiones de Office encontradas? (S/N): ").strip().lower()
    if uninstall_office == 's':
        for installation in installations:
            uninstall_office_version(installation)
            ask_proceed_installation(install_folder)
    else:
        print(Fore.RED + "Advertencia: Tener más de una versión de Office instalada puede ocasionar problemas.")
        ask_proceed_installation(install_folder)

def ask_proceed_installation(install_folder):
    proceed_install = input("¿Desea proceder con la instalación de Office? (S/N): ").strip().lower()
    if proceed_install == 's':
        print("Iniciando instalación de Office...")
        execute_installation_process(install_folder)
    else:
        print("Proceso cancelado.")
        remove_temp_folder(install_folder)

def execute_installation_process(install_folder):
    create_office_selection_window(install_folder)

    install_thread = threading.Thread(target=run_setup_commands, args=(install_folder,))
    install_thread.start()
    install_thread.join()

    remove_temp_folder(install_folder)

def main():
    if not is_admin():
        print("Este script necesita ejecutarse con privilegios de administrador para continuar.")
        input("Presiona Enter para salir y vuelve a ejecutar como administrador...")
        sys.exit(0)
    init(autoreset=True)
    manage_installation()

if __name__ == "__main__":
    main()