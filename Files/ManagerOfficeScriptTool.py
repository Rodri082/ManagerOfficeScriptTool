import os
import subprocess
import sys
import requests
import platform
import winreg
import zipfile
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

global_folder_destination = None

script_dir = (
    os.path.dirname(os.path.abspath(sys.argv[0]))
    if getattr(sys, 'frozen', False)
    else os.path.dirname(os.path.abspath(__file__))
)

files_dir = os.path.join(script_dir, 'Files')

ODT_2013 = "https://download.microsoft.com/download/6/2/3/6230F7A2-D8A9-478B-AC5C-57091B632FCF/officedeploymenttool_x86_5031-1000.exe"
ODT_2016_2019_2021 = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"
SARA_URL = "https://aka.ms/SaRA_CommandLineVersionFiles"

class OfficeInstallation:
    def __init__(self, name: str, version: str, install_path: str, click_to_run: bool, bitness: str,
                 computer_name: str, updates_enabled: bool, update_url: str, client_culture: str, version_to_report: str,
                 proplus_retail_media_type: str, auto_upgrade: str, uninstall_string: str):
        self.name = name
        self.version = version
        self.install_path = install_path
        self.click_to_run = click_to_run
        self.bitness = bitness
        self.computer_name = computer_name
        self.updates_enabled = updates_enabled
        self.update_url = update_url
        self.client_culture = client_culture
        self.version_to_report = version_to_report
        self.proplus_retail_media_type = proplus_retail_media_type
        self.auto_upgrade = auto_upgrade
        self.uninstall_string = uninstall_string

def get_registry_keys(key: str) -> List[str]:
    try:
        root_key = winreg.HKEY_LOCAL_MACHINE
        if platform.machine().endswith('64'):
            access_flag = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
        else:
            access_flag = winreg.KEY_READ
        
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

registry_cache = {}

def get_registry_value(key: str, value_name: str) -> str:
    if (key, value_name) in registry_cache:
        return registry_cache[(key, value_name)]
    try:
        root_key = winreg.HKEY_LOCAL_MACHINE
        if platform.machine().endswith('64'):
            access_flag = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
        else:
            access_flag = winreg.KEY_READ
        
        with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
            value, _ = winreg.QueryValueEx(key_handle, value_name)
            registry_cache[(key, value_name)] = value
            return value
    except OSError as e:
        return ""

def get_office_installations(computer_name: str, show_all_installed_products: bool = False) -> List[OfficeInstallation]:
    office_keys = [
        r"SOFTWARE\Microsoft\Office",
        r"SOFTWARE\Wow6432Node\Microsoft\Office"
    ]

    uninstall_keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    installations = []
    found_names = set()
    
    for office_key in office_keys:
        for version in get_registry_keys(office_key):
            if version in ['Common', 'Delivery', 'DmsClient', 'Outlook']:
                continue
            
            office_version_key = os.path.join(office_key, version)
            install_path = get_registry_value(office_version_key, "InstallationPath")
            click_to_run_path = os.path.join(office_version_key, "ClickToRun\\Configuration")
            client_culture = get_registry_value(click_to_run_path, "ClientCulture")
            updates_enabled = get_registry_value(click_to_run_path, "UpdatesEnabled") == "True"
            update_url = get_registry_value(click_to_run_path, "CDNBaseUrl")
            platform = get_registry_value(click_to_run_path, "Platform")

            bitness = "64-Bit" if "x64" in install_path or platform == "x64" else "32-Bit"

            version_to_report = get_registry_value(click_to_run_path, "VersionToReport")
            proplus_retail_media_type = get_registry_value(click_to_run_path, "ProPlusRetail.MediaType")
            auto_upgrade = get_registry_value(click_to_run_path, "AutoUpgrade")

            for uninstall_key in uninstall_keys:
                for uninstall_subkey in get_registry_keys(uninstall_key):
                    uninstall_key_path = os.path.join(uninstall_key, uninstall_subkey)
                    display_name = get_registry_value(uninstall_key_path, "DisplayName")
                    
                    if display_name and display_name not in found_names:
                        found_names.add(display_name)
                    
                        if "Microsoft Office" not in display_name:
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
                                bitness=bitness,
                                computer_name=computer_name,
                                updates_enabled=updates_enabled,
                                update_url=update_url,
                                client_culture=client_culture,
                                version_to_report=version_to_report,
                                proplus_retail_media_type=proplus_retail_media_type,
                                auto_upgrade=auto_upgrade,
                                uninstall_string=uninstall_string,
                            )
                        )

    if not show_all_installed_products:
        installations = [inst for inst in installations if "Microsoft Office" in inst.name]
    
    return installations

def get_computer_architecture() -> str:
    return platform.architecture()[0]

def Detect_OfficeVersion(install_folder, uninstall_folder):
    computer_name = platform.node()
    installations = get_office_installations(computer_name, show_all_installed_products=False)

    if not installations:
        print(Fore.GREEN + "No se encontraron instalaciones de Microsoft Office.")
        ask_proceed_installation(install_folder, uninstall_folder)

    office_names = []

    print("-------------------------------------------------------------------------------------------------------------")
    print(Fore.RED + "Se encontraron las siguientes instalaciones de Microsoft Office.")
    print("-------------------------------------------------------------------------------------------------------------")
    for install in installations:
        office_names.append(install.name)
        office_info = {
            "DisplayName": install.name,
            "Version": install.version,
            "InstallPath": install.install_path,
            "ClickToRun": install.click_to_run,
            "Bitness": install.bitness,
            "ComputerName": install.computer_name,
            "UpdatesEnabled": install.updates_enabled,
            "UpdateUrl": install.update_url,
            "ClientCulture": install.client_culture,
            "ProPlusRetail.MediaType": install.proplus_retail_media_type,
            "AutoUpgrade": install.auto_upgrade
        }

        for key, value in office_info.items():
            print(f"{key}: {'-' * (35 - len(key))} " + str(value))
    print("-------------------------------------------------------------------------------------------------------------")
    ask_uninstall_and_proceed(install_folder, uninstall_folder, office_names)

    return office_names

def Uninstall_OfficeVersion(office_names):
    global global_folder_destination
    download_file()

    office_versions = {
        "2007": "2007",
        "2013": "2013",
        "2016": "2016",
        "2019": "2019",
        "2021": "2021",
        "365": "M365"
    }

    found_versions = []

    for name in office_names:
        for version_key, version_value in office_versions.items():
            if version_key in name:
                if version_value not in found_versions:
                    found_versions.append(version_value)
    
    if len(found_versions) > 1:
        office_version = "All"
    elif found_versions:
        office_version = found_versions[0]
    else:
        office_version = None

    if office_version:
        try:
            sara_cmd_path = global_folder_destination
            command0 = f"SaRACmd.exe -DisplayEULA"
            print(Fore.RED + f"Ejecutando {command0} para mostrar los Términos y condiciones de Microsoft para SaRACmd.exe")
            subprocess.run(command0, cwd=sara_cmd_path, stdout=sys.stdout, stderr=sys.stderr, check=True, shell=True)
            print()

            uninstall_office = input(f"""¿Desea aceptar los Términos y condiciones de Microsoft para utilizar SaRA_CommandLineVersion \ny continuar con la desinstalación de Microsoft Office {office_version}? (S/N): """).strip().lower()

            if uninstall_office == 's':
                command = f"SaRACmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion {office_version}"
                print()
                print(f"Ejecutando {command} para desinstalar Microsoft Office {office_version}")
                print()
                print(Fore.RED + f"Por favor, no cierre esta ventana. Se desinstalará Microsoft Office {office_version}")
                subprocess.run(command, cwd=sara_cmd_path, stdout=sys.stdout, stderr=sys.stderr, check=True, shell=True)
                print("Desinstalación completada")

                restart_windows = input(
                    f"""Se recomienda reiniciar el equipo para finalizar las tareas de limpieza restantes.
                    ¿Quiere reinciar ahora? (S/N): """).strip().lower()

                if restart_windows == 's':
                    os.system("shutdown /r /t 5")
                    sys.exit(0)
                else:
                    return

            else:
                input(f"No se procederá con la desinstalación de Microsoft Office {office_version}. Presione una tecla para continuar...")
                print()
                print("Desinsatalación abortada")

        except FileNotFoundError as e:
            log_error(f"Error: Archivo no encontrado. {e}")
        except subprocess.CalledProcessError as e:
            log_error(f"Error: El comando falló. {e}")
        except Exception as e:
            log_error(f"Error inesperado: {e}")

def get_final_url():
    url = SARA_URL
    final_url = None
    try:
        response_original = requests.head(url, allow_redirects=False)
        response_original.raise_for_status()
        
        if 'Location' in response_original.headers:
            final_url = response_original.headers['Location']
        else:
            log_error("No se encontró el encabezado 'Location'.")
    
    except requests.exceptions.RequestException as e:
        log_error(f"Error al obtener los encabezados: {e}")
    
    return final_url

def download_file():
    final_url = get_final_url()
    if final_url:
        try:
            uninstall_folder = os.path.join(files_dir, "Uninstall")
            os.makedirs(uninstall_folder, exist_ok=True)
            
            file_name = os.path.basename(urlparse(final_url).path)
            file_destination = os.path.join(uninstall_folder, file_name)
            
            print(f"Descargando {file_name} desde {final_url}")
            response = requests.get(final_url, stream=True)
            response.raise_for_status()
            
            total_size = int(response.headers.get('Content-Length', 0))
            if total_size == 0:
                log_error("No se pudo obtener el tamaño del archivo.")
                return

            with open(file_destination, 'wb') as f, tqdm(
                total=total_size, unit='B', unit_scale=True, desc="Progreso"
            ) as bar:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
                        bar.update(len(chunk))

            print(f"\nArchivo descargado con éxito en: {file_destination}")
            
            if file_destination.lower().endswith('.zip'):
                extract_zip_file(file_destination, uninstall_folder)
        
        except requests.exceptions.RequestException as e:
            log_error(f"Error al descargar el archivo: {e}")
    else:
        log_error("No se pudo obtener la URL final.")

def extract_zip_file(zip_file, destination):
    global global_folder_destination
    try:
        folder_name = os.path.splitext(os.path.basename(zip_file))[0]
        folder_destination = os.path.join(destination, folder_name)
        os.makedirs(folder_destination, exist_ok=True)

        global_folder_destination = folder_destination
        
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(folder_destination)
    except zipfile.BadZipFile:
        log_error(f"El archivo {zip_file} no es un archivo .zip válido.")
    except Exception as e:
        log_error(f"Error al extraer el archivo .zip: {e}")

def create_tooltip(widget, text):
    tooltip = tk.Toplevel(widget)
    tooltip.withdraw()
    tooltip.overrideredirect(True)
    tooltip_label = ttk.Label(tooltip, text=text, background="yellow", relief="solid", borderwidth=1, wraplength=200)
    tooltip_label.pack()

    def show_tooltip(event):
        tooltip.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        tooltip.deiconify()

    def hide_tooltip(event):
        tooltip.withdraw()

    widget.bind("<Enter>", show_tooltip)
    widget.bind("<Leave>", hide_tooltip)

def create_office_selection_window(install_folder, uninstall_folder):
    def on_closing():
        if messagebox.askyesno("Eliminar archivos temporales", "¿Deseas eliminar los archivos temporales creados por el script?"):
            try:
                remove_temp_folder(install_folder, uninstall_folder)
                messagebox.showinfo("Archivos eliminados", "Los archivos temporales han sido eliminados.")
                sys.exit(0)
            except Exception as e:
                messagebox.showwarning("Error", f"No se pudieron eliminar los archivos temporales: {e}")
                sys.exit(1)
        root.destroy()

    def update_apps(event=None):
        selected_version = version_var.get()
        app_options = get_app_options(selected_version)
        for app in app_checkboxes.values():
            app.destroy()
        app_checkboxes.clear()
        for app_name in app_options:
            app_vars[app_name] = tk.BooleanVar()
            app_vars[app_name].set(False)
            app_checkboxes[app_name] = tk.Checkbutton(apps_frame, text=app_name, variable=app_vars[app_name])
            app_checkboxes[app_name].pack(anchor=tk.W)

    def get_app_options(selected_version):
        if selected_version == "Microsoft Office 2021":
            return ["Word", "Excel", "Access", "Publisher", "Groove", "Lync", "OneNote", "Outlook", "PowerPoint", "OneDrive", "Teams"]
        else:
            return ["Word", "Excel", "Access", "Publisher", "Groove", "Lync", "OneNote", "Outlook", "PowerPoint", "OneDrive"]

    root = tk.Tk()
    root.title("Personalizar Instalación de Office")
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.attributes('-topmost', True)
    root.update()
    root.attributes('-topmost', False)

    version_frame = tk.Frame(root)
    version_frame.grid(row=0, column=0, padx=10, pady=10)
    version_label = tk.Label(version_frame, text="Versión de Office:")
    version_label.pack()
    version_options = ["Microsoft Office 2013", "Microsoft Office 2016", "Microsoft Office 2019", "Microsoft Office 2021"]
    version_var = tk.StringVar(root, version_options[0])
    version_combo = tk.OptionMenu(version_frame, version_var, *version_options, command=update_apps)
    version_combo.pack()
    create_tooltip(version_combo, "Selecciona la versión de Office que deseas instalar.")

    architecture_frame = tk.Frame(root)
    architecture_frame.grid(row=1, column=0, padx=10, pady=10)
    architecture_label = tk.Label(architecture_frame, text="Arquitectura:")
    architecture_label.pack()
    architecture_var = tk.StringVar()
    architecture_var.set("32 Bits")
    bits32_radio = tk.Radiobutton(architecture_frame, text="32 Bits", variable=architecture_var, value="32 Bits")
    bits32_radio.pack(side=tk.LEFT)
    bits64_radio = tk.Radiobutton(architecture_frame, text="64 Bits", variable=architecture_var, value="64 Bits")
    bits64_radio.pack(side=tk.LEFT)
    create_tooltip(bits32_radio, "Elige si la arquitectura del sistema es 32 bits.")
    create_tooltip(bits64_radio, "Elige si la arquitectura del sistema es 64 bits.")

    edition_frame = tk.Frame(root)
    edition_frame.grid(row=2, column=0, padx=10, pady=10)
    edition_label = tk.Label(edition_frame, text="Edición de Office:")
    edition_label.pack()
    edition_var = tk.StringVar()
    edition_var.set("ProPlus")
    proplus_radio = tk.Radiobutton(edition_frame, text="ProPlus", variable=edition_var, value="ProPlus")
    proplus_radio.pack(side=tk.LEFT)
    standard_radio = tk.Radiobutton(edition_frame, text="Standard", variable=edition_var, value="Standard")
    standard_radio.pack(side=tk.LEFT)
    create_tooltip(proplus_radio, "Selecciona para obtener la version Professional.")
    create_tooltip(standard_radio, "Selecciona para obtener la version Estandar.")

    apps_frame = tk.Frame(root)
    apps_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=10)
    apps_label = tk.Label(apps_frame, text="Aplicaciones a instalar:")
    apps_label.pack()
    app_vars = {}
    app_checkboxes = {}

    update_apps()

    def accept():
        selected_apps = [app for app in app_vars if app_vars[app].get()]
        selected_architecture = architecture_var.get()
        selected_edition = edition_var.get()
        selected_version = version_var.get()
        selected_apps_str = ",".join(selected_apps)

        selected_options = {
            "Version": selected_version,
            "Edition": selected_edition,
            "Architecture": selected_architecture,
            "Applications": selected_apps_str
        }

        try:
            local_exe_folder = os.path.join(files_dir, "InstallOfficeFiles")
            os.makedirs(local_exe_folder, exist_ok=True)
            file_path = os.path.join(install_folder, "selected_options.txt")
            with open(file_path, "w") as file:
                for key, value in selected_options.items():
                    file.write(f"{key}: {value}\n")
            messagebox.showinfo("Opciones guardadas", "Tus selecciones han sido guardadas con éxito.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron guardar las opciones: {e}")
            return

        root.destroy()

    accept_button = tk.Button(root, text="Aceptar", command=accept)
    accept_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
    create_tooltip(accept_button, "Haz clic para continuar con la instalación.")

    center_window(root)
    root.mainloop()

def center_window(window):
    window_width = window.winfo_reqwidth()
    window_height = window.winfo_reqheight()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = int((screen_width - window_width) / 2)
    y_coordinate = int((screen_height - window_height) / 2)
    window.geometry(f"+{x_coordinate}+{y_coordinate}")

def read_selected_options(file_path):
    selected_options = {}
    try:
        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                for line in file:
                    key, value = line.strip().split(": ")
                    selected_options[key] = value
    except FileNotFoundError:
        log_error(f"El archivo {file_path} no fue encontrado.")
    except IOError as e:
        log_error(f"Error al leer el archivo {file_path}: {e}")
    except ValueError as e:
        log_error(f"Formato inválido en el archivo {file_path}: {e}")
    
    return selected_options


def extract_year(version_string):
    words = version_string.split()
    for word in words:
        if word.isdigit():
            return word
    
    return None

def format_edition(version, edition):
    if version == "2013" or version == "2016":
        return f"{edition}Retail"
    elif version == "2019":
        return f"{edition}2019Retail"
    elif version == "2021":
        return f"{edition}2021Retail"
    else:
        return "Unknown"
    
def format_architecture(architecture):
    if architecture == "32 Bits":
        return "32"
    elif architecture == "64 Bits":
        return "64"
    else:
        return "Unknown"
    
def format_applications(applications):
    formatted_apps = applications.replace(",", "\n")
    return formatted_apps

def download_and_extract_ODT(version, install_folder):
    local_exe_folder = os.path.join(files_dir, "InstallOfficeFiles")
    os.makedirs(local_exe_folder, exist_ok=True)

    url = ODT_2013 if version == "2013" else ODT_2016_2019_2021

    if not url:
        log_error("La URL proporcionada es inválida.")
        return

    try:
        exe_file_name = os.path.basename(urlparse(url).path)
        exe_file_path = os.path.join(local_exe_folder, exe_file_name)

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

    temp_extract_folder = tempfile.mkdtemp()
    try:
        print(f"Extrayendo archivos en: {temp_extract_folder}")
        subprocess.run(
            [exe_file_path, "/quiet", f"/extract:{temp_extract_folder}"],
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
        setup_file = os.path.join(temp_extract_folder, "setup.exe")
        if os.path.exists(setup_file):
            shutil.copy(setup_file, install_folder)
            print(f"'setup.exe' copiado con éxito a: {install_folder}")
        else:
            log_error(f"'setup.exe' no encontrado en {temp_extract_folder}.")
    except Exception as e:
        log_error(f"Error al copiar 'setup.exe': {e}")

    finally:
        shutil.rmtree(temp_extract_folder, ignore_errors=True)
        print("Archivos temporales eliminados.")

def install_office(install_folder):
    file_path = os.path.join(install_folder, "selected_options.txt")

    try:
        options = read_selected_options(file_path)
    except Exception as e:
        log_error(f"Error al leer las opciones seleccionadas: {e}")
        return

    if "Version" in options:
        version_string = options["Version"]
        year = extract_year(version_string)

        if year:
            version = year

        download_and_extract_ODT(version, install_folder)
        xml_file_name = f"OfficeConfig{version}.xml"
        xml_path = os.path.join(files_dir, "ODT_ConfigXML", xml_file_name)

        if os.path.exists(xml_path):
            with open(xml_path, "r") as xml_file:
                xml_content = xml_file.read()

            architecture = format_architecture(options.get("Architecture", "Unknown"))
            modified_xml_content = xml_content.replace('OfficeClientEdition="64"', f'OfficeClientEdition="{architecture}"')

            edition = format_edition(version, options.get("Edition", "Unknown"))
            modified_xml_content = modified_xml_content.replace('Product ID="ProplusRetail"', f'Product ID="{edition}"')

            applications = format_applications(options.get("Applications", ""))
            app_names = applications.split("\n")

            for app_name in app_names:
                if app_name.strip():
                    exclusion_tag = f'<ExcludeApp ID="{app_name.strip()}" />'
                    modified_xml_content = modified_xml_content.replace(exclusion_tag, "")
                    
            modified_xml_content = '\n'.join(line for line in modified_xml_content.split('\n') if line.strip())
            modified_xml_path = os.path.join(install_folder, "configuration.xml")
            with open(modified_xml_path, "w") as modified_xml_file:
                modified_xml_file.write(modified_xml_content)

        else:
            log_error(f"El archivo {xml_file_name} no se encontró en la ubicación esperada.")

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

def check_and_create_install_folder():
    install_folder = os.path.join(files_dir, "InstallOfficeFiles")

    try:
        if os.path.exists(install_folder):
            shutil.rmtree(install_folder)

    except OSError as e:
        log_error(f"Error al crear el directorio {install_folder}: {e}")
        raise

    return install_folder

def check_and_create_uninstall_folder():
    uninstall_folder = os.path.join(files_dir, "Uninstall")
    try:
        if os.path.exists(uninstall_folder):
            shutil.rmtree(uninstall_folder)
            
    except OSError as e:
        log_error(f"Error al crear el directorio {uninstall_folder}: {e}")
        raise

    return uninstall_folder

def remove_temp_folder(install_folder, uninstall_folder):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    if not messagebox.askyesno("Eliminar archivos temporales", "¿Deseas eliminar los archivos temporales creados por el script?"):
        messagebox.showinfo("Cancelado", "No se eliminaron los archivos temporales.")
        return

    try:
        for folder in [install_folder, uninstall_folder]:
            if os.path.isdir(folder):
                shutil.rmtree(folder)
        
        messagebox.showinfo("Archivos eliminados", "Los archivos temporales han sido eliminados correctamente.")
    
    except Exception as e:
        messagebox.showwarning("Error", f"No se pudieron eliminar los archivos temporales: {e}")
        sys.exit(1)

    sys.exit(0)

def manage_installation():
    check_log_retention(os.path.join(files_dir, "logs", "log_error.txt"))

    install_folder = check_and_create_install_folder()
    uninstall_folder = check_and_create_uninstall_folder()

    detect_office = input("¿Desea detectar las versiones de Office instaladas? (S/N): ").strip().lower()
    if detect_office == 's':
        office_names = Detect_OfficeVersion(install_folder, uninstall_folder)
        uninstall_decision = ask_uninstall_and_proceed(install_folder, uninstall_folder, office_names)
    else:
        print(Fore.RED + "Advertencia: Tener más de una versión de Office instalada puede ocasionar problemas.")
        uninstall_decision = ask_proceed_installation(install_folder, uninstall_folder)

    if uninstall_decision:
        print("Instalación completada.")
    else:
        print("Proceso cancelado.")
        remove_temp_folder(install_folder, uninstall_folder)

def ask_uninstall_and_proceed(install_folder, uninstall_folder, office_names):
    uninstall_office = input("¿Desea desinstalar todas las versiones de Office encontradas? (S/N): ").strip().lower()
    if uninstall_office == 's':
        Uninstall_OfficeVersion(office_names)
        ask_proceed_installation(install_folder, uninstall_folder)
    else:
        print(Fore.RED + "Advertencia: Tener más de una versión de Office instalada puede ocasionar problemas.")
        ask_proceed_installation(install_folder, uninstall_folder)


def ask_proceed_installation(install_folder, uninstall_folder):
    proceed_install = input("¿Desea proceder con la instalación de Office? (S/N): ").strip().lower()
    if proceed_install == 's':
        print("Iniciando instalación de Office...")
        execute_installation_process(install_folder, uninstall_folder)
    else:
        print("Proceso cancelado.")
        remove_temp_folder(install_folder, uninstall_folder)

def execute_installation_process(install_folder, uninstall_folder):
    create_office_selection_window(install_folder, uninstall_folder)

    install_thread = threading.Thread(target=install_office, args=(install_folder,))
    install_thread.start()
    install_thread.join()

    run_setup_commands(install_folder)

    remove_temp_folder(install_folder, uninstall_folder)

def main():
    init(autoreset=True)
    manage_installation()

if __name__ == "__main__":
    main()
