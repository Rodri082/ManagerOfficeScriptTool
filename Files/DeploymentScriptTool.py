import os
import subprocess
import sys
import requests
from urllib.parse import urlparse
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import shutil
import threading
import datetime
from colorama import init, Fore, Style

#ODT_2013 = https://www.microsoft.com/en-us/download/details.aspx?id=36778
#ODT_2016_2019_2021 = https://www.microsoft.com/en-us/download/details.aspx?id=49117

ODT_2013 = "https://download.microsoft.com/download/6/2/3/6230F7A2-D8A9-478B-AC5C-57091B632FCF/officedeploymenttool_x86_5031-1000.exe"
ODT_2016_2019_2021 = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"


def log_error(error_message):
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    logs_folder = os.path.join(script_dir, "logs")

    if not os.path.exists(logs_folder):
        os.makedirs(logs_folder)

    log_path = os.path.join(logs_folder, "log.txt")

    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a") as log_file:
        log_file.write(f"{current_time} - {error_message}\n")

    check_log_retention(log_path)

def check_log_retention(log_path):
    if os.path.exists(log_path):
        log_creation_time = datetime.datetime.fromtimestamp(os.path.getctime(log_path))
        current_time = datetime.datetime.now()
        retention_period = datetime.timedelta(days=5)
        
        if current_time - log_creation_time >= retention_period:
            os.remove(log_path)

def check_and_create_install_folder():
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    install_folder = os.path.join(script_dir, "InstallOfficeFiles")

    if os.path.exists(install_folder):
        remove_install_folder(install_folder)

    try:
        os.mkdir(install_folder)
    except OSError as e:
        log_error(f"Error al crear la carpeta {install_folder}: {e}")
        sys.exit(1)
    except Exception as e:
        log_error(f"Error desconocido al crear la carpeta {install_folder}: {e}")
        sys.exit(1)

    return install_folder

def check_and_create_uninstall_folder():
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    uninstall_folder = os.path.join(script_dir, "Uninstall")

    if os.path.exists(uninstall_folder):
        return uninstall_folder

def remove_install_folder(install_folder):
    try:
        if os.path.exists(install_folder):
            shutil.rmtree(install_folder)

    except FileNotFoundError:
        pass
    except (PermissionError, OSError) as e:
        log_error(f"Error al eliminar la carpeta: {e}")
    except Exception as e:
        log_error(f"Error desconocido al elimina la carpeta: {e}")

def remove_uninstall_folder(uninstall_folder):
    try:
        if os.path.exists(uninstall_folder):
            shutil.rmtree(uninstall_folder)

    except FileNotFoundError:
        pass
    except (PermissionError, OSError) as e:
        log_error(f"Error al eliminar la carpeta: {e}")
    except Exception as e:
        log_error(f"Error desconocido al elimina la carpeta: {e}")

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
                remove_install_folder(install_folder)
                remove_uninstall_folder(uninstall_folder)
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

def extract_odt_files(version, install_folder):
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    local_exe_folder = os.path.join(script_dir, "InstallOfficeFiles")
    os.makedirs(local_exe_folder, exist_ok=True)

    url = ODT_2013 if version == "2013" else ODT_2016_2019_2021

    try:
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            parsed_url = urlparse(url)
            exe_file_name = os.path.basename(parsed_url.path)
            exe_file_path = os.path.join(local_exe_folder, exe_file_name)

            with open(exe_file_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=8192):
                    file.write(chunk)

            print(f"Archivo descargado correctamente: {exe_file_path}")
        else:
            log_error(f"Error al descargar el archivo desde {url}: {response.status_code}")
            return
    except Exception as e:
        log_error(f"Ocurrió un error al descargar {url}: {e}")
        return

    try:
        subprocess.run([exe_file_path, "/quiet", f"/extract:{install_folder}"], cwd=install_folder, check=True)
    except subprocess.CalledProcessError as e:
        log_error(f"Error al extraer archivos con el comando: {e}")
        return
    except Exception as e:
        log_error(f"Error desconocido al extraer archivos: {e}")
        return

    try:
        for file in os.listdir(install_folder):
            if file.endswith(".xml"):
                os.remove(os.path.join(install_folder, file))
    except OSError as e:
        log_error(f"Error al eliminar archivo XML en {install_folder}: {e}")
    except Exception as e:
        log_error(f"Error desconocido al eliminar archivo XML: {e}")

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

        extract_odt_files(version, install_folder)
        if getattr(sys, 'frozen', False):
            script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        xml_file_name = f"OfficeConfig{version}.xml"
        xml_path = os.path.join(script_dir, "ODT_ConfigXML", xml_file_name)

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
        print(Fore.RED + f"Por favor, no cierre esta ventana. Se Instalará la Versión de Microsoft Office Seleccionada..." + Style.RESET_ALL)
        subprocess.run(command, cwd=working_directory, stdout=sys.stdout, stderr=sys.stderr, check=True)
    except FileNotFoundError as e:
        log_error(f"El archivo setup.exe no se encuentra en la ubicación esperada: {e}")
    except Exception as e:
        log_error(f"Error al ejecutar comando: {e}")
        sys.exit(1)

def manage_installation():
    install_folder = check_and_create_install_folder()
    uninstall_folder = check_and_create_uninstall_folder()
    create_office_selection_window(install_folder, uninstall_folder)

    install_thread = threading.Thread(target=install_office, args=(install_folder,))
    install_thread.start()
    install_thread.join()

    run_setup_commands(install_folder)

    if messagebox.askyesno("Eliminar archivos temporales", "¿Deseas eliminar los archivos temporales creados por el script?"):
        try:
            if os.path.exists(install_folder):
                remove_install_folder(install_folder)
            if os.path.exists(uninstall_folder):
                remove_uninstall_folder(uninstall_folder)

            messagebox.showinfo("Archivos eliminados", "Los archivos temporales han sido eliminados.")
            sys.exit(0)

        except Exception as e:
            messagebox.showwarning("Error", f"No se pudieron eliminar los archivos temporales: {e}")
            sys.exit(1)

def main():
    init()
    manage_installation()

if __name__ == "__main__":
    main()