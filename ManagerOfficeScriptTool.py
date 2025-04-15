# Módulos integrados
import logging
import platform
import re
import shutil
import subprocess
import tempfile
import time
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import winreg
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Tuple
from urllib.parse import urlparse

# Módulos de terceros
import requests
from colorama import Fore, init
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm
from webdriver_manager.chrome import ChromeDriverManager

# Constantes para URLs de herramientas
ODT_DOWNLOAD_ID_OFFICE_2013 = "36778"  # 2013
ODT_DOWNLOAD_ID_OFFICE_2016_AND_LATER = "49117"  # 2016, 2019, 2021LTSC, 2024LTSC y 365.


# Directorios temporales y rutas de trabajo
temp_dir = Path(tempfile.mkdtemp(prefix="ManagerOfficeScriptTool_"))
logs_folder = temp_dir / "logs"
office_install_dir = temp_dir / "InstallOfficeFiles"
office_uninstall_dir = temp_dir / "UninstallOfficeFiles"


def safe_log_path(path: Path) -> str:
    """
    Convierte rutas sensibles a una forma anonimizada para el log.
    Reemplaza la carpeta del usuario con %USERPROFILE%.
    """
    try:
        return str(path).replace(str(Path.home()), "%USERPROFILE%")
    except Exception:
        return str(path)


def safe_log_registry_key(reg_path: str) -> str:
    """
    Sanitiza rutas de registro para ocultar posibles datos sensibles.
    """
    return reg_path.replace("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\", "HKLM\\...\\")


def init_logging(logs_path: str):
    """
    Inicializa el sistema de logging creando el directorio y configurando el archivo de log.
    Usando pathlib para un manejo de rutas más pythonic.
    """
    # Convertir el string a un objeto Path y crear el directorio, incluyendo los padres, si es necesario.
    path = Path(logs_path)
    path.mkdir(parents=True, exist_ok=True)

    # Definir la ruta del archivo de log usando la sintaxis de pathlib
    log_file = path / "application.log"

    # Configurar el logging utilizando la ruta convertida a string
    logging.basicConfig(
        filename=str(log_file),
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )


def ask_yes_no(message: str) -> bool:
    """
    Pregunta al usuario en consola y retorna True si la respuesta es afirmativa.

    Args:
        message (str): Mensaje a mostrar al usuario.

    Returns:
        bool: True si la respuesta es afirmativa, False en caso contrario.
    """
    while True:
        respuesta = input(f"{message} (S/N): ").strip().lower()
        if respuesta in ("s", "sí", "si"):
            return True
        elif respuesta in ("n", "no"):
            return False
        else:
            print(Fore.YELLOW + "Respuesta no válida. Por favor ingresa 'S' o 'N'.")


def clean_temp_folders() -> None:
    """
    Elimina las carpetas temporales generadas por el script si el usuario lo confirma.
    Aplica sanitización a las rutas antes de mostrar o registrar mensajes.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if not messagebox.askyesno(
        "Eliminar archivos temporales",
        "¿Deseas eliminar las carpetas temporales creadas por el script?",
    ):
        messagebox.showinfo("Cancelado", "No se eliminaron las carpetas temporales.")
        root.destroy()
        return

    eliminadas = []
    errores = []

    for folder_path in [Path(office_install_dir), Path(office_uninstall_dir)]:
        sanitized_folder = safe_log_path(folder_path)

        if folder_path.is_dir():
            try:
                shutil.rmtree(folder_path)
                eliminadas.append(sanitized_folder)
                logging.info(f"Carpeta eliminada: {sanitized_folder}")
            except PermissionError:
                logging.error(
                    f"Permiso denegado al eliminar la carpeta: {sanitized_folder}"
                )
                errores.append(f"Permiso denegado: {sanitized_folder}")
            except FileNotFoundError:
                logging.warning(f"La carpeta ya no existe: {sanitized_folder}")
            except OSError as e:
                logging.exception(f"Error eliminando {sanitized_folder}: {e}")
                errores.append(f"{sanitized_folder}: {e}")

    if errores:
        messagebox.showwarning(
            "Error",
            "Ocurrieron errores al eliminar algunas carpetas:\n\n" + "\n".join(errores),
        )
    elif eliminadas:
        messagebox.showinfo(
            "Archivos eliminados",
            "Las carpetas temporales han sido eliminadas correctamente.",
        )
    else:
        messagebox.showinfo(
            "Nada que eliminar", "No se encontraron carpetas temporales para eliminar."
        )

    root.destroy()


class RegistryReader:
    """
    Clase para leer claves y valores del registro de Windows de forma eficiente.

    Utiliza un caché interno (_cache) para evitar lecturas repetidas de la misma clave.
    """

    def __init__(self) -> None:
        self._cache: Dict[Tuple[str, str], str] = {}

    def get_registry_keys(self, key: str) -> List[str]:
        """
        Devuelve las subclaves de una clave del registro.

        Args:
            key (str): Ruta de la clave.

        Returns:
            List[str]: Subclaves encontradas o lista vacía si hay error.
        """
        root_key = winreg.HKEY_LOCAL_MACHINE
        access_flag = winreg.KEY_READ | (
            winreg.KEY_WOW64_64KEY if platform.machine().endswith("64") else 0
        )

        subkeys: List[str] = []
        sanitized_key = safe_log_registry_key(key)
        try:
            with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
                index = 0
                while True:
                    try:
                        subkey = winreg.EnumKey(key_handle, index)
                        subkeys.append(subkey)
                        index += 1
                    except OSError:
                        break
        except FileNotFoundError:
            logging.warning(f"Clave del registro no encontrada: '{sanitized_key}'")
        except PermissionError:
            logging.error(f"Permiso denegado al acceder a la clave: '{sanitized_key}'")
        except OSError as e:
            logging.error(f"Error OS al abrir clave '{sanitized_key}': {e}")
        except Exception as e:
            logging.exception(
                f"Excepción inesperada al acceder a la clave '{sanitized_key}': {e}"
            )

        return subkeys

    def get_registry_value(self, key: str, value_name: str) -> str:
        """
        Obtiene un valor de una clave del registro (caché integrada).

        Args:
            key (str): Ruta de la clave.
            value_name (str): Nombre del valor.

        Returns:
            str: Valor encontrado o cadena vacía.
        """
        cache_key = (key, value_name)
        if cache_key in self._cache:
            return self._cache[cache_key]

        root_key = winreg.HKEY_LOCAL_MACHINE
        access_flag = winreg.KEY_READ | (
            winreg.KEY_WOW64_64KEY if platform.machine().endswith("64") else 0
        )

        sanitized_key = safe_log_registry_key(key)

        try:
            with winreg.OpenKey(root_key, key, 0, access_flag) as key_handle:
                try:
                    value, _ = winreg.QueryValueEx(key_handle, value_name)
                    self._cache[cache_key] = value
                    return value
                except FileNotFoundError:
                    logging.warning(
                        f"Valor '{value_name}' no encontrado en clave: '{sanitized_key}'"
                    )
                except PermissionError:
                    logging.error(
                        f"Permiso denegado para leer '{value_name}' en clave: '{sanitized_key}'"
                    )
                except OSError as e:
                    logging.error(
                        f"Error OS al leer '{value_name}' en clave '{sanitized_key}': {e}"
                    )
        except FileNotFoundError:
            logging.warning(
                f"Clave no encontrada al buscar valor '{value_name}': '{sanitized_key}'"
            )
        except PermissionError:
            logging.error(
                f"Permiso denegado al abrir clave '{sanitized_key}' para leer '{value_name}'"
            )
        except Exception as e:
            logging.exception(
                f"Excepción inesperada al leer '{value_name}' en clave '{sanitized_key}': {e}"
            )

        return ""


class OfficeInstallation:
    """
    Representa una instalación de Microsoft Office detectada en el sistema.

    Atributos:
        name (str): Nombre del producto Office (p.ej., "Microsoft Office 365").
        version (str): Versión instalada (p.ej., "16.0.12345.10000").
        install_path (str): Ruta de instalación en el sistema.
        click_to_run (bool): Indica si la instalación es del tipo Click-to-Run.
        product (str): ID del producto asociado a la instalación.
        bitness (str): Arquitectura (p.ej., "32-Bits" o "64-Bits").
        updates_enabled (bool): Indica si las actualizaciones están habilitadas.
        update_url (str): URL del canal de actualizaciones.
        client_culture (str): Idioma de la instalación.
        version_to_report (str): Versión informada del cliente.
        media_type (str): Tipo de medio de instalación.
        uninstall_string (str): Comando de desinstalación desde el registro.
    """

    def __init__(
        self,
        name: str,
        version: str,
        install_path: str,
        click_to_run: bool,
        product: str,
        bitness: str,
        updates_enabled: bool,
        update_url: str,
        client_culture: str,
        media_type: str,
        uninstall_string: str,
    ) -> None:
        self.name = name
        self.version = version
        self.install_path = install_path
        self.click_to_run = click_to_run
        self.uninstall_string = uninstall_string
        self.updates_enabled = updates_enabled
        self.bitness = bitness
        self.update_url = update_url
        self.media_type = media_type

        # Extraer ClientCulture desde UninstallString (buscando "culture=")
        match_culture = re.search(r"culture=([a-zA-Z\-]+)", uninstall_string)
        self.client_culture = (
            match_culture.group(1) if match_culture else client_culture
        )

        # Extraer Bitness desde UninstallString (buscando "platform=")
        match_platform = re.search(
            r"platform=(x86|x64)", uninstall_string, re.IGNORECASE
        )
        if match_platform:
            self.bitness = (
                "64-Bits" if match_platform.group(1).lower() == "x64" else "32-Bits"
            )

        # Extraer ProductID desde UninstallString (buscando "productstoremove=" y extrayendo el valor antes del guion bajo)
        match_product = re.search(r"productstoremove=([A-Za-z]+)", uninstall_string)
        self.product = match_product.group(1) if match_product else product

    def __repr__(self) -> str:
        """
        Devuelve una representación legible de la instalación,
        útil para depuración y visualización rápida.
        """
        return f"<OfficeInstallation: {self.name} ({self.version}) - {self.client_culture} - {self.bitness}>"


class OfficeManager:
    """
    Administra la detección y visualización de instalaciones de Microsoft Office.

    Args:
        show_all (bool): Si es True, mostrará todas las instalaciones encontradas.
                         Si es False, solo mostrará las instalaciones de Microsoft Office o 365.
    """

    def __init__(self, show_all: bool = False) -> None:
        self.show_all = show_all
        self.registry = RegistryReader()
        self.installations: List[OfficeInstallation] = []

    def _get_installations(self) -> List[OfficeInstallation]:
        """
        Busca y almacena las instalaciones de Microsoft Office detectadas en el sistema.

        Returns:
            List[OfficeInstallation]: Lista de objetos OfficeInstallation encontrados.
        """
        # Claves base del registro para Office y desinstalación
        office_keys = [
            r"SOFTWARE\Microsoft\Office",
            r"SOFTWARE\Wow6432Node\Microsoft\Office",
        ]
        uninstall_keys = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        ]
        # Versiones a buscar (incluye ClickToRun y versiones numéricas)
        versions = ["ClickToRun", "15.0", "16.0"]
        found_names: set[str] = set()
        installations: List[OfficeInstallation] = []

        for office_key in office_keys:
            for version in versions:
                # Construir la clave del registro dependiendo si es ClickToRun o no.
                if version == "ClickToRun":
                    version_key = str(Path(office_key) / version / "Configuration")
                else:
                    version_key = str(
                        Path(office_key) / version / "ClickToRun" / "Configuration"
                    )

                # Obtener valores generales de la instalación
                platform_value = self.registry.get_registry_value(
                    version_key, "Platform"
                )
                if not platform_value:
                    continue

                updates_enabled = (
                    self.registry.get_registry_value(version_key, "UpdatesEnabled")
                    == "True"
                )
                update_url = self.registry.get_registry_value(version_key, "CDNBaseUrl")

                product = self.registry.get_registry_value(
                    version_key, "ProductReleaseIds"
                )
                media_type = self.registry.get_registry_value(
                    version_key, f"{product}.MediaType"
                )

                # Buscar entradas de desinstalación para asociarlas a la instalación Office.
                for uninstall_key in uninstall_keys:
                    subkeys = self.registry.get_registry_keys(uninstall_key)
                    for subkey in subkeys:
                        uninstall_key_path = str(Path(uninstall_key) / subkey)
                        display_name = self.registry.get_registry_value(
                            uninstall_key_path, "DisplayName"
                        )
                        if not display_name or display_name in found_names:
                            continue

                        # Si show_all es False, se filtran las instalaciones que no contengan Office o Microsoft 365.
                        if not self.show_all and not (
                            "Microsoft Office" in display_name
                            or "Microsoft 365" in display_name
                        ):
                            continue

                        found_names.add(display_name)
                        display_version = self.registry.get_registry_value(
                            uninstall_key_path, "DisplayVersion"
                        )
                        install_location = self.registry.get_registry_value(
                            uninstall_key_path, "InstallLocation"
                        )
                        uninstall_string = self.registry.get_registry_value(
                            uninstall_key_path, "UninstallString"
                        )
                        click_to_run = "ClickToRun" in uninstall_string

                        # Crear la instancia OfficeInstallation
                        installations.append(
                            OfficeInstallation(
                                name=display_name,
                                version=display_version,
                                install_path=install_location,
                                click_to_run=click_to_run,
                                product="",
                                bitness="",
                                updates_enabled=updates_enabled,
                                update_url=update_url,
                                client_culture="",
                                media_type=media_type,
                                uninstall_string=uninstall_string,
                            )
                        )

        self.installations = installations
        return installations

    def display_installations(self) -> None:
        """
        Imprime por consola las instalaciones de Office encontradas con formato visual.
        """
        installations = self._get_installations()
        if not installations:
            print(Fore.YELLOW + "No se encontraron instalaciones de Microsoft Office.")
            return

        print(Fore.CYAN + "=" * 110)
        print(
            Fore.GREEN
            + "Se encontraron las siguientes instalaciones de Microsoft Office:"
        )
        print(Fore.CYAN + "=" * 110)

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
                "MediaType": install.media_type,
            }
            for key, value in office_info.items():
                print(Fore.LIGHTWHITE_EX + f"{key:<20}: {str(value)}")
            print(Fore.CYAN + "=" * 110)

    def get_installations(self) -> List[OfficeInstallation]:
        """
        Retorna la lista de instalaciones encontradas.

        Returns:
            List[OfficeInstallation]: Lista de instalaciones actuales.
        """
        return self._get_installations()


class ODTManager:
    def __init__(self, office_install_dir: str) -> None:
        if not office_install_dir:
            raise ValueError("El directorio de instalación no puede ser None.")
        self.office_dir = office_install_dir

    def get_download_url_from_microsoft(self, version_identifier: str) -> str | None:
        version_id = (
            ODT_DOWNLOAD_ID_OFFICE_2013
            if "2013" in version_identifier
            else ODT_DOWNLOAD_ID_OFFICE_2016_AND_LATER
        )
        url_page = (
            f"https://www.microsoft.com/en-us/download/details.aspx?id={version_id}"
        )

        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")

        try:
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=options
            )
            driver.get(url_page)

            # Esperar a que window.__DLCDetails__ esté disponible
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script(
                    "return typeof window.__DLCDetails__ !== 'undefined'"
                )
            )

            # Obtener el objeto directamente desde JS
            dlc_details = driver.execute_script("return window.__DLCDetails__")

            file_info = dlc_details["dlcDetailsView"]["downloadFile"][0]
            url = file_info["url"]
            self.expected_name = file_info["name"]
            self.expected_size = int(file_info["size"])

            parsed = urlparse(url)
            if parsed.scheme != "https" or parsed.netloc != "download.microsoft.com":
                logging.error(
                    f"Dominio o esquema no válido: {parsed.scheme}://{parsed.netloc}"
                )
                return None

            logging.info(f"URL válida obtenida: {url}")
            return url

        except Exception as e:
            logging.exception(
                "Error al obtener la URL de descarga desde window.__DLCDetails__:"
            )
            return None
        finally:
            try:
                driver.quit()
            except Exception:
                pass

    def download_and_extract(
        self, version_identifier: str, max_retries: int = 3
    ) -> bool:
        if not self.office_dir:
            logging.error("El directorio de instalación no está definido.")
            return False

        url = self.get_download_url_from_microsoft(version_identifier)
        if not url:
            logging.error("No se pudo obtener la URL de descarga.")
            return False

        office_dir = Path(self.office_dir)
        office_dir.mkdir(parents=True, exist_ok=True)

        exe_file_name = Path(urlparse(url).path).name
        exe_file_path = office_dir / exe_file_name
        sanitized_path = safe_log_path(exe_file_path)
        sanitized_dir = safe_log_path(self.office_dir)

        for intento in range(1, max_retries + 1):
            try:
                # Si ya existe y es válido, usarlo
                if exe_file_path.exists():
                    actual_size = exe_file_path.stat().st_size
                    if (
                        exe_file_path.name == self.expected_name
                        and actual_size == self.expected_size
                    ):
                        logging.info(
                            f"Archivo ya descargado y válido: {sanitized_path}"
                        )
                        break  # Salir del bucle, está todo bien
                    else:
                        logging.warning(
                            f"Archivo existente inválido. Eliminando: {sanitized_path}"
                        )
                        exe_file_path.unlink()

                logging.info(f"Intento {intento}: Descargando ODT desde: {url}")
                response = requests.get(url, stream=True, timeout=30)
                response.raise_for_status()

                total_size = int(response.headers.get("Content-Length", 0))
                with open(exe_file_path, "wb") as file, tqdm(
                    total=total_size,
                    unit="B",
                    unit_scale=True,
                    desc=f"Intento {intento} - Descargando ODT",
                ) as bar:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            file.write(chunk)
                            bar.update(len(chunk))

                # Validar archivo descargado
                actual_size = exe_file_path.stat().st_size
                if (
                    exe_file_path.name == self.expected_name
                    and actual_size == self.expected_size
                ):
                    logging.info(f"Archivo descargado correctamente: {sanitized_path}")
                    break  # Salir del bucle si todo está bien
                else:
                    logging.error("Archivo descargado no válido. Reintentando...")
                    exe_file_path.unlink()

            except (requests.RequestException, OSError) as e:
                logging.warning(f"Error durante la descarga (intento {intento}): {e}")
                if intento < max_retries:
                    time.sleep(5)  # Esperar antes del siguiente intento
                else:
                    logging.error(
                        "Se alcanzó el número máximo de reintentos de descarga."
                    )
                    return False

        # Extraer ODT
        try:
            command = [str(exe_file_path), "/quiet", f"/extract:{self.office_dir}"]
            subprocess.run(
                command,
                cwd=str(self.office_dir),
                capture_output=True,
                text=True,
                check=True,
            )
            logging.info(f"ODT extraído correctamente en {sanitized_dir}")
            return True

        except subprocess.CalledProcessError as e:
            logging.error(f"Error al ejecutar el ODT:\n{e.stderr}")
        except requests.RequestException as e:
            logging.error(f"Error en la descarga de ODT: {e}")
        except PermissionError:
            logging.error("Permiso denegado al intentar escribir en el disco.")
        except OSError as e:
            logging.error(f"Error del sistema de archivos: {e}")
        except Exception:
            logging.exception(
                "Error inesperado durante la descarga y extracción de ODT."
            )

        return False


class OfficeUninstaller:
    """
    Encargado de desinstalar una instalación específica de Microsoft Office
    utilizando el Office Deployment Tool (ODT).

    Args:
        office_uninstall_dir (str): Directorio temporal donde se descargará y ejecutará ODT.
        installation (OfficeInstallation): Objeto que representa la instalación de Office
            que será eliminada.
    """

    def __init__(
        self, office_uninstall_dir: str, installation: OfficeInstallation
    ) -> None:
        self.office_uninstall_dir = office_uninstall_dir
        self.installation = installation
        self.setup_path: Path | None = None

        # Log para confirmar que se está usando la instalación correcta
        logging.info(
            f"Inicializando OfficeUninstaller para: {self.installation.name} | "
            f"Product: {self.installation.product} | "
            f"Idioma: {self.installation.client_culture}"
        )

    def _generar_configuracion_remocion(self) -> str:
        """
        Genera un archivo XML de configuración para desinstalar la instalación de Office actual.

        Returns:
            str: Ruta del archivo de configuración generado.
        """
        xml_content = f"""<Configuration>
 <Remove>
    <Product ID="{self.installation.product}">
      <Language ID="{self.installation.client_culture}" />
    </Product>
 </Remove>
</Configuration>"""

        uninstall_dir = Path(self.office_uninstall_dir)
        file_path = uninstall_dir / "configuration.xml"
        sanitized_uninstall_path = safe_log_path(uninstall_dir)
        sanitized_file_path = safe_log_path(file_path)
        try:
            uninstall_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            logging.error(
                f"No se pudo crear el directorio de desinstalación: {sanitized_uninstall_path}"
            )
            raise

        try:
            file_path.write_text(xml_content, encoding="utf-8")
            return str(file_path)
        except Exception as e:
            logging.exception(
                f"Error al generar el archivo XML de desinstalación en {sanitized_file_path}"
            )
            raise

    def ejecutar_desinstalacion(self) -> bool:
        """
        Ejecuta la desinstalación de Office usando el setup.exe y el archivo XML generado.

        Returns:
            bool: True si la desinstalación fue exitosa, False si falló.
        """
        odt_manager = ODTManager(self.office_uninstall_dir)
        version_identifier = self.installation.name

        print(
            Fore.GREEN
            + f"Iniciando proceso de desinstalación para: {version_identifier}"
        )
        if not odt_manager.download_and_extract(version_identifier):
            print(Fore.RED + "Error al descargar y extraer ODT.")
            logging.error("Fallo la descarga o extracción del ODT.")
            return False

        office_uninstall_path = Path(self.office_uninstall_dir)
        self.setup_path = office_uninstall_path / "setup.exe"
        sanitized_uninstall_path = safe_log_path(office_uninstall_path)

        if not self.setup_path.exists():
            print(
                Fore.RED + f"No se encontró 'setup.exe' en {sanitized_uninstall_path}."
            )
            logging.error(f"'setup.exe' no encontrado en {sanitized_uninstall_path}")
            return False

        try:
            config_path = self._generar_configuracion_remocion()
        except Exception:
            print(Fore.RED + "Error al generar el archivo de configuración XML.")
            return False

        try:
            subprocess.run(
                [str(self.setup_path), "/configure", config_path],
                cwd=str(office_uninstall_path),
                capture_output=True,
                text=True,
                check=True,
            )

            logging.info(f"Office desinstalado: {self.installation.name}")
            return True

        except subprocess.CalledProcessError as e:
            logging.error(
                f"setup.exe falló con código {e.returncode}.\nStderr:\n{e.stderr}"
            )
        except Exception as e:
            logging.exception(f"Error inesperado durante la desinstalación. {e}")

        return False

    def execute(self) -> str:
        """
        Punto de entrada externo para iniciar la desinstalación.

        Returns:
            str: Mensaje de estado coloreado según el resultado.
        """
        if self.ejecutar_desinstalacion():
            return f"{Fore.GREEN}Desinstalado: {self.installation.name}"
        else:
            return f"{Fore.RED}Error al desinstalar: {self.installation.name}"


class OfficeInstaller:
    """
    Encargado de ejecutar el proceso de instalación de Microsoft Office
    utilizando el archivo setup.exe y configuration.xml generados previamente.

    Args:
        office_install_dir (str): Ruta donde se encuentra el setup.exe y el archivo de
            configuración XML.
    """

    def __init__(self, office_install_dir: str) -> None:
        """
        Inicializa la instancia con la ruta de instalación de Office.

        Args:
            office_install_dir (str): Carpeta que contiene los archivos necesarios para
                instalar Office.
        """
        self.office_install_dir = office_install_dir

    def run_setup_commands(self) -> None:
        """
        Ejecuta el comando de instalación de Office utilizando setup.exe y configuration.xml.

        Valida la existencia de los archivos requeridos y maneja cualquier excepción
        durante la ejecución del proceso.
        """
        office_dir = Path(self.office_install_dir)
        setup_path = office_dir / "setup.exe"
        config_path = office_dir / "configuration.xml"
        sanitized_setup_path = safe_log_path(setup_path)
        sanitized_config_path = safe_log_path(config_path)

        if not setup_path.exists():
            logging.error("No se encontró 'setup.exe': %s", sanitized_setup_path)
            print(Fore.RED + "Error: 'setup.exe' no está disponible.")
            return

        if not config_path.exists():
            logging.error(
                "No se encontró 'configuration.xml': %s", sanitized_config_path
            )
            print(Fore.RED + "Error: 'configuration.xml' no está disponible.")
            return

        command = [str(setup_path), "/configure", str(config_path)]

        try:
            print(
                Fore.YELLOW
                + "Instalando Microsoft Office. Por favor, no cierre esta ventana..."
            )

            subprocess.run(
                command,
                cwd=str(office_dir),
                capture_output=True,
                text=True,
                check=True,
            )

            print(Fore.GREEN + "Instalación completada exitosamente.")
            logging.info("Instalación de Office finalizada correctamente.")

        except subprocess.CalledProcessError as e:
            print(Fore.RED + "La instalación falló.")
            logging.error(
                "setup.exe falló con código %d\nComando: %s\nStderr:\n%s",
                e.returncode,
                f"{sanitized_setup_path} /configure {sanitized_config_path}",
                e.stderr.strip(),
            )

        except PermissionError:
            logging.error("Permiso denegado al ejecutar setup.exe.")
            print(Fore.RED + "Permiso denegado al ejecutar la instalación.")

        except OSError as e:
            logging.error(
                f"Error del sistema operativo al ejecutar la instalación: {e}"
            )
            print(Fore.RED + f"Error del sistema al iniciar la instalación")

        except Exception as e:
            logging.exception("Error inesperado durante la instalación.")
            print(Fore.RED + "Ocurrió un error inesperado durante la instalación.")


class OfficeSelectionWindow:
    """
    Ventana gráfica para seleccionar opciones de instalación de Microsoft Office.

    Permite al usuario escoger la versión, arquitectura, idioma, y aplicaciones a instalar,
    generando luego un archivo de configuración XML compatible con Office Deployment Tool (ODT).
    """

    def __init__(self) -> None:
        """
        Inicializa la interfaz gráfica del selector de versiones de Office.

        Carga los datos de versiones disponibles junto con las aplicaciones
        asociadas a cada versión, y prepara la estructura base de la ventana.
        """
        # Diccionario con las apps disponibles para cada versión de Office
        self.all_apps: dict[str, list[str]] = {
            "Office Standard 2013": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "Groove",
            ],
            "Office Professional Plus 2013": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "Access",
                "InfoPath",
                "Lync",
                "Groove",
            ],
            "Office Standard 2016": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "OneDrive",
                "Groove",
            ],
            "Office Professional Plus 2016": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "OneDrive",
                "Access",
                "Lync",
                "Groove",
            ],
            "Office Standard 2019": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "OneDrive",
                "Groove",
            ],
            "Office Professional Plus 2019": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "OneDrive",
                "Access",
                "Lync",
                "Groove",
            ],
            "Office LTSC Standard 2021": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "OneDrive",
            ],
            "Office LTSC Professional Plus 2021": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "OneDrive",
                "Access",
                "Teams",
                "Lync",
            ],
            "Office LTSC Standard 2024": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "OneDrive",
            ],
            "Office LTSC Professional Plus 2024": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "OneDrive",
                "Access",
                "Lync",
            ],
            "Microsoft 365 HomePremium - Current": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Publisher",
                "Access",
                "OneDrive",
            ],
            "Microsoft 365 Apps for business - Current": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Teams",
                "Publisher",
                "Access",
                "Lync",
                "OneDrive",
                "Groove",
            ],
            "Microsoft 365 Apps for business - MonthlyEnterprise": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Teams",
                "Publisher",
                "Access",
                "Lync",
                "OneDrive",
                "Groove",
            ],
            "Microsoft 365 Apps for business - SemiAnnual": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Teams",
                "Publisher",
                "Access",
                "Lync",
                "OneDrive",
                "Groove",
            ],
            "Microsoft 365 Apps for enterprise - Current": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Teams",
                "Publisher",
                "Access",
                "Lync",
                "OneDrive",
                "Groove",
            ],
            "Microsoft 365 Apps for enterprise - MonthlyEnterprise": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Teams",
                "Publisher",
                "Access",
                "Lync",
                "OneDrive",
                "Groove",
            ],
            "Microsoft 365 Apps for enterprise - SemiAnnual": [
                "Word",
                "Excel",
                "PowerPoint",
                "Outlook",
                "OneNote",
                "Teams",
                "Publisher",
                "Access",
                "Lync",
                "OneDrive",
                "Groove",
            ],
        }

        # Diccionario con el ID de producto y canal correspondiente a cada versión
        self.versiones: dict[str, tuple[str, str]] = {
            "Office Standard 2013": ("StandardRetail", "Current"),
            "Office Professional Plus 2013": ("ProplusRetail", "Current"),
            "Office Standard 2016": ("StandardRetail", "Current"),
            "Office Professional Plus 2016": ("ProplusRetail", "Current"),
            "Office Standard 2019": ("Standard2019Volume", "PerpetualVL2019"),
            "Office Professional Plus 2019": ("ProPlus2019Volume", "PerpetualVL2019"),
            "Office LTSC Standard 2021": ("Standard2021Volume", "PerpetualVL2021"),
            "Office LTSC Professional Plus 2021": (
                "ProPlus2021Volume",
                "PerpetualVL2021",
            ),
            "Office LTSC Standard 2024": ("Standard2024Volume", "PerpetualVL2024"),
            "Office LTSC Professional Plus 2024": (
                "ProPlus2024Volume",
                "PerpetualVL2024",
            ),
            "Microsoft 365 HomePremium - Current": ("O365HomePremRetail", "Current"),
            "Microsoft 365 Apps for business - Current": (
                "O365BusinessRetail",
                "Current",
            ),
            "Microsoft 365 Apps for business - MonthlyEnterprise": (
                "O365BusinessRetail",
                "MonthlyEnterprise",
            ),
            "Microsoft 365 Apps for business - SemiAnnual": (
                "O365BusinessRetail",
                "SemiAnnual",
            ),
            "Microsoft 365 Apps for enterprise - Current": (
                "O365ProPlusRetail",
                "Current",
            ),
            "Microsoft 365 Apps for enterprise - MonthlyEnterprise": (
                "O365ProPlusRetail",
                "MonthlyEnterprise",
            ),
            "Microsoft 365 Apps for enterprise - SemiAnnual": (
                "O365ProPlusRetail",
                "SemiAnnual",
            ),
        }

        self.root = tk.Tk()
        self.app_vars: dict[str, tk.BooleanVar] = {}
        self.app_checkbuttons: list[ttk.Checkbutton] = []
        self.cancelled = False

    def on_closing(self) -> None:
        """
        Manejador para el evento de cierre de la ventana.
        Cancela la instalación y limpia los temporales.
        """
        messagebox.showwarning("Advertencia", "Instalación de Office cancelada...")
        self.cancelled = True
        self.root.destroy()

    def update_apps(self, event=None) -> None:
        """
        Actualiza la lista de aplicaciones mostradas según la versión seleccionada.
        """
        selected_version = self.combo_version.get()
        available_apps = self.all_apps.get(selected_version, [])

        # Limpiar checkbuttons anteriores
        for widget in self.app_checkbuttons:
            widget.grid_forget()
        self.app_vars.clear()
        self.app_checkbuttons.clear()

        # Crear nuevos checkbuttons para las apps disponibles
        for i, app in enumerate(available_apps):
            var = tk.BooleanVar()
            self.app_vars[app] = var
            cb = ttk.Checkbutton(self.frame_apps, text=app, variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, pady=3)
            self.app_checkbuttons.append(cb)

    def generate_configuration(
        self,
        selected_version: str,
        bits: str,
        language: str,
        remove_msi: bool,
        selected_apps: list[str],
    ) -> str | None:
        """
        Genera el archivo configuration.xml con las opciones seleccionadas por el usuario.

        Returns:
            str | None: Ruta del archivo de configuración generado o None en caso de error.
        """

        self.root.destroy()

        odt_manager = ODTManager(office_install_dir)

        if not odt_manager.download_and_extract(selected_version):
            messagebox.showerror("Error", "No se pudo descargar y extraer ODT.")
            return None

        if selected_version not in self.versiones:
            logging.error(f"Versión seleccionada no válida: {selected_version}")
            messagebox.showerror("Error", "Versión seleccionada no válida.")
            return None

        available_apps = self.all_apps.get(selected_version, [])
        apps_to_exclude = [app for app in available_apps if app not in selected_apps]
        exclude_apps = (
            "\n".join(
                f'            <ExcludeApp ID="{app}" />' for app in apps_to_exclude
            )
            if apps_to_exclude
            else ""
        )
        remove_msi_tag = "\n    <RemoveMSI />" if remove_msi else ""

        config = f"""<Configuration>
    <Add OfficeClientEdition="{bits}" Channel="{self.versiones[selected_version][1]}">
        <Product ID="{self.versiones[selected_version][0]}">
            <Language ID="{language}" />{("\n" + exclude_apps) if exclude_apps else ""}
        </Product>
    </Add>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="AUTOACTIVATE" Value="1" />
    <Updates Enabled="TRUE" />{remove_msi_tag}
    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>"""

        config_file_path = Path(office_install_dir) / "configuration.xml"
        sanitized_config_path = safe_log_path(config_file_path)

        try:
            config_file_path.write_text(config, encoding="utf-8")
            print(
                Fore.GREEN
                + f"Archivo de configuración generado exitosamente en: {sanitized_config_path}"
            )
            return str(config_file_path)

        except PermissionError:
            logging.error(
                f"Permiso denegado al escribir archivo: {sanitized_config_path}"
            )
            messagebox.showerror("Error", "Permiso denegado al guardar el archivo.")
        except OSError as e:
            logging.error(
                f"Error del sistema de archivos al guardar {sanitized_config_path}: {e}"
            )
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{e}")
        except Exception as e:
            logging.exception(
                "Error inesperado al escribir el archivo de configuración:"
            )
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")

        return None

    def install_office(self) -> None:
        """
        Valida los campos seleccionados por el usuario y genera el archivo de configuración.
        """
        selected_version = self.combo_version.get()
        bits = self.combo_arch.get()
        language = self.combo_language.get()
        remove_msi = self.remove_msi_var.get()
        selected_apps = [app for app, var in self.app_vars.items() if var.get()]

        if not selected_apps:
            messagebox.showwarning(
                "Advertencia",
                "Por favor selecciona al menos una aplicación a instalar.",
            )
            return
        if not selected_version or not bits or not language:
            messagebox.showwarning(
                "Advertencia",
                "Por favor selecciona una versión, arquitectura y lenguaje.",
            )
            return
        if selected_version not in self.versiones:
            messagebox.showerror(
                "Error", "No se encontró la configuración para la versión seleccionada."
            )
            return

        self.generate_configuration(
            selected_version, bits, language, remove_msi, selected_apps
        )

    def show(self) -> None:
        """
        Muestra la ventana de selección de Office con todos los componentes gráficos.
        """
        self.root.title("Selector de Versiones de Microsoft Office")
        self.root.geometry("380x600")
        self.root.config(bg="#f4f4f4")
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", lambda: self.on_closing())

        # Crear estructura de la interfaz
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

        ttk.Label(frame, text="Selecciona la versión de Office:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.combo_version = ttk.Combobox(frame, width=50, state="readonly")
        self.combo_version["values"] = list(self.all_apps.keys())
        self.combo_version.set("Office Standard 2013")
        self.combo_version.grid(row=1, column=0, sticky=tk.W, pady=5)
        self.combo_version.bind("<<ComboboxSelected>>", self.update_apps)

        ttk.Label(frame, text="Selecciona la arquitectura:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.combo_arch = ttk.Combobox(
            frame, width=10, state="readonly", values=["32", "64"]
        )
        self.combo_arch.set("64")
        self.combo_arch.grid(row=3, column=0, sticky=tk.W, pady=5)

        ttk.Label(frame, text="Selecciona el idioma:").grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        self.combo_language = ttk.Combobox(
            frame, width=20, state="readonly", values=["es-es", "en-us"]
        )
        self.combo_language.set("es-es")
        self.combo_language.grid(row=5, column=0, sticky=tk.W, pady=5)

        self.remove_msi_var = tk.BooleanVar()
        ttk.Checkbutton(
            frame, text="Agregar RemoveMSI", variable=self.remove_msi_var
        ).grid(row=6, column=0, sticky=tk.W, pady=5)

        self.frame_apps = ttk.Frame(frame)
        self.frame_apps.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)

        self.app_vars = {}
        self.app_checkbuttons = []
        self.update_apps()

        ttk.Button(frame, text="Instalar Office", command=self.install_office).grid(
            row=17, column=0, pady=10
        )

        self.root.mainloop()
        return self.cancelled


def run_uninstallers(
    installations: List[OfficeInstallation], uninstall_dir: Path
) -> None:
    """
    Ejecuta la desinstalación de múltiples instalaciones de Office utilizando ODT.

    Args:
        installations (List[OfficeInstallation]): Lista de instalaciones a desinstalar.
        uninstall_dir (Path): Directorio temporal para ODT.
    """
    with ThreadPoolExecutor(max_workers=1) as executor:
        futures = [
            executor.submit(OfficeUninstaller(uninstall_dir, inst).execute)
            for inst in installations
        ]
        for future in as_completed(futures):
            result = future.result()
            if result:
                print(result)


def main() -> None:
    """
    Función principal del script.

    Ejecuta el flujo general de detección, desinstalación e instalación de Microsoft Office.
    - Detecta instalaciones existentes.
    - Ofrece la opción de desinstalar versiones encontradas con ODT.
    - Permite configurar e instalar una nueva versión de Office.
    - Gestiona errores y limpia archivos temporales al finalizar.
    """
    try:
        init_logging(logs_folder)
        init(autoreset=True)

        if ask_yes_no("¿Desea detectar las versiones de Office instaladas?"):
            manager = OfficeManager(show_all=False)
            installations = manager.get_installations()
            manager.display_installations()

            if installations:
                if len(installations) == 1:
                    if ask_yes_no("¿Desea desinstalar la versión encontrada con ODT?"):
                        run_uninstallers(installations, office_uninstall_dir)
                    else:
                        print(
                            Fore.RED
                            + "Advertencia: Tener múltiples versiones puede causar conflictos."
                        )
                else:
                    print(Fore.LIGHTWHITE_EX + "Seleccione una opción:")
                    print(Fore.LIGHTWHITE_EX + "1 - No desinstalar ninguna versión")
                    print(
                        Fore.LIGHTWHITE_EX
                        + "2 - Desinstalar todas las versiones encontradas"
                    )
                    print(
                        Fore.LIGHTWHITE_EX
                        + "3 - Elegir una versión específica para desinstalar"
                    )
                    opcion = input(Fore.LIGHTWHITE_EX + "Opción (1/2/3): ").strip()

                    if opcion == "2":
                        run_uninstallers(installations, office_uninstall_dir)

                    elif opcion == "3":
                        print(
                            Fore.LIGHTWHITE_EX
                            + "Seleccione la versión que desea desinstalar:"
                        )
                        for idx, install in enumerate(installations, 1):
                            print(Fore.LIGHTCYAN_EX + f"{idx} - {install.name}")

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
                                + f"Versión seleccionada: {seleccionada.name} ({seleccionada.client_culture})"
                            )
                            run_uninstallers([seleccionada], office_uninstall_dir)
                        else:
                            print(
                                Fore.YELLOW
                                + "Selección inválida. No se realizará ninguna desinstalación."
                            )
                    else:
                        print(Fore.YELLOW + "No se realizará ninguna desinstalación.")
        else:
            print(
                Fore.RED
                + "Advertencia: Si tiene más de una versión de Office instalada, podría ocasionar problemas."
            )

        if ask_yes_no("¿Desea proceder con una nueva instalación de Office?"):
            print("Iniciando configuración de instalación de Office...")
            selection_window = OfficeSelectionWindow()
            cancelled = selection_window.show()

            if cancelled:
                print(Fore.YELLOW + "El usuario canceló la instalación.")
            else:
                installer = OfficeInstaller(office_install_dir)
                installer.run_setup_commands()
        else:
            print(Fore.YELLOW + "Proceso cancelado por el usuario.")

    except KeyboardInterrupt:
        logging.warning("Ejecución interrumpida por el usuario (Ctrl+C).")
        print(Fore.YELLOW + "\nProceso cancelado manualmente.")

    except Exception as e:
        logging.exception("Ocurrió un error inesperado en la ejecución principal.")
        print(Fore.RED + f"Error crítico: {e}")

    finally:
        clean_temp_folders()


if __name__ == "__main__":
    main()
