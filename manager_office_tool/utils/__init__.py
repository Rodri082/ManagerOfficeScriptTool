"""
Subpaquete utils de ManagerOfficeScriptTool.

Contiene utilidades generales para consola, GUI, logging y manejo seguro de
rutas y carpetas temporales.
"""

from .console_utils import (
    ask_menu_option,
    ask_multiple_valid_indices,
    ask_single_valid_index,
    ask_yes_no,
)
from .gui_utils import center_window
from .logging_utils import init_logging
from .path_utils import (
    clean_folders,
    ensure_subfolder,
    get_data_path,
    get_temp_dir,
    safe_log_path,
    safe_log_registry_key,
)

__all__ = [
    "ask_yes_no",
    "ask_menu_option",
    "ask_single_valid_index",
    "ask_multiple_valid_indices",
    "center_window",
    "init_logging",
    "clean_folders",
    "ensure_subfolder",
    "get_data_path",
    "get_temp_dir",
    "safe_log_path",
    "safe_log_registry_key",
]
