import tkinter as tk
from pathlib import Path
from tkinter import messagebox

from .path_utils import clean_folders


def center_window(win, width=300, height=200):
    win.update_idletasks()
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (width // 2)
    y = (hs // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")


def clean_temp_folders_ui(folders: list[Path]) -> None:

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if not messagebox.askyesno(
        "Eliminar archivos temporales",
        "¿Deseas eliminar las carpetas temporales creadas por el script?",
    ):
        messagebox.showinfo(
            "Cancelado", "No se eliminaron las carpetas temporales."
        )
        root.destroy()
        return

    eliminadas, errores = clean_folders(folders)

    if errores:
        messagebox.showwarning(
            "Error",
            "Ocurrieron errores al eliminar algunas carpetas:\n\n"
            + "\n".join(errores),
        )
    elif eliminadas:
        messagebox.showinfo(
            "Archivos eliminados",
            "Las carpetas temporales han sido eliminadas correctamente.",
        )
    else:
        messagebox.showinfo(
            "Nada que eliminar",
            "No se encontraron carpetas temporales para eliminar.",
        )

    root.destroy()
