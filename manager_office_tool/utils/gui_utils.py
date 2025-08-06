"""
gui_utils.py

Incluye funciones para centrar ventanas GUI.
"""


def center_window(win, width=300, height=200) -> None:
    """
    Centra una ventana de Tkinter en la pantalla.

    Args:
        win: Ventana de Tkinter.
        width (int): Ancho de la ventana.
        height (int): Alto de la ventana.
    """
    win.update_idletasks()
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (width // 2)
    y = (hs // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")
