"""
Point d'entrée du Générateur d'Assignation.

Usage :
    python main.py
"""

import tkinter as tk

from src.gui import AssignationApp


def main() -> None:
    root = tk.Tk()

    # Icône (si présente dans le dossier)
    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass

    app = AssignationApp(root)  # noqa: F841
    root.mainloop()


if __name__ == "__main__":
    main()
