from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path


def _configure_tcl_tk() -> None:
    app_root = Path(sys.executable).resolve().parent
    tcl_library = app_root / "tcl" / "tcl8.6"
    tk_library = app_root / "tcl" / "tk8.6"
    dll_directory = app_root / "DLLs"
    os.chdir(app_root)
    if tcl_library.exists():
        os.environ.setdefault("TCL_LIBRARY", "tcl/tcl8.6")
    if tk_library.exists():
        os.environ.setdefault("TK_LIBRARY", "tcl/tk8.6")
    if dll_directory.exists() and hasattr(os, "add_dll_directory"):
        os.add_dll_directory(str(dll_directory))


def _show_startup_error(message: str) -> None:
    error_path = Path.cwd() / "PFAS_CCS_Screening_startup_error.txt"
    error_path.write_text(message, encoding="utf-8")
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "PFAS CCS Screening",
            f"The application could not start.\n\nDetails: {error_path}",
        )
        root.destroy()
    except Exception:
        pass


_configure_tcl_tk()

if not os.environ.get("PFAS_APP_SKIP_LAUNCH"):
    try:
        from pfas_screening_app.gui import main

        main()
    except Exception:
        _show_startup_error(traceback.format_exc())
