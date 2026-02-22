import logging
import sys
import tkinter as tk
from tkinter import messagebox

from contract_cleaner.logging import get_log_dir

def show_fatal_error(exc: BaseException):
    log_dir = get_log_dir()
    log_file = log_dir / "cleaner.log"

    logging.error("Unhandled exception", exc_info=exc)

    root = tk.Tk()
    root.withdraw()

    messagebox.showerror(
        "Contract Cleaner - Error",
        (
            "Something went wrong.\n\n"
            "No files were modified.\n\n"
            f"A log file was written to:\n\n{log_file}\n\n"
            "Please send this file to support."
        ),
    )

def install_exception_hook():
    def excepthook(exc_type, exc, tb):
        show_fatal_error(exc)
        sys.exit(1)

    sys.excepthook = excepthook
