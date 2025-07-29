#!/usr/bin/env python3
"""
syndrom_db_ui.py
----------------
A simple Tkinter UI for managing the SyndromDB.

Features:
1. Enter Syndrom name (folder will be auto-sanitized for Windows)
2. Enter Monitor name
3. Browse and select Golden and Defect images
4. Enter a free-form description
5. Save button creates/updates folder inside `SyndromDB/` with:
   - golden.jpg
   - defect.jpg
   - description.txt  (first line contains Monitor name, rest is description)

Run the script from the project root:
    python syndrom_db_ui.py
"""

import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# Path to the SyndromDB directory (relative to this script)
SYNDROM_DB_PATH = os.path.join(os.path.dirname(__file__), "SyndromDB")

# Regex pattern for characters that are illegal in Windows folder names
ILLEGAL_CHARS = r"[\\/:*?\"<>|]"

def sanitize_name(name: str) -> str:
    """Replace illegal path characters with a hyphen."""
    return re.sub(ILLEGAL_CHARS, "-", name).strip()

class SyndromDBUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SyndromDB Manager")
        self.geometry("650x500")
        self.resizable(False, False)

        # --- Widgets -------------------------------------------------------
        tk.Label(self, text="Syndrom Name:").grid(row=0, column=0, sticky="e", pady=5, padx=5)
        self.syndrom_entry = tk.Entry(self, width=50)
        self.syndrom_entry.grid(row=0, column=1, columnspan=3, sticky="w", padx=5)

        # Golden image
        tk.Label(self, text="Golden Image:").grid(row=2, column=0, sticky="e", pady=5, padx=5)
        self.golden_path = tk.StringVar()
        tk.Entry(self, textvariable=self.golden_path, width=40, state="readonly").grid(row=2, column=1, sticky="w", padx=5)
        tk.Button(self, text="Browse", command=self._browse_golden).grid(row=2, column=2, padx=5, sticky="w")

        # Defect image
        tk.Label(self, text="Defect Image:").grid(row=3, column=0, sticky="e", pady=5, padx=5)
        self.defect_path = tk.StringVar()
        tk.Entry(self, textvariable=self.defect_path, width=40, state="readonly").grid(row=3, column=1, sticky="w", padx=5)
        tk.Button(self, text="Browse", command=self._browse_defect).grid(row=3, column=2, padx=5, sticky="w")

        # Description
        tk.Label(self, text="Description:").grid(row=4, column=0, sticky="ne", pady=5, padx=5)
        self.desc_text = scrolledtext.ScrolledText(self, width=48, height=10)
        self.desc_text.grid(row=4, column=1, columnspan=3, sticky="w", padx=5)

        # Action buttons
        tk.Button(self, text="Save", width=12, command=self._save).grid(row=5, column=1, pady=15, sticky="e")
        tk.Button(self, text="Clear", width=12, command=self._clear).grid(row=5, column=2, pady=15, sticky="w")

    # ------------------------------------------------------------------
    def _browse_golden(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Golden Image",
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.bmp")],
        )
        if path:
            self.golden_path.set(path)

    def _browse_defect(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Defect Image",
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.bmp")],
        )
        if path:
            self.defect_path.set(path)

    def _save(self) -> None:
        syndrom = self.syndrom_entry.get().strip()
        golden = self.golden_path.get()
        defect = self.defect_path.get()
        description = self.desc_text.get("1.0", tk.END).strip()

        if not syndrom:
            messagebox.showerror("Input Error", "Syndrom name is required.")
            return

        # Prepare target directory
        folder_name = sanitize_name(syndrom)
        target_dir = os.path.join(SYNDROM_DB_PATH, folder_name)

        # Check for duplicate Syndrom or Monitor name ----------------------
        # 1) Syndrom folder already exists
        if os.path.isdir(target_dir):
            messagebox.showwarning(
                "Duplicate Syndrom",
                f"A record for the syndrom '{syndrom}' already exists.\n"
                "Please choose a different name or delete the existing entry first.",
            )
            return

        # All good — create directory and copy files -----------------------
        os.makedirs(target_dir, exist_ok=False)  # Should succeed because we checked above

        try:
            if golden:
                shutil.copy(golden, os.path.join(target_dir, "golden.jpg"))
            if defect:
                shutil.copy(defect, os.path.join(target_dir, "defect.jpg"))

            # Write description.txt
            with open(os.path.join(target_dir, "description.txt"), "w", encoding="utf-8") as fh:
                fh.write(f"{description}\n")
        except Exception as exc:
            messagebox.showerror("Save Error", f"Failed to save files:\n{exc}")
            return

        messagebox.showinfo("Success", f"Syndrom '{syndrom}' saved to {target_dir}")
        self._clear()

    def _clear(self) -> None:
        self.syndrom_entry.delete(0, tk.END)
        self.golden_path.set("")
        self.defect_path.set("")
        self.desc_text.delete("1.0", tk.END)

# ----------------------------------------------------------------------
if __name__ == "__main__":
    # Ensure the SyndromDB directory exists
    if not os.path.isdir(SYNDROM_DB_PATH):
        os.makedirs(SYNDROM_DB_PATH, exist_ok=True)

    app = SyndromDBUI()
    app.mainloop() 