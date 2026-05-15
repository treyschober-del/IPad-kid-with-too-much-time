import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil

DARK_BG = "#121212"
DARK_FG = "#f0f0f0"

class ItemCreator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Standalone Item Creator")
        self.configure(bg=DARK_BG)

        tk.Label(
            self,
            text="Create New Item",
            bg=DARK_BG,
            fg=DARK_FG,
            font=("Papyrus", 18, "bold")
        ).grid(row=0, column=0, columnspan=3, pady=10)

        # Item name
        tk.Label(self, text="Item Name:", bg=DARK_BG, fg=DARK_FG).grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.name_var = tk.StringVar()
        tk.Entry(self, textvariable=self.name_var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG).grid(row=1, column=1, padx=5, pady=2)

        # Basic stats
        self.stat_vars = {}
        for i, stat in enumerate(["Body", "Mind", "Soul", "Unconscious"], start=2):
            tk.Label(self, text=f"{stat} Δ:", bg=DARK_BG, fg=DARK_FG).grid(row=i, column=0, sticky="e", padx=5, pady=2)
            var = tk.StringVar(value="0")
            tk.Entry(self, textvariable=var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG, width=8).grid(row=i, column=1, sticky="w", padx=5, pady=2)
            self.stat_vars[stat.lower()] = var

        # Extra headers
        tk.Label(self, text="Extra headers (key:value per line):", bg=DARK_BG, fg=DARK_FG).grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 2))
        self.extras_text = tk.Text(self, height=6, width=40, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG)
        self.extras_text.grid(row=7, column=0, columnspan=2, padx=5, pady=2)

        # PNG selection
        tk.Label(self, text="Optional PNG:", bg=DARK_BG, fg=DARK_FG).grid(row=8, column=0, sticky="e", padx=5, pady=(10, 2))
        self.png_path_var = tk.StringVar()
        tk.Entry(self, textvariable=self.png_path_var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG, width=30).grid(row=8, column=1, sticky="w", padx=5, pady=(10, 2))

        tk.Button(self, text="Browse...", command=self._browse_png, bg="#1e1e1e", fg=DARK_FG).grid(row=8, column=2, padx=5, pady=(10, 2))

        # Save button
        tk.Button(
            self,
            text="Save Item",
            command=self._save_item,
            bg="#1e1e1e",
            fg=DARK_FG
        ).grid(row=9, column=0, columnspan=3, pady=15)

    def _browse_png(self):
        path = filedialog.askopenfilename(filetypes=[("PNG files", "*.png")])
        if path:
            self.png_path_var.set(path)

    def _save_item(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Item name is required.")
            return

        # Create folder
        items_root = "items"
        os.makedirs(items_root, exist_ok=True)
        item_folder = os.path.join(items_root, name)
        os.makedirs(item_folder, exist_ok=True)

        # Build data dict
        data = {"name": name}
        for key, var in self.stat_vars.items():
            try:
                data[key] = int(var.get().strip() or "0")
            except ValueError:
                messagebox.showerror("Error", f"Invalid value for {key}.")
                return

        extra_lines = self.extras_text.get("1.0", "end-1c").splitlines()
        for line in extra_lines:
            if ":" in line:
                k, v = line.split(":", 1)
                data[k.strip()] = v.strip()

        # Write item.txt
        txt_path = os.path.join(item_folder, "item.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            for k, v in data.items():
                f.write(f"{k}\t{v}\n")

        # Copy PNG if provided
        png_src = self.png_path_var.get().strip()
        if png_src and os.path.exists(png_src):
            png_dst = os.path.join(item_folder, os.path.basename(png_src))
            shutil.copy2(png_src, png_dst)

        messagebox.showinfo("Item Saved", f"Item '{name}' saved to {item_folder}.")

if __name__ == "__main__":
    app = ItemCreator()
    app.mainloop()
