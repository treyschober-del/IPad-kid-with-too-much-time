import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil

DARK_BG = "#121212"
DARK_FG = "#f0f0f0"
ACCENT = "#3f51b5"

class ItemCreator(tk.Toplevel):
    def __init__(self, parent=None):
        if parent is None:
            parent = tk.Tk()
            parent.withdraw()
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Standalone Item Creator")
        self.configure(bg=DARK_BG)

        tk.Label(
            self,
            text="Create New Item",
            bg=DARK_BG,
            fg=DARK_FG,
            font=("Papyrus", 20, "bold")
        ).grid(row=0, column=0, columnspan=3, pady=10)

        row_num = 1

        # Item name
        tk.Label(self, text="Item Name:", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, sticky="e", padx=5, pady=2)
        self.name_var = tk.StringVar()
        tk.Entry(self, textvariable=self.name_var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG).grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1

        # Basic stats (base stat modifiers)
        tk.Label(self, text="Base Stat Modifiers", bg=DARK_BG, fg=DARK_FG, font=("Papyrus", 12, "bold")).grid(row=row_num, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 5))
        row_num += 1

        self.stat_vars = {}
        for stat in ["Body", "Mind", "Soul", "Unconscious"]:
            tk.Label(self, text=f"{stat} Δ:", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, sticky="e", padx=5, pady=2)
            var = tk.StringVar(value="0")
            tk.Entry(self, textvariable=var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG, width=8).grid(row=row_num, column=1, sticky="w", padx=5, pady=2)
            self.stat_vars[stat.lower()] = var
            row_num += 1

        # Health stat modifiers
        tk.Label(self, text="Health Stat Modifiers", bg=DARK_BG, fg=DARK_FG, font=("Papyrus", 12, "bold")).grid(row=row_num, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 5))
        row_num += 1

        self.health_stat_vars = {}
        for stat in ["Body Health", "Mind Health", "Soul Health"]:
            tk.Label(self, text=f"{stat} Δ:", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, sticky="e", padx=5, pady=2)
            var = tk.StringVar(value="0")
            tk.Entry(self, textvariable=var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG, width=8).grid(row=row_num, column=1, sticky="w", padx=5, pady=2)
            self.health_stat_vars[stat.lower()] = var
            row_num += 1

        # Item size
        tk.Label(self, text="Item Size:", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, sticky="e", padx=5, pady=(10, 2))
        self.size_var = tk.StringVar(value="1")
        size_frame = tk.Frame(self, bg=DARK_BG)
        size_frame.grid(row=row_num, column=1, sticky="w", padx=5, pady=(10, 2))
        for size in ["1", "2", "4"]:
            tk.Radiobutton(
                size_frame,
                text=f"{size} slot(s)",
                variable=self.size_var,
                value=size,
                bg=DARK_BG,
                fg=DARK_FG,
                selectcolor=ACCENT,
                activebackground="#303030",
                highlightthickness=1,
                highlightbackground=DARK_BG,
                highlightcolor=ACCENT,
                font=("Papyrus", 10, "bold"),
                width=10,
            ).pack(side="left", padx=5)
        size_label = tk.Label(
            self,
            textvariable=self.size_var,
            bg=DARK_BG,
            fg=ACCENT,
            font=("Papyrus", 10, "bold"),
            width=4,
            anchor="w",
        )
        size_label.grid(row=row_num, column=2, sticky="w", padx=5)
        row_num += 1

        self.orientation_var = tk.StringVar(value="vertical")
        
        # Weapon option
        tk.Label(self, text="Is Weapon?", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, sticky="e", padx=5, pady=(10, 2))
        self.weapon_var = tk.StringVar(value="no")
        weapon_frame = tk.Frame(self, bg=DARK_BG)
        weapon_frame.grid(row=row_num, column=1, sticky="w", padx=5, pady=(10, 2))
        tk.Radiobutton(
            weapon_frame,
            text="Yes",
            variable=self.weapon_var,
            value="yes",
            bg=DARK_BG,
            fg=DARK_FG,
            selectcolor=ACCENT,
            activebackground="#303030",
            highlightthickness=1,
            highlightbackground=DARK_BG,
            highlightcolor=ACCENT,
            font=("Papyrus", 10, "bold"),
            width=6,
        ).pack(side="left")
        tk.Radiobutton(
            weapon_frame,
            text="No",
            variable=self.weapon_var,
            value="no",
            bg=DARK_BG,
            fg=DARK_FG,
            selectcolor=ACCENT,
            activebackground="#303030",
            highlightthickness=1,
            highlightbackground=DARK_BG,
            highlightcolor=ACCENT,
            font=("Papyrus", 10, "bold"),
            width=6,
        ).pack(side="left", padx=10)
        weapon_label = tk.Label(
            self,
            textvariable=self.weapon_var,
            bg=DARK_BG,
            fg=ACCENT,
            font=("Papyrus", 10, "bold"),
            width=4,
            anchor="w",
        )
        weapon_label.grid(row=row_num, column=2, sticky="w", padx=5)
        row_num += 1

        # Extra headers
        tk.Label(self, text="Extra headers (key:value per line):", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 2))
        row_num += 1

        self.extras_text = tk.Text(self, height=6, width=40, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG)
        self.extras_text.grid(row=row_num, column=0, columnspan=2, padx=5, pady=2)
        row_num += 1

        # PNG selection with drag-and-drop
        tk.Label(self, text="PNG Image:", bg=DARK_BG, fg=DARK_FG).grid(row=row_num, column=0, sticky="e", padx=5, pady=(10, 2))
        self.png_path_var = tk.StringVar()
        self.png_entry = tk.Entry(self, textvariable=self.png_path_var, bg="#1e1e1e", fg=DARK_FG, insertbackground=DARK_FG, width=30)
        self.png_entry.grid(row=row_num, column=1, sticky="w", padx=5, pady=(10, 2))
        
        # Try to enable drag-and-drop
        try:
            from tkinterdnd2 import DND_FILES, Earth
            self.png_entry.drop_target_register(DND_FILES)
            self.png_entry.dnd_bind('<<Drop>>', self._drop_png)
        except ImportError:
            pass

        tk.Button(self, text="Browse...", command=self._browse_png, bg="#1e1e1e", fg=DARK_FG).grid(row=row_num, column=2, padx=5, pady=(10, 2))
        row_num += 1

        # Save button
        tk.Button(
            self,
            text="Save Item",
            command=self._save_item,
            bg="#1e1e1e",
            fg=DARK_FG
        ).grid(row=row_num, column=0, columnspan=3, pady=15)

    def _browse_png(self):
        path = filedialog.askopenfilename(filetypes=[("PNG files", "*.png")])
        if path:
            self.png_path_var.set(path)

    def _drop_png(self, event):
        """Handle drag-and-drop PNG files"""
        files = self.parse_dnd_files(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.png'):
                self.png_path_var.set(file_path)
            else:
                messagebox.showwarning("Invalid File", "Please drop a PNG file.")

    def parse_dnd_files(self, data):
        """Parse drag-and-drop file data"""
        # Handle both Windows and Unix-style paths
        if data.startswith('{'):
            # Tcl list format
            files = self.tk.splitlist(data)
        else:
            # Space-separated paths (may have issues with spaces in filenames)
            files = data.split()
        return files

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
        
        # Add base stat modifiers
        for key, var in self.stat_vars.items():
            try:
                data[key] = int(var.get().strip() or "0")
            except ValueError:
                messagebox.showerror("Error", f"Invalid value for {key}.")
                return
        
        # Add health stat modifiers
        for key, var in self.health_stat_vars.items():
            try:
                value = int(var.get().strip() or "0")
                if value != 0:  # Only add if non-zero
                    data[key] = value
            except ValueError:
                messagebox.showerror("Error", f"Invalid value for {key}.")
                return
        
        # Add weapon option
        weapon = self.weapon_var.get().strip()
        if weapon:
            data["weapon"] = weapon

        # Add size
        size = self.size_var.get().strip()
        data["size"] = size
        
        # If size is 4 (2x2), prompt for orientation
        if size == "4":
            orientation_window = tk.Toplevel(self)
            orientation_window.title("Choose Orientation")
            orientation_window.configure(bg=DARK_BG)
            orientation_window.grab_set()
            
            tk.Label(
                orientation_window,
                text="2x2 items need orientation.\nChoose vertical or horizontal:",
                bg=DARK_BG,
                fg=DARK_FG
            ).pack(padx=20, pady=10)
            
            orientation = tk.StringVar(value="vertical")
            
            tk.Radiobutton(
                orientation_window,
                text="Vertical (2 wide, 2 tall)",
                variable=orientation,
                value="vertical",
                bg=DARK_BG,
                fg=DARK_FG
            ).pack(padx=20, pady=5)
            
            tk.Radiobutton(
                orientation_window,
                text="Horizontal (1 wide, 4 tall)",
                variable=orientation,
                value="horizontal",
                bg=DARK_BG,
                fg=DARK_FG
            ).pack(padx=20, pady=5)
            
            def confirm():
                data["orientation"] = orientation.get()
                orientation_window.destroy()
                self._write_item_data(name, data, item_folder)
            
            tk.Button(
                orientation_window,
                text="OK",
                command=confirm,
                bg="#1e1e1e",
                fg=DARK_FG
            ).pack(pady=10)
        else:
            self._write_item_data(name, data, item_folder)

    def _write_item_data(self, name, data, item_folder):
        # Add extra fields
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
        # Clear the form
        self.name_var.set("")
        for var in self.stat_vars.values():
            var.set("0")
        for var in self.health_stat_vars.values():
            var.set("0")
        self.weapon_var.set("no")
        self.size_var.set("1")
        self.extras_text.delete("1.0", "end")
        self.png_path_var.set("")

        messagebox.showinfo("Item Saved", f"Item '{name}' saved to {item_folder}.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    ItemCreator(root)
    root.mainloop()
