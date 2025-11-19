# python
import os
import json
import sqlite3
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

# Application directories and config
APP_DIR = Path(os.getenv("APPDATA") or Path.home()) / "openmsx_frontend"
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APP_DIR / "app_config.json"
DEFAULT_DB = str(APP_DIR / "app_data.db")


def ensure_config_file():
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps({"db_path": DEFAULT_DB}, indent=2), encoding="utf-8")


def load_config() -> dict:
    ensure_config_file()
    with CONFIG_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


class DBManager:
    """Simple SQLite config storage. Uses `INSERT OR REPLACE` for compatibility."""
    def __init__(self, db_path: str):
        self.db_path = str(Path(db_path).resolve())
        try:
            self.conn = sqlite3.connect(self.db_path, timeout=5)
            self._ensure_table()
        except Exception as e:
            messagebox.showerror("DB Error", f"Failed opening DB `{self.db_path}`:\n{e}")
            raise

    def _ensure_table(self):
        cur = self.conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        self.conn.commit()

    def get(self, key: str) -> Optional[str]:
        cur = self.conn.cursor()
        cur.execute("SELECT value FROM config WHERE key = ?", (key,))
        row = cur.fetchone()
        return row[0] if row else None

    def set(self, key: str, value: str):
        cur = self.conn.cursor()
        cur.execute("INSERT OR REPLACE INTO config(key, value) VALUES(?, ?)", (key, value))
        self.conn.commit()

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass


class OpenMSXFrontend:
    def __init__(self):
        if sys.platform != "win32":
            messagebox.showerror("Unsupported platform", "This program runs only on Windows.")
            sys.exit(1)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        cfg = load_config()
        db_path = cfg.get("db_path", DEFAULT_DB)
        self.db = DBManager(db_path)

        self.root = ctk.CTk()
        self.root.title("openMSX Frontend")
        self.root.geometry("760x420")
        self.root.minsize(640, 320)

        self.status_var = ctk.StringVar()
        self.pid_var = ctk.StringVar(value="PID: Not started")
        self.socket_var = ctk.StringVar(value="Socket:\n-")
        self.machine_var = ctk.StringVar(value="")
        self.current_socket_path: Optional[str] = None

        self._machines_cache: List[str] = []
        self._extensions_cache: List[str] = []
        self.listbox_extensions: Optional[tk.Listbox] = None

        self._build_main_ui()
        self._load_machines()
        self._load_extensions()
        self._update_status()

        if not self.db.get("openmsx_dir"):
            self.open_config_window(initial=True)

    def _build_main_ui(self):
        frame = ctk.CTkFrame(self.root, corner_radius=8)
        frame.pack(padx=16, pady=16, fill="both", expand=True)

        header = ctk.CTkLabel(frame, text="openMSX Frontend", font=ctk.CTkFont(size=20, weight="bold"))
        header.pack(pady=(6, 12))

        controls = ctk.CTkFrame(frame, corner_radius=0)
        controls.pack(fill="x", padx=8)

        btn_start = ctk.CTkButton(controls, text="Start openMSX", command=self.start_openmsx)
        btn_start.pack(side="left", padx=(0, 8))

        btn_config = ctk.CTkButton(controls, text="Configuration", command=lambda: self.open_config_window(initial=False))
        btn_config.pack(side="left", padx=(0, 8))

        # PID / Socket display
        info = ctk.CTkFrame(frame, corner_radius=0)
        info.pack(fill="x", pady=(12, 0), padx=8)

        pid_label = ctk.CTkLabel(info, textvariable=self.pid_var, anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        pid_label.pack(fill="x", anchor="w")

        # Socket on the next line for readability; allow wrapping
        socket_label = ctk.CTkLabel(info, textvariable=self.socket_var, anchor="w", font=ctk.CTkFont(size=10), wraplength=720, justify="left")
        socket_label.pack(fill="x", pady=(4, 0), anchor="w")

        # Machine combobox
        machine_row = ctk.CTkFrame(frame, corner_radius=0)
        machine_row.pack(fill="x", pady=(12, 0), padx=8)

        machine_label = ctk.CTkLabel(machine_row, text="MSX Model:", anchor="w", font=ctk.CTkFont(size=12))
        machine_label.pack(side="left")

        self.combo_machines = ctk.CTkComboBox(machine_row, values=[], variable=self.machine_var, width=420, command=self._on_machine_selected)
        self.combo_machines.pack(side="left", padx=(8, 0))
        self.combo_machines.configure(values=[], state="disabled")

        # Inserir área de extensões abaixo do machine_row
        ext_row = ctk.CTkFrame(frame, corner_radius=0)
        ext_row.pack(fill="both", pady=(12, 0), padx=8, expand=False)

        ext_label = ctk.CTkLabel(ext_row, text="Extensions:", anchor="w", font=ctk.CTkFont(size=12))
        ext_label.pack(anchor="w")

        # Frame para listbox + scrollbar (usar widgets tk dentro do CTkFrame)
        list_frame = ctk.CTkFrame(ext_row, corner_radius=0)
        list_frame.pack(fill="both", pady=(6, 0), expand=True)

        # listbox tkinter com seleção múltipla e altura 8
        lb = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, height=8, exportselection=False)
        lb.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.config(yscrollcommand=sb.set)

        # armazenar referência e bind para salvar seleção quando mudar
        self.listbox_extensions = lb
        lb.bind("<<ListboxSelect>>", self._on_extensions_selected)

        # Buttons area
        bottom = ctk.CTkFrame(frame, corner_radius=0)
        bottom.pack(fill="x", pady=(12, 0), padx=8)

        self.btn_socket = ctk.CTkButton(bottom, text="Open Socket", command=self._check_socket, state="disabled", fg_color="gray")
        self.btn_socket.pack(side="left", padx=(0, 8))

        status_label = ctk.CTkLabel(bottom, textvariable=self.status_var)
        status_label.pack(side="left", padx=(8, 0))

    def _machines_dir(self, openmsx_dir: str) -> Path:
        return Path(openmsx_dir) / "share" / "machines"

    def _extensions_dir(self, openmsx_dir: str) -> Path:
        return Path(openmsx_dir) / "share" / "extensions"

    def _get_machines(self) -> List[str]:
        """Return sorted list of machine names (file stems) from share/machines."""
        openmsx_dir = self.db.get("openmsx_dir") or ""
        machines: List[str] = []
        if openmsx_dir:
            p = self._machines_dir(openmsx_dir)
            if p.exists() and p.is_dir():
                for f in sorted(p.glob("*.xml")):
                    machines.append(f.stem)
        return machines

    def _get_extensions(self) -> List[str]:
        """Retorna nomes (sem .xml) das extensões em share/extensions."""
        openmsx_dir = self.db.get("openmsx_dir") or ""
        exts: List[str] = []
        if openmsx_dir:
            p = self._extensions_dir(openmsx_dir)
            if p.exists() and p.is_dir():
                for f in sorted(p.iterdir()):
                    try:
                        if f.is_file() and f.suffix.lower() == ".xml":
                            exts.append(f.stem)
                    except Exception:
                        continue
        return exts

    def _load_machines(self):
        machines = self._get_machines()
        self._machines_cache = machines
        if machines:
            self.combo_machines.configure(values=machines, state="normal")
            # restore saved selection if valid
            saved = self.db.get("openmsx_machine")
            if saved and saved in machines:
                self.machine_var.set(saved)
            else:
                # default to first
                self.machine_var.set(machines[0])
                try:
                    self.db.set("openmsx_machine", machines[0])
                except Exception:
                    pass
        else:
            self.combo_machines.configure(values=[], state="disabled")
            self.machine_var.set("")
            try:
                self.db.set("openmsx_machine", "")
            except Exception:
                pass

    def _load_extensions(self):
        """Preenche a listbox de extensões e restaura seleção do DB."""
        exts = self._get_extensions()
        self._extensions_cache = exts

        if self.listbox_extensions:
            lb = self.listbox_extensions
            lb.delete(0, tk.END)
            for item in exts:
                lb.insert(tk.END, item)

            if exts:
                # restaurar seleção salva (JSON)
                saved = self.db.get("openmsx_extensions")
                try:
                    if saved:
                        sel = json.loads(saved)
                        if isinstance(sel, list):
                            lb.selection_clear(0, tk.END)
                            for i, name in enumerate(exts):
                                if name in sel:
                                    lb.selection_set(i)
                except Exception:
                    pass
                lb.configure(state="normal")
            else:
                lb.configure(state="disabled")

    def _on_machine_selected(self, value: str):
        """Called when user chooses a machine in the combobox."""
        try:
            if value:
                self.db.set("openmsx_machine", value)
            else:
                self.db.set("openmsx_machine", "")
        except Exception:
            pass

    def _get_selected_extensions(self) -> List[str]:
        """Retorna lista de nomes selecionados na listbox."""
        if not self.listbox_extensions:
            return []
        sel_idxs = self.listbox_extensions.curselection()
        return [self._extensions_cache[i] for i in sel_idxs if 0 <= i < len(self._extensions_cache)]

    def _on_extensions_selected(self, event=None):
        """Salva seleção atual das extensões no DB."""
        selected = self._get_selected_extensions()
        try:
            self.db.set("openmsx_extensions", json.dumps(selected))
        except Exception:
            pass

    def start_openmsx(self):
        dir_path = self.db.get("openmsx_dir")
        if not dir_path:
            messagebox.showerror("Error", "openMSX directory not configured.")
            return

        exe_path = str(Path(dir_path) / "openmsx.exe")
        if not Path(exe_path).is_file():
            messagebox.showerror("Error", f"`openmsx.exe` not found in:\n{exe_path}")
            return

        machine = (self.machine_var.get() or self.db.get("openmsx_machine") or "").strip()
        if not machine:
            messagebox.showwarning("Machine not selected",
                                   "Please select an MSX model from the combobox before starting.")
            return

        try:
            # Build command with -machine option and -ext for cada extensão selecionada
            cmd = [exe_path, "-machine", machine]

            # anexar extensões selecionadas
            exts = self._get_selected_extensions()
            for e in exts:
                cmd.extend(["-ext", e])

            proc = subprocess.Popen(cmd, cwd=dir_path)
            initial_pid = proc.pid
            real_pid = initial_pid

            # Try to refine using psutil with a short retry loop to find the actual openmsx.exe process
            try:
                import psutil
                deadline = time.time() + 2.0  # up to 2 seconds
                while time.time() < deadline:
                    matches = []
                    for p in psutil.process_iter(['pid', 'exe', 'ppid']):
                        try:
                            exe = p.info.get('exe')
                            if exe and os.path.normcase(os.path.normpath(exe)) == os.path.normcase(os.path.normpath(exe_path)):
                                matches.append(p.pid)
                        except Exception:
                            continue
                    if matches:
                        real_pid = matches[-1]
                        break
                    # fallback: check children of initial process
                    try:
                        parent = psutil.Process(initial_pid)
                        children = parent.children(recursive=True)
                        if children:
                            # choose last child
                            real_pid = children[-1].pid
                            break
                    except Exception:
                        pass
                    time.sleep(0.15)
            except ImportError:
                messagebox.showwarning("psutil not installed",
                                       "Install `psutil` to improve PID detection (pip install psutil). Using initial PID.")
            except Exception:
                pass

            # Persist status
            try:
                self.db.set("openmsx_pid", str(real_pid))
            except Exception:
                pass
            self.pid_var.set(f"PID: {real_pid}")

            # Update status string including extensions
            ext_part = " ".join([f"-ext {e}" for e in exts]) if exts else ""
            self.status_var.set(f"openMSX started: {exe_path} -machine {machine} {ext_part}".strip())

            # Socket path & UI update
            self._update_socket_button(real_pid)
        except Exception as e:
            messagebox.showerror("Start Error", f"Failed to start openMSX:\n{e}")

    def _update_socket_button(self, pid: int):
        temp_dir = os.getenv("TEMP") or os.getcwd()
        socket_path = os.path.join(temp_dir, "openmsx-default", f"socket.{pid}")
        self.current_socket_path = socket_path
        # Show socket on next line for readability
        self.socket_var.set(f"Socket:\n{socket_path}")

        if hasattr(self, "btn_socket"):
            try:
                exists = os.path.exists(socket_path)
                # enable button even if not present, color indicates presence
                self.btn_socket.configure(state="normal", fg_color="green" if exists else "red")
            except Exception:
                # fallback: enable but gray
                try:
                    self.btn_socket.configure(state="normal", fg_color="gray")
                except Exception:
                    pass

    def _check_socket(self):
        path = self.current_socket_path
        if not path:
            messagebox.showwarning("Socket", "No socket path available.")
            return
        if os.path.exists(path):
            try:
                # Use explorer to select the file
                subprocess.Popen(['explorer', f'/select,{path}'])
            except Exception:
                # fallback: open folder
                try:
                    os.startfile(os.path.dirname(path))
                except Exception as e:
                    messagebox.showerror("Open Error", f"Can't open explorer:\n{e}")
        else:
            messagebox.showwarning("Socket not found", f"Socket doesn't exist at:\n{path}")

    def open_config_window(self, initial: bool = False):
        win = ctk.CTkToplevel(self.root)
        win.title("Configuration")
        win.geometry("640x160")
        win.grab_set()

        cur_dir = self.db.get("openmsx_dir") or ""
        entry_var = ctk.StringVar(value=cur_dir)

        lbl = ctk.CTkLabel(win, text="openMSX directory (folder containing openmsx.exe):")
        lbl.pack(padx=12, pady=(12, 6), anchor="w")

        entry_row = ctk.CTkFrame(win, corner_radius=0)
        entry_row.pack(fill="x", padx=12, pady=(0, 6))

        entry = ctk.CTkEntry(entry_row, textvariable=entry_var)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        def choose_dir():
            d = filedialog.askdirectory(title="Select openMSX directory")
            if d:
                entry_var.set(d)

        btn_choose = ctk.CTkButton(entry_row, text="Browse", width=120, command=choose_dir)
        btn_choose.pack(side="right")

        btn_frame = ctk.CTkFrame(win, corner_radius=0)
        btn_frame.pack(fill="x", padx=12, pady=(8, 12))

        def reset():
            entry_var.set("")

        def save():
            path = entry_var.get().strip()
            if not path:
                messagebox.showwarning("Validation", "Directory cannot be empty.")
                return
            exe_path = os.path.join(path, "openmsx.exe")
            if not os.path.isfile(exe_path):
                messagebox.showerror("Not found", f"`openmsx.exe` not found at:\n{exe_path}")
                return
            try:
                self.db.set("openmsx_dir", path)
                self._load_machines()
                self._load_extensions()
                self._update_status()
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed saving configuration:\n{e}")
            finally:
                win.destroy()

        def cancel():
            win.destroy()
            if initial and not self.db.get("openmsx_dir"):
                messagebox.showinfo("Configuration required", "Configuration not completed. You can configure later using the 'Configuration' button.")

        btn_reset = ctk.CTkButton(btn_frame, text="Reset", command=reset)
        btn_reset.pack(side="left", padx=6, pady=6)

        spacer = ctk.CTkLabel(btn_frame, text="")
        spacer.pack(side="left", expand=True)

        btn_cancel = ctk.CTkButton(btn_frame, text="Cancel", command=cancel)
        btn_cancel.pack(side="right", padx=6, pady=6)

        btn_save = ctk.CTkButton(btn_frame, text="Save", command=save)
        btn_save.pack(side="right", padx=6, pady=6)

    def _update_status(self):
        openmsx_dir = self.db.get("openmsx_dir")
        if openmsx_dir:
            self.status_var.set(f"openMSX: {openmsx_dir}")
        else:
            self.status_var.set("openMSX not configured")

        pid = self.db.get("openmsx_pid")
        if pid:
            try:
                pid_int = int(pid)
                self.pid_var.set(f"PID: {pid_int}")
                self._update_socket_button(pid_int)
            except (ValueError, TypeError):
                self.pid_var.set("PID: invalid")
                self.socket_var.set("Socket:\n-")
                self.current_socket_path = None
                if hasattr(self, "btn_socket"):
                    try:
                        self.btn_socket.configure(state="disabled", fg_color="gray")
                    except Exception:
                        pass
        else:
            self.pid_var.set("PID: Not started")
            self.socket_var.set("Socket:\n-")
            self.current_socket_path = None
            if hasattr(self, "btn_socket"):
                try:
                    self.btn_socket.configure(state="disabled", fg_color="gray")
                except Exception:
                    pass

        # Restore selected machine into combobox if present
        saved = self.db.get("openmsx_machine")
        if saved and saved in self._machines_cache:
            self.machine_var.set(saved)

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        try:
            self.db.close()
        finally:
            self.root.destroy()


if __name__ == "__main__":
    app = OpenMSXFrontend()
    app.run()
