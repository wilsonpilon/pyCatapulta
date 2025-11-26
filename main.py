# python
import os
import json
import sqlite3
import subprocess
import sys
import time
import glob
import socket
import threading
from pathlib import Path
from typing import List, Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

# Application directories and config
APP_DIR = Path(os.getenv("APPDATA") or Path.home()) / "openmsx_frontend"
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APP_DIR / "app_config.json"
DEFAULT_DB = str(APP_DIR / "app_data.db")
MAX_HISTORY = 20


def ensure_config_file():
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps({"db_path": DEFAULT_DB}, indent=2), encoding="utf-8")


def load_config() -> dict:
    ensure_config_file()
    with CONFIG_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


class DBManager:
    """Simple SQLite config storage."""
    def __init__(self, db_path: str):
        self.db_path = str(Path(db_path).resolve())
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self._ensure_table()

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


class OpenMSXClientWindow:
    """
    Simple integrated TCP client to connect to openMSX TCP port (if available).
    """
    def __init__(self, parent: ctk.CTk):
        self.parent = parent
        self.win = ctk.CTkToplevel(parent)
        self.win.title("openMSX TCP Client")
        self.win.geometry("900x520")
        self.win.minsize(700, 420)
        self.win.protocol("WM_DELETE_WINDOW", self.on_close)

        self.frame = ctk.CTkFrame(self.win, corner_radius=8)
        self.frame.pack(fill="both", expand=True, padx=16, pady=16)

        header = ctk.CTkLabel(self.frame, text="openMSX TCP Client", font=ctk.CTkFont(size=18, weight="bold"))
        header.pack(pady=(6, 10))

        self.input = ctk.CTkTextbox(master=self.frame, width=860, height=300)
        self.input.insert("0.0", "# Digite o(s) comando(s) para openMSX aqui\n")
        self.input.pack(padx=10, pady=(10, 8), fill="both", expand=False)

        btn_row = ctk.CTkFrame(master=self.frame, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=(0, 8))

        self.send_btn = ctk.CTkButton(master=btn_row, text="Enviar", width=140, command=self.on_send)
        self.send_btn.pack(side="left", padx=(0, 8))

        self.close_btn = ctk.CTkButton(master=btn_row, text="Fechar", width=120, fg_color="red", hover_color="#ff6666", command=self.on_close)
        self.close_btn.pack(side="left")

        self.status = ctk.CTkLabel(master=self.frame, text="Status: inicializando...", anchor="w")
        self.status.pack(fill="x", padx=10, pady=(6, 0))

        self.resp = ctk.CTkTextbox(master=self.frame, width=860, height=120)
        self.resp.insert("0.0", "Resposta do servidor:\n")
        self.resp.configure(state="disabled")
        self.resp.pack(padx=10, pady=(8, 10), fill="both", expand=True)

        self.sock = None
        self.sock_lock = threading.Lock()
        self.tcp_port = None
        self._stop = False

        threading.Thread(target=self._background_find_port, daemon=True).start()

    def set_status(self, txt: str):
        try:
            self.status.configure(text=f"Status: {txt}")
        except Exception:
            pass

    def append_response(self, txt: str):
        def _append():
            try:
                self.resp.configure(state="normal")
                self.resp.insert("end", txt + "\n")
                self.resp.see("end")
                self.resp.configure(state="disabled")
            except Exception:
                pass
        try:
            self.win.after(0, _append)
        except Exception:
            pass

    def _background_find_port(self):
        self.set_status("Procurando arquivo de porta em \\%LOCALAPPDATA\\%\\Temp\\openmsx-default ...")
        for _ in range(12):
            if self._stop:
                return
            port = find_port_from_temp()
            if port:
                self.tcp_port = port
                self.set_status(f"Porta encontrada: {port}")
                return
            time.sleep(1.0)
        self.set_status("Não encontrou arquivo de porta.")

    def _ensure_connected(self) -> bool:
        if self.tcp_port is None:
            self.set_status("Porta desconhecida.")
            return False
        with self.sock_lock:
            if self.sock:
                return True
            try:
                s = socket.create_connection(("127.0.0.1", int(self.tcp_port)), timeout=1.5)
                self.sock = s
                self.set_status(f"Conectado a {self.tcp_port}")
                return True
            except Exception as e:
                self.set_status(f"Falha conectar: {e}")
                self.sock = None
                return False

    def _recv_all_until_quiet(self, s: socket.socket, idle_timeout: float = 0.25, max_total: float = 3.0) -> str:
        end_time = time.time() + max_total
        parts = []
        s.setblocking(False)
        try:
            while time.time() < end_time:
                try:
                    chunk = s.recv(4096)
                    if chunk:
                        parts.append(chunk.decode(errors="ignore"))
                        # extend end_time a little to allow more data
                        end_time = max(end_time, time.time() + idle_timeout)
                    else:
                        time.sleep(0.05)
                except BlockingIOError:
                    time.sleep(0.05)
        finally:
            try:
                s.setblocking(True)
            except Exception:
                pass
        return "".join(parts)

    def send_command_thread(self, cmd: str):
        if not self._ensure_connected():
            return
        with self.sock_lock:
            s = self.sock
        if not s:
            return
        try:
            data = cmd.strip()
            if not data.endswith("\n"):
                data += "\n"
            s.sendall(data.encode("utf-8"))
            resp = self._recv_all_until_quiet(s)
            self.append_response(resp or "<no response>")
            self.set_status("Comando enviado.")
        except Exception as e:
            self.append_response(f"<error: {e}>")
            try:
                with self.sock_lock:
                    if self.sock:
                        self.sock.close()
                    self.sock = None
            except Exception:
                pass

    def on_send(self):
        text = self.input.get("0.0", "end").strip()
        if not text:
            self.set_status("Nada para enviar.")
            return
        self.set_status("Enviando...")
        threading.Thread(target=self.send_command_thread, args=(text,), daemon=True).start()

    def on_close(self):
        self._stop = True
        with self.sock_lock:
            try:
                if self.sock:
                    self.sock.close()
            except Exception:
                pass
            self.sock = None
        try:
            self.win.destroy()
        except Exception:
            pass


class OpenMSXFrontend:
    def __init__(self):
        if sys.platform != "win32":
            messagebox.showerror("Platform", "This frontend runs on Windows only.")
            sys.exit(1)

        # initial theme defaults (can be overridden by DB)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        cfg = load_config()
        db_path = cfg.get("db_path", DEFAULT_DB)
        self.db = DBManager(db_path)

        # apply saved UI theme and other settings early
        saved_theme = (self.db.get("ui_theme") or "Light")
        if saved_theme == "Green":
            ctk.set_appearance_mode("Light")
            try:
                ctk.set_default_color_theme("green")
            except Exception:
                ctk.set_default_color_theme("blue")
        elif saved_theme == "Dark":
            ctk.set_appearance_mode("Dark")
            ctk.set_default_color_theme("blue")
        else:
            ctk.set_appearance_mode("Light")
            ctk.set_default_color_theme("blue")

        self.file_hunter_url = self.db.get("file_hunter_url") or "https://download.file-hunter.com/"
        self.msx_default_dir = self.db.get("msx_default_dir") or r"C:\msx"

        self.root = ctk.CTk()
        self.root.title("openMSX Frontend")
        self.root.geometry("1000x640")
        self.root.minsize(800, 480)

        self.status_var = ctk.StringVar()
        self.pid_var = ctk.StringVar(value="PID: Not started")
        self.socket_var = ctk.StringVar(value="Socket:\n-")
        self.machine_var = ctk.StringVar(value="")
        self.current_socket_path: Optional[str] = None

        # Disk A/B mode and history
        self.disk_a_mode = self.db.get("disk_a_mode") or "image"  # "image" or "directory"
        try:
            self.disk_a_history = json.loads(self.db.get("disk_a_history") or "[]")
            if not isinstance(self.disk_a_history, list):
                self.disk_a_history = []
        except Exception:
            self.disk_a_history = []
        self.disk_a_var = tk.StringVar(value="")

        self.disk_b_mode = self.db.get("disk_b_mode") or "image"
        try:
            self.disk_b_history = json.loads(self.db.get("disk_b_history") or "[]")
            if not isinstance(self.disk_b_history, list):
                self.disk_b_history = []
        except Exception:
            self.disk_b_history = []
        self.disk_b_var = tk.StringVar(value="")

        # Cartridge A/B history
        try:
            self.cart_a_history = json.loads(self.db.get("cart_a_history") or "[]")
            if not isinstance(self.cart_a_history, list):
                self.cart_a_history = []
        except Exception:
            self.cart_a_history = []
        self.cart_a_var = tk.StringVar(value="")

        try:
            self.cart_b_history = json.loads(self.db.get("cart_b_history") or "[]")
            if not isinstance(self.cart_b_history, list):
                self.cart_b_history = []
        except Exception:
            self.cart_b_history = []
        self.cart_b_var = tk.StringVar(value="")

        self._machines_cache: List[str] = []
        self._extensions_cache: List[str] = []
        self.listbox_extensions: Optional[tk.Listbox] = None

        self._build_main_ui()
        self._load_machines()
        self._load_extensions()
        self._update_status()

        # Start background port watcher to update Client TCP button color
        try:
            threading.Thread(target=self._start_port_watcher, daemon=True).start()
        except Exception:
            pass

        if not self.db.get("openmsx_dir"):
            # prompt configuration at startup
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

        # Button to open integrated TCP client window (store on self)
        self.btn_client = ctk.CTkButton(controls, text="Client TCP", command=self.open_client_window, fg_color="gray")
        self.btn_client.pack(side="left", padx=(0, 8))

        btn_exit = ctk.CTkButton(controls, text="Exit", fg_color="red", hover_color="#cc6666", command=self._on_close)
        btn_exit.pack(side="right", padx=(0, 8))

        # PID / Socket display
        info = ctk.CTkFrame(frame, corner_radius=0)
        info.pack(fill="x", pady=(12, 0), padx=8)

        pid_label = ctk.CTkLabel(info, textvariable=self.pid_var, anchor="w", font=ctk.CTkFont(size=12, weight="bold"))
        pid_label.pack(fill="x", anchor="w")

        socket_label = ctk.CTkLabel(info, textvariable=self.socket_var, anchor="w", font=ctk.CTkFont(size=10), wraplength=960, justify="left")
        socket_label.pack(fill="x", pady=(4, 0), anchor="w")

        # Machine combobox
        machine_row = ctk.CTkFrame(frame, corner_radius=0)
        machine_row.pack(fill="x", pady=(12, 0), padx=8)

        machine_label = ctk.CTkLabel(machine_row, text="MSX Model:", anchor="w", font=ctk.CTkFont(size=12))
        machine_label.pack(side="left")

        self.combo_machines = ctk.CTkComboBox(machine_row, values=[], variable=self.machine_var, width=600, command=self._on_machine_selected)
        self.combo_machines.pack(side="left", padx=(8, 0))
        self.combo_machines.configure(values=[], state="disabled")

        # Disk A media row
        media_row = ctk.CTkFrame(frame, corner_radius=0)
        media_row.pack(fill="x", pady=(12, 0), padx=8)

        media_label = ctk.CTkLabel(media_row, text="Disk A:", anchor="w", font=ctk.CTkFont(size=12))
        media_label.pack(side="left")

        # Use ttk.Combobox (editable) for path history + manual entry
        self.disk_a_combobox = ttk.Combobox(media_row, textvariable=self.disk_a_var, values=self.disk_a_history, width=80)
        self.disk_a_combobox.pack(side="left", padx=(8, 6), fill="x", expand=True)
        self.disk_a_combobox.state(["!readonly"])  # make editable

        # Browse button (disk icon)
        self.btn_disk_browse = ctk.CTkButton(media_row, text="💾", width=40, command=self._browse_disk_a)
        self.btn_disk_browse.pack(side="left", padx=(6, 6))

        # Eject button (eject icon)
        self.btn_disk_eject = ctk.CTkButton(media_row, text="⏏️ Eject", width=100, fg_color="#ff6666", hover_color="#ff8888", command=self._eject_disk_a)
        self.btn_disk_eject.pack(side="left", padx=(6, 6))

        # Options button to open the small listbox (image/default or directory)
        self.btn_disk_options = ctk.CTkButton(media_row, text="Options ▾", width=110, command=self._open_disk_a_options)
        self.btn_disk_options.pack(side="left", padx=(6, 0))

        # Disk B media row (mirror of Disk A)
        media_row_b = ctk.CTkFrame(frame, corner_radius=0)
        media_row_b.pack(fill="x", pady=(6, 0), padx=8)

        media_label_b = ctk.CTkLabel(media_row_b, text="Disk B:", anchor="w", font=ctk.CTkFont(size=12))
        media_label_b.pack(side="left")

        self.disk_b_combobox = ttk.Combobox(media_row_b, textvariable=self.disk_b_var, values=self.disk_b_history, width=80)
        self.disk_b_combobox.pack(side="left", padx=(8, 6), fill="x", expand=True)
        self.disk_b_combobox.state(["!readonly"])

        self.btn_disk_b_browse = ctk.CTkButton(media_row_b, text="💾", width=40, command=self._browse_disk_b)
        self.btn_disk_b_browse.pack(side="left", padx=(6, 6))

        self.btn_disk_b_eject = ctk.CTkButton(media_row_b, text="⏏️ Eject", width=100, fg_color="#ff6666", hover_color="#ff8888", command=self._eject_disk_b)
        self.btn_disk_b_eject.pack(side="left", padx=(6, 6))

        self.btn_disk_b_options = ctk.CTkButton(media_row_b, text="Options ▾", width=110, command=self._open_disk_b_options)
        self.btn_disk_b_options.pack(side="left", padx=(6, 0))

        # Cartridge A row (similar to Disk rows)
        cart_row = ctk.CTkFrame(frame, corner_radius=0)
        cart_row.pack(fill="x", pady=(12, 0), padx=8)

        cart_label = ctk.CTkLabel(cart_row, text="Cart A:", anchor="w", font=ctk.CTkFont(size=12))
        cart_label.pack(side="left")

        self.cart_a_combobox = ttk.Combobox(cart_row, textvariable=self.cart_a_var, values=self.cart_a_history, width=80)
        self.cart_a_combobox.pack(side="left", padx=(8, 6), fill="x", expand=True)
        self.cart_a_combobox.state(["!readonly"])

        # Browse button (cart icon)
        self.btn_cart_a_browse = ctk.CTkButton(cart_row, text="🎮", width=40, command=self._browse_cart_a)
        self.btn_cart_a_browse.pack(side="left", padx=(6, 6))

        # Eject button
        self.btn_cart_a_eject = ctk.CTkButton(cart_row, text="⏏️ Eject", width=100, fg_color="#ff6666", hover_color="#ff8888", command=self._eject_cart_a)
        self.btn_cart_a_eject.pack(side="left", padx=(6, 6))

        # Cartridge B row
        cart_row_b = ctk.CTkFrame(frame, corner_radius=0)
        cart_row_b.pack(fill="x", pady=(6, 0), padx=8)

        cart_label_b = ctk.CTkLabel(cart_row_b, text="Cart B:", anchor="w", font=ctk.CTkFont(size=12))
        cart_label_b.pack(side="left")

        self.cart_b_combobox = ttk.Combobox(cart_row_b, textvariable=self.cart_b_var, values=self.cart_b_history, width=80)
        self.cart_b_combobox.pack(side="left", padx=(8, 6), fill="x", expand=True)
        self.cart_b_combobox.state(["!readonly"])

        self.btn_cart_b_browse = ctk.CTkButton(cart_row_b, text="🎮", width=40, command=self._browse_cart_b)
        self.btn_cart_b_browse.pack(side="left", padx=(6, 6))

        self.btn_cart_b_eject = ctk.CTkButton(cart_row_b, text="⏏️ Eject", width=100, fg_color="#ff6666", hover_color="#ff8888", command=self._eject_cart_b)
        self.btn_cart_b_eject.pack(side="left", padx=(6, 6))

        # Extensions area
        ext_row = ctk.CTkFrame(frame, corner_radius=0)
        ext_row.pack(fill="both", pady=(12, 0), padx=8, expand=False)

        ext_label = ctk.CTkLabel(ext_row, text="Extensions:", anchor="w", font=ctk.CTkFont(size=12))
        ext_label.pack(anchor="w")

        list_frame = ctk.CTkFrame(ext_row, corner_radius=0)
        list_frame.pack(fill="both", pady=(6, 0), expand=True)

        lb = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, height=12, exportselection=False)
        lb.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.config(yscrollcommand=sb.set)

        self.listbox_extensions = lb
        lb.bind("<<ListboxSelect>>", self._on_extensions_selected)

        bottom = ctk.CTkFrame(frame, corner_radius=0)
        bottom.pack(fill="x", pady=(12, 0), padx=8)

        self.btn_socket = ctk.CTkButton(bottom, text="Open Socket", command=self._check_socket, state="disabled", fg_color="gray")
        self.btn_socket.pack(side="left", padx=(0, 8))

        status_label = ctk.CTkLabel(bottom, textvariable=self.status_var)
        status_label.pack(side="left", padx=(8, 0))

        # initialize display of disk mode in options buttons
        self._update_disk_options_button_text()

    def _update_disk_options_button_text(self):
        try:
            self.btn_disk_options.configure(text=f"Mode: {self.disk_a_mode} ▾")
            self.btn_disk_b_options.configure(text=f"Mode: {self.disk_b_mode} ▾")
        except Exception:
            pass

    def _browse_disk_a(self):
        try:
            if self.disk_a_mode == "directory":
                sel = filedialog.askdirectory(title="Select Disk A directory")
                if sel:
                    self.disk_a_var.set(sel)
                    self._add_disk_a_history(sel)
                    self.db.set("disk_a_current", sel)
            else:
                sel = filedialog.askopenfilename(title="Select Disk A image", filetypes=[("Disk images", "*.dsk *.img *.zip *.vhd *.iso"), ("All files", "*.*")])
                if sel:
                    self.disk_a_var.set(sel)
                    self._add_disk_a_history(sel)
                    self.db.set("disk_a_current", sel)
        except Exception as e:
            messagebox.showerror("Browse Disk A", str(e))

    def _eject_disk_a(self):
        try:
            self.disk_a_var.set("")
            self.db.set("disk_a_current", "")
        except Exception:
            pass

    def _open_disk_a_options(self):
        try:
            win = ctk.CTkToplevel(self.root)
            win.title("Disk A Options")
            win.geometry("320x160")
            win.grab_set()

            var = tk.StringVar(value=self.disk_a_mode)
            r1 = ctk.CTkRadioButton(win, text="Image file", variable=var, value="image")
            r1.pack(anchor="w", padx=12, pady=8)
            r2 = ctk.CTkRadioButton(win, text="Directory", variable=var, value="directory")
            r2.pack(anchor="w", padx=12, pady=8)

            def do_save():
                self.disk_a_mode = var.get()
                self.db.set("disk_a_mode", self.disk_a_mode)
                self._update_disk_options_button_text()
                win.destroy()

            btn = ctk.CTkButton(win, text="Save", command=do_save)
            btn.pack(pady=8)
        except Exception as e:
            messagebox.showerror("Disk A Options", str(e))

    def _add_disk_a_history(self, path: str):
        try:
            if not path:
                return
            if path in self.disk_a_history:
                self.disk_a_history.remove(path)
            self.disk_a_history.insert(0, path)
            self.disk_a_history = self.disk_a_history[:MAX_HISTORY]
            self.db.set("disk_a_history", json.dumps(self.disk_a_history))
            self.disk_a_combobox['values'] = self.disk_a_history
        except Exception:
            pass

    def _browse_disk_b(self):
        try:
            if self.disk_b_mode == "directory":
                sel = filedialog.askdirectory(title="Select Disk B directory")
                if sel:
                    self.disk_b_var.set(sel)
                    self._add_disk_b_history(sel)
                    self.db.set("disk_b_current", sel)
            else:
                sel = filedialog.askopenfilename(title="Select Disk B image", filetypes=[("Disk images", "*.dsk *.img *.zip *.vhd *.iso"), ("All files", "*.*")])
                if sel:
                    self.disk_b_var.set(sel)
                    self._add_disk_b_history(sel)
                    self.db.set("disk_b_current", sel)
        except Exception as e:
            messagebox.showerror("Browse Disk B", str(e))

    def _eject_disk_b(self):
        try:
            self.disk_b_var.set("")
            self.db.set("disk_b_current", "")
        except Exception:
            pass

    def _open_disk_b_options(self):
        try:
            win = ctk.CTkToplevel(self.root)
            win.title("Disk B Options")
            win.geometry("320x160")
            win.grab_set()

            var = tk.StringVar(value=self.disk_b_mode)
            r1 = ctk.CTkRadioButton(win, text="Image file", variable=var, value="image")
            r1.pack(anchor="w", padx=12, pady=8)
            r2 = ctk.CTkRadioButton(win, text="Directory", variable=var, value="directory")
            r2.pack(anchor="w", padx=12, pady=8)

            def do_save():
                self.disk_b_mode = var.get()
                self.db.set("disk_b_mode", self.disk_b_mode)
                self._update_disk_options_button_text()
                win.destroy()

            btn = ctk.CTkButton(win, text="Save", command=do_save)
            btn.pack(pady=8)
        except Exception as e:
            messagebox.showerror("Disk B Options", str(e))

    def _add_disk_b_history(self, path: str):
        try:
            if not path:
                return
            if path in self.disk_b_history:
                self.disk_b_history.remove(path)
            self.disk_b_history.insert(0, path)
            self.disk_b_history = self.disk_b_history[:MAX_HISTORY]
            self.db.set("disk_b_history", json.dumps(self.disk_b_history))
            self.disk_b_combobox['values'] = self.disk_b_history
        except Exception:
            pass

    def _browse_cart_a(self):
        """Open file dialog to select a cartridge image for Cart A (.rom / .zip)."""
        try:
            sel = filedialog.askopenfilename(title="Select Cart A ROM", filetypes=[("ROMs and zips", "*.rom *.zip *.bin"), ("All files", "*.*")])
            if sel:
                self.cart_a_var.set(sel)
                self._add_cart_a_history(sel)
                self.db.set("cart_a_current", sel)
        except Exception as e:
            messagebox.showerror("Browse Cart A", str(e))

    def _eject_cart_a(self):
        """Eject Cart A: clear field and persist empty current."""
        try:
            self.cart_a_var.set("")
            self.db.set("cart_a_current", "")
        except Exception:
            pass

    def _add_cart_a_history(self, path: str):
        try:
            if not path:
                return
            if path in self.cart_a_history:
                self.cart_a_history.remove(path)
            self.cart_a_history.insert(0, path)
            self.cart_a_history = self.cart_a_history[:MAX_HISTORY]
            self.db.set("cart_a_history", json.dumps(self.cart_a_history))
            self.cart_a_combobox['values'] = self.cart_a_history
        except Exception:
            pass

    def _browse_cart_b(self):
        """Open file dialog to select a cartridge image for Cart B (.rom / .zip)."""
        try:
            sel = filedialog.askopenfilename(title="Select Cart B ROM", filetypes=[("ROMs and zips", "*.rom *.zip *.bin"), ("All files", "*.*")])
            if sel:
                self.cart_b_var.set(sel)
                self._add_cart_b_history(sel)
                self.db.set("cart_b_current", sel)
        except Exception as e:
            messagebox.showerror("Browse Cart B", str(e))

    def _eject_cart_b(self):
        """Eject Cart B: clear field and persist empty current."""
        try:
            self.cart_b_var.set("")
            self.db.set("cart_b_current", "")
        except Exception:
            pass

    def _add_cart_b_history(self, path: str):
        try:
            if not path:
                return
            if path in self.cart_b_history:
                self.cart_b_history.remove(path)
            self.cart_b_history.insert(0, path)
            self.cart_b_history = self.cart_b_history[:MAX_HISTORY]
            self.db.set("cart_b_history", json.dumps(self.cart_b_history))
            self.cart_b_combobox['values'] = self.cart_b_history
        except Exception:
            pass

    def _machines_dir(self, openmsx_dir: str) -> Path:
        return Path(openmsx_dir) / "share" / "machines"

    def _extensions_dir(self, openmsx_dir: str) -> Path:
        return Path(openmsx_dir) / "share" / "extensions"

    def _get_machines(self) -> List[str]:
        openmsx_dir = self.db.get("openmsx_dir") or ""
        machines: List[str] = []
        if openmsx_dir:
            md = self._machines_dir(openmsx_dir)
            if md.is_dir():
                for child in sorted(md.iterdir()):
                    if child.is_file():
                        machines.append(child.stem)
                    elif child.is_dir():
                        machines.append(child.name)
        return machines

    def _get_extensions(self) -> List[str]:
        openmsx_dir = self.db.get("openmsx_dir") or ""
        exts: List[str] = []
        if openmsx_dir:
            ed = self._extensions_dir(openmsx_dir)
            if ed.is_dir():
                for child in sorted(ed.iterdir()):
                    if child.is_dir() or child.suffix:
                        exts.append(child.name)
        return exts

    def _load_machines(self):
        machines = self._get_machines()
        self._machines_cache = machines
        if machines:
            try:
                self.combo_machines.configure(values=machines, state="normal")
                saved = self.db.get("openmsx_machine") or ""
                if saved and saved in machines:
                    self.machine_var.set(saved)
            except Exception:
                pass
        else:
            self.combo_machines.configure(values=[], state="disabled")

    def _load_extensions(self):
        exts = self._get_extensions()
        self._extensions_cache = exts
        if self.listbox_extensions:
            self.listbox_extensions.delete(0, "end")
            for e in exts:
                self.listbox_extensions.insert("end", e)
            # restore selection
            try:
                saved = json.loads(self.db.get("openmsx_extensions") or "[]")
                for i, name in enumerate(exts):
                    if name in saved:
                        self.listbox_extensions.selection_set(i)
            except Exception:
                pass

    def _on_machine_selected(self, value: str):
        try:
            if value:
                self.db.set("openmsx_machine", value)
        except Exception:
            pass

    def _get_selected_extensions(self) -> List[str]:
        if not self.listbox_extensions:
            return []
        sel_idxs = self.listbox_extensions.curselection()
        return [self._extensions_cache[i] for i in sel_idxs if 0 <= i < len(self._extensions_cache)]

    def _on_extensions_selected(self, event=None):
        selected = self._get_selected_extensions()
        try:
            self.db.set("openmsx_extensions", json.dumps(selected))
        except Exception:
            pass

    def start_openmsx(self):
        dir_path = self.db.get("openmsx_dir")
        if not dir_path:
            messagebox.showerror("Start", "openMSX directory not configured.")
            return

        exe_path = str(Path(dir_path) / "openmsx.exe")
        if not Path(exe_path).is_file():
            messagebox.showerror("Start", f"openmsx.exe not found in {exe_path}")
            return

        machine = (self.machine_var.get() or self.db.get("openmsx_machine") or "").strip()
        if not machine:
            messagebox.showerror("Start", "No MSX machine selected.")
            return

        # Build a basic command. Users can modify the frontend later to pass media/extension args.
        try:
            proc = subprocess.Popen([exe_path], cwd=dir_path)
            pid = proc.pid
            self.db.set("openmsx_pid", str(pid))
            self.pid_var.set(f"PID: {pid}")
            self._update_socket_button(pid)
            self.status_var.set("openMSX started.")
        except Exception as e:
            messagebox.showerror("Start", f"Failed to start openMSX:\n{e}")

    def _update_socket_button(self, pid: int):
        temp_dir = os.getenv("TEMP") or os.getcwd()
        socket_path = os.path.join(temp_dir, "openmsx-default", f"socket.{pid}")
        self.current_socket_path = socket_path
        self.socket_var.set(f"Socket:\n{socket_path}")

        if hasattr(self, "btn_socket"):
            if os.path.exists(socket_path):
                self.btn_socket.configure(fg_color="#66cc66")  # green
                self.btn_socket.configure(state="normal")
            else:
                self.btn_socket.configure(fg_color="gray")
                self.btn_socket.configure(state="normal")

    def _check_socket(self):
        path = self.current_socket_path
        if not path:
            messagebox.showinfo("Socket", "No socket path available.")
            return
        if os.path.exists(path):
            messagebox.showinfo("Socket", f"Socket exists:\n{path}")
        else:
            messagebox.showwarning("Socket", f"Socket not found:\n{path}")

    def open_config_window(self, initial: bool = False):
        win = ctk.CTkToplevel(self.root)
        win.title("Configuration")
        win.geometry("640x240")
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

        # Theme selection, File Hunter URL and MSX default directory
        theme_row = ctk.CTkFrame(win, corner_radius=0)
        theme_row.pack(fill="x", padx=12, pady=(4, 6))
        theme_lbl = ctk.CTkLabel(theme_row, text="UI Theme:")
        theme_lbl.pack(side="left", padx=(0, 8))
        theme_var = ctk.StringVar(value=self.db.get("ui_theme") or "Light")
        theme_cb = ctk.CTkComboBox(theme_row, values=["Light", "Dark", "Green"], variable=theme_var, width=220)
        theme_cb.pack(side="left")

        fh_row = ctk.CTkFrame(win, corner_radius=0)
        fh_row.pack(fill="x", padx=12, pady=(4, 6))
        fh_lbl = ctk.CTkLabel(fh_row, text="File Hunter URL:")
        fh_lbl.pack(side="left", padx=(0, 8))
        fh_var = ctk.StringVar(value=self.db.get("file_hunter_url") or "https://download.file-hunter.com/")
        fh_entry = ctk.CTkEntry(fh_row, textvariable=fh_var)
        fh_entry.pack(side="left", fill="x", expand=True)

        msx_row = ctk.CTkFrame(win, corner_radius=0)
        msx_row.pack(fill="x", padx=12, pady=(4, 6))
        msx_lbl = ctk.CTkLabel(msx_row, text="Default MSX directory:")
        msx_lbl.pack(side="left", padx=(0, 8))
        msx_var = ctk.StringVar(value=self.db.get("msx_default_dir") or r"C:\msx")
        msx_entry = ctk.CTkEntry(msx_row, textvariable=msx_var)
        msx_entry.pack(side="left", fill="x", expand=True)

        def choose_msx_dir():
            d = filedialog.askdirectory(title="Select default MSX directory")
            if d:
                msx_var.set(d)

        btn_msx_choose = ctk.CTkButton(msx_row, text="Browse", width=90, command=choose_msx_dir)
        btn_msx_choose.pack(side="right", padx=(6, 0))

        btn_frame = ctk.CTkFrame(win, corner_radius=0)
        btn_frame.pack(fill="x", padx=12, pady=(8, 12))

        def reset():
            entry_var.set("")
            theme_var.set("Light")
            fh_var.set("https://download.file-hunter.com/")
            msx_var.set(r"C:\msx")

        def save():
            path = entry_var.get().strip()
            if not path:
                messagebox.showerror("Configuration", "openMSX directory cannot be empty.")
                return
            exe_path = os.path.join(path, "openmsx.exe")
            if not os.path.isfile(exe_path):
                messagebox.showerror("Configuration", f"`openmsx.exe` not found in:\n{exe_path}")
                return
            try:
                self.db.set("openmsx_dir", path)
                self.db.set("ui_theme", theme_var.get())
                self.db.set("file_hunter_url", fh_var.get().strip())
                self.db.set("msx_default_dir", msx_var.get().strip())
                # apply theme immediately
                chosen = theme_var.get()
                if chosen == "Green":
                    ctk.set_appearance_mode("Light")
                    try:
                        ctk.set_default_color_theme("green")
                    except Exception:
                        ctk.set_default_color_theme("blue")
                elif chosen == "Dark":
                    ctk.set_appearance_mode("Dark")
                    ctk.set_default_color_theme("blue")
                else:
                    ctk.set_appearance_mode("Light")
                    ctk.set_default_color_theme("blue")
                self.file_hunter_url = fh_var.get().strip()
                self.msx_default_dir = msx_var.get().strip()
                win.destroy()
                # reload machines/extensions with new openmsx_dir
                self._load_machines()
                self._load_extensions()
            except Exception as e:
                messagebox.showerror("Configuration", f"Failed saving configuration:\n{e}")

        def cancel():
            win.destroy()
            if initial and not self.db.get("openmsx_dir"):
                try:
                    self.root.destroy()
                except Exception:
                    pass

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
            self.status_var.set(f"openMSX dir: {openmsx_dir}")
        else:
            self.status_var.set("openMSX dir: not set")

        pid = self.db.get("openmsx_pid")
        if pid:
            try:
                p = int(pid)
                self.pid_var.set(f"PID: {p}")
                self._update_socket_button(p)
            except Exception:
                self.pid_var.set("PID: invalid")
        else:
            self.pid_var.set("PID: Not started")

        saved = self.db.get("openmsx_machine")
        if saved and saved in self._machines_cache:
            self.machine_var.set(saved)

        last_disk = self.db.get("disk_a_current")
        if last_disk:
            self.disk_a_var.set(last_disk)

        last_disk_b = self.db.get("disk_b_current")
        if last_disk_b:
            self.disk_b_var.set(last_disk_b)

        last_cart = self.db.get("cart_a_current")
        if last_cart:
            self.cart_a_var.set(last_cart)

        last_cart_b = self.db.get("cart_b_current")
        if last_cart_b:
            self.cart_b_var.set(last_cart_b)

    def open_client_window(self):
        try:
            OpenMSXClientWindow(self.root)
        except Exception as e:
            messagebox.showerror("Client", str(e))

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        try:
            # persist histories and current selections
            try:
                self.db.set("disk_a_history", json.dumps(self.disk_a_history))
                self.db.set("disk_b_history", json.dumps(self.disk_b_history))
                self.db.set("cart_a_history", json.dumps(self.cart_a_history))
                self.db.set("cart_b_history", json.dumps(self.cart_b_history))
                # current selections
                self.db.set("disk_a_current", self.disk_a_var.get() or "")
                self.db.set("disk_b_current", self.disk_b_var.get() or "")
                self.db.set("cart_a_current", self.cart_a_var.get() or "")
                self.db.set("cart_b_current", self.cart_b_var.get() or "")
            except Exception:
                pass
        finally:
            try:
                self.db.close()
            except Exception:
                pass
            try:
                self.root.destroy()
            except Exception:
                pass

    def _check_local_port(self, port: int) -> bool:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except Exception:
            return False

    def _start_port_watcher(self, interval: float = 2.0):
        while True:
            try:
                port = find_port_from_temp()
                if port and self._check_local_port(port):
                    # update button color on main thread
                    def _set_green():
                        try:
                            self.btn_client.configure(fg_color="#66cc66")
                        except Exception:
                            pass
                    self.root.after(0, _set_green)
                else:
                    def _set_gray():
                        try:
                            self.btn_client.configure(fg_color="gray")
                        except Exception:
                            pass
                    self.root.after(0, _set_gray)
                # update socket path if pid stored
                pid = self.db.get("openmsx_pid")
                if pid:
                    try:
                        p = int(pid)
                        self._update_socket_button(p)
                    except Exception:
                        pass
            except Exception:
                pass
            time.sleep(interval)


def find_port_from_temp() -> Optional[int]:
    local_appdata = os.environ.get("LOCALAPPDATA")
    if not local_appdata:
        return None
    base = os.path.join(local_appdata, "Temp", "openmsx-default")
    if not os.path.isdir(base):
        return None
    files = [f for f in glob.glob(os.path.join(base, "*")) if os.path.isfile(f)]
    if not files:
        return None
    latest = max(files, key=os.path.getmtime)
    try:
        with open(latest, "r", encoding="utf-8", errors="ignore") as fh:
            content = fh.read().strip()
            # try parse as int or extract number
            try:
                return int(content)
            except Exception:
                # try find first number in text
                import re
                m = re.search(r"\d+", content)
                if m:
                    return int(m.group(0))
    except Exception:
        return None
    return None


if __name__ == "__main__":
    app = OpenMSXFrontend()
    app.run()