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


def find_port_from_temp() -> Optional[int]:
    """
    Locate the directory:
      %LOCALAPPDATA%\Temp\openmsx-default
    and read the most recently modified file. That file should contain the port number.
    Returns int(port) or None.
    """
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
            return int(content)
    except Exception:
        return None


class OpenMSXClientWindow:
    """
    Integrated TCP client window.
    - Sends commands terminated by '\\n'.
    - Reads responses by repeatedly receiving data until a short recv idle timeout occurs.
    - Keeps a single socket per window and closes it on errors/exit.
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

        # Start background task to find port
        threading.Thread(target=self._background_find_port, daemon=True).start()

    def set_status(self, txt: str):
        try:
            self.status.configure(text=f"Status: {txt}")
        except Exception:
            pass

    def append_response(self, txt: str):
        def _append():
            self.resp.configure(state="normal")
            self.resp.insert("end", txt + "\n")
            self.resp.see("end")
            self.resp.configure(state="disabled")
        try:
            self.win.after(0, _append)
        except Exception:
            pass

    def _background_find_port(self):
        self.set_status("Procurando arquivo de porta em \\%LOCALAPPDATA\\%\\Temp\\openmsx-default ...")
        for _ in range(12):
            if self._stop:
                return
            p = find_port_from_temp()
            if p:
                self.tcp_port = p
                self.set_status(f"Porta encontrada: {p}")
                return
            time.sleep(0.5)
        self.set_status("Não encontrou arquivo de porta.")

    def _ensure_connected(self) -> bool:
        """
        Ensure self.sock is connected. If missing, create new connection to found port.
        Returns True if ready to send.
        """
        if self.tcp_port is None:
            self.append_response("Porta não disponível.")
            return False

        with self.sock_lock:
            if self.sock:
                return True

        try:
            s = socket.create_connection(("127.0.0.1", self.tcp_port), timeout=3.0)
            # Use a short recv timeout for responsive reads
            s.settimeout(0.6)
            with self.sock_lock:
                self.sock = s
            self.set_status(f"Conectado 127.0.0.1:{self.tcp_port}")
            return True
        except Exception as e:
            self.append_response(f"Falha ao conectar TCP: {e}")
            self.set_status("Conexão falhou")
            return False

    def _recv_all_until_quiet(self, s: socket.socket, idle_timeout: float = 0.25, max_total: float = 3.0) -> str:
        """
        Read repeatedly until no data arrives for `idle_timeout` seconds or `max_total` reached.
        This tolerates multi-line replies from the server.
        """
        end_time = time.time() + max_total
        parts = []
        try:
            s.settimeout(idle_timeout)
            while time.time() < end_time:
                try:
                    chunk = s.recv(4096)
                    if not chunk:
                        break
                    parts.append(chunk)
                except socket.timeout:
                    break
                except Exception:
                    break
        finally:
            try:
                s.settimeout(0.6)
            except Exception:
                pass
        if parts:
            try:
                return b"".join(parts).decode("utf-8", errors="ignore").strip()
            except Exception:
                return "".join([p.decode("utf-8", errors="ignore") if isinstance(p, bytes) else str(p) for p in parts]).strip()
        return ""

    def send_command_thread(self, cmd: str):
        if not self._ensure_connected():
            return
        with self.sock_lock:
            s = self.sock
        if not s:
            self.append_response("Sem socket disponível.")
            return
        try:
            if not cmd.endswith("\n"):
                cmd += "\n"
            s.sendall(cmd.encode("utf-8"))
            reply = self._recv_all_until_quiet(s, idle_timeout=0.25, max_total=2.0)
            if reply:
                self.append_response(f"Resposta:\n{reply}")
            else:
                self.append_response("Comando enviado; sem resposta recebida (timeout).")
        except Exception as e:
            self.append_response(f"Erro ao enviar: {e}")
            with self.sock_lock:
                try:
                    if self.sock:
                        self.sock.close()
                except Exception:
                    pass
                self.sock = None
            self.set_status("Conexão perdida")

    def on_send(self):
        text = self.input.get("0.0", "end").strip()
        if not text:
            self.append_response("Nada para enviar.")
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
            messagebox.showerror("Unsupported platform", "This program runs only on Windows.")
            sys.exit(1)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        cfg = load_config()
        db_path = cfg.get("db_path", DEFAULT_DB)
        self.db = DBManager(db_path)

        self.root = ctk.CTk()
        self.root.title("openMSX Frontend")
        self.root.geometry("1000x640")
        self.root.minsize(800, 480)

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

        # Start background port watcher to update Client TCP button color
        try:
            threading.Thread(target=self._start_port_watcher, daemon=True).start()
        except Exception:
            pass

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

    def _machines_dir(self, openmsx_dir: str) -> Path:
        return Path(openmsx_dir) / "share" / "machines"

    def _extensions_dir(self, openmsx_dir: str) -> Path:
        return Path(openmsx_dir) / "share" / "extensions"

    def _get_machines(self) -> List[str]:
        openmsx_dir = self.db.get("openmsx_dir") or ""
        machines: List[str] = []
        if openmsx_dir:
            p = self._machines_dir(openmsx_dir)
            if p.exists() and p.is_dir():
                for f in sorted(p.glob("*.xml")):
                    machines.append(f.stem)
        return machines

    def _get_extensions(self) -> List[str]:
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
            saved = self.db.get("openmsx_machine")
            if saved and saved in machines:
                self.machine_var.set(saved)
            else:
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
        exts = self._get_extensions()
        self._extensions_cache = exts

        if self.listbox_extensions:
            lb = self.listbox_extensions
            lb.delete(0, tk.END)
            for item in exts:
                lb.insert(tk.END, item)

            if exts:
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
        try:
            if value:
                self.db.set("openmsx_machine", value)
            else:
                self.db.set("openmsx_machine", "")
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
            cmd = [exe_path, "-machine", machine]

            exts = self._get_selected_extensions()
            for e in exts:
                cmd.extend(["-ext", e])

            proc = subprocess.Popen(cmd, cwd=dir_path)
            initial_pid = proc.pid
            real_pid = initial_pid

            try:
                import psutil
                deadline = time.time() + 2.0
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
                    try:
                        parent = psutil.Process(initial_pid)
                        children = parent.children(recursive=True)
                        if children:
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

            try:
                self.db.set("openmsx_pid", str(real_pid))
            except Exception:
                pass
            self.pid_var.set(f"PID: {real_pid}")

            ext_part = " ".join([f"-ext {e}" for e in exts]) if exts else ""
            self.status_var.set(f"openMSX started: {exe_path} -machine {machine} {ext_part}".strip())

            self._update_socket_button(real_pid)
        except Exception as e:
            messagebox.showerror("Start Error", f"Failed to start openMSX:\n{e}")

    def _update_socket_button(self, pid: int):
        temp_dir = os.getenv("TEMP") or os.getcwd()
        socket_path = os.path.join(temp_dir, "openmsx-default", f"socket.{pid}")
        self.current_socket_path = socket_path
        self.socket_var.set(f"Socket:\n{socket_path}")

        if hasattr(self, "btn_socket"):
            try:
                exists = os.path.exists(socket_path)
                self.btn_socket.configure(state="normal", fg_color="green" if exists else "red")
            except Exception:
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
                subprocess.Popen(['explorer', f'/select,{path}'])
            except Exception:
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

        saved = self.db.get("openmsx_machine")
        if saved and saved in self._machines_cache:
            self.machine_var.set(saved)

    def open_client_window(self):
        try:
            if getattr(self, "_client_window", None):
                try:
                    if self._client_window.win.winfo_exists():
                        self._client_window.win.lift()
                        return
                except Exception:
                    pass
            self._client_window = OpenMSXClientWindow(self.root)
        except Exception as e:
            messagebox.showerror("Client Error", f"Cannot open client window:\n{e}")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        try:
            self.db.close()
        finally:
            self.root.destroy()

    def _check_local_port(self, port: int) -> bool:
        """Return True if a TCP connection to localhost:port can be established quickly."""
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except Exception:
            return False

    def _start_port_watcher(self, interval: float = 2.0):
        """
        Background loop: read the port file and test connection.
        Updates `Client TCP` button color on the GUI thread:
          - green if connection succeeds
          - gray if no open port found or connection fails
        """
        while True:
            try:
                port = find_port_from_temp()
                is_open = False
                if port:
                    is_open = self._check_local_port(port)

                def _update_button(opened=is_open):
                    try:
                        if hasattr(self, "btn_client"):
                            color = "green" if opened else "gray"
                            self.btn_client.configure(fg_color=color)
                    except Exception:
                        pass

                try:
                    self.root.after(0, _update_button)
                except Exception:
                    pass
            except Exception:
                pass
            time.sleep(interval)


if __name__ == "__main__":
    app = OpenMSXFrontend()
    app.run()
