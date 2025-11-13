# python
import os
import json
import sqlite3
import subprocess
import sys
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Arquivos de configuração
CONFIG_FILE = Path("app_config.json")
DEFAULT_DB = "app_data.db"


def ensure_config_file():
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps({"db_path": DEFAULT_DB}, indent=2))


def load_config():
    ensure_config_file()
    with CONFIG_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


class DBManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(self.db_path)
        self._ensure_table()

    def _ensure_table(self):
        cur = self.conn.cursor()
        cur.execute(
            "CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT NOT NULL)"
        )
        self.conn.commit()

    def get(self, key):
        cur = self.conn.cursor()
        cur.execute("SELECT value FROM config WHERE key = ?", (key,))
        row = cur.fetchone()
        return row[0] if row else None

    def set(self, key, value):
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO config(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.conn.commit()

    def close(self):
        self.conn.close()


class OpenMSXFrontend:
    def __init__(self):
        if sys.platform != "win32":
            messagebox.showerror("Plataforma não suportada", "Este programa roda apenas no Windows.")
            sys.exit(1)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        cfg = load_config()
        db_path = cfg.get("db_path", DEFAULT_DB)
        self.db = DBManager(db_path)

        self.root = ctk.CTk()
        self.root.title("openMSX Frontend")
        self.root.geometry("700x320")
        self.root.minsize(640, 320)

        self.status_var = ctk.StringVar()
        self.pid_var = ctk.StringVar(value="PID: Não iniciado")
        self.socket_var = ctk.StringVar(value="Socket: -")
        self.current_socket_path = None

        self._build_main_ui()

        if not self.db.get("openmsx_dir"):
            self.open_config_window(initial=True)

    def _build_main_ui(self):
        frame = ctk.CTkFrame(self.root, corner_radius=8)
        frame.pack(padx=20, pady=20, fill="both", expand=True)

        label = ctk.CTkLabel(frame, text="openMSX Frontend", font=ctk.CTkFont(size=20, weight="bold"))
        label.pack(pady=(8, 16))

        btn_start = ctk.CTkButton(frame, text="Iniciar openMSX", command=self.start_openmsx)
        btn_start.pack(pady=6, ipadx=10, ipady=6)

        btn_config = ctk.CTkButton(frame, text="Configuração", command=lambda: self.open_config_window(initial=False))
        btn_config.pack(pady=6, ipadx=10, ipady=6)

        # Linha com PID e Socket em linhas separadas (PID acima, Socket abaixo)
        info_row = ctk.CTkFrame(frame, corner_radius=0)
        info_row.pack(fill="x", pady=(12, 0), padx=10)

        pid_label = ctk.CTkLabel(info_row, textvariable=self.pid_var, anchor="w",
                                 font=ctk.CTkFont(size=12, weight="bold"))
        pid_label.pack(fill="x", anchor="w")

        socket_label = ctk.CTkLabel(info_row, textvariable=self.socket_var, anchor="w",
                                    font=ctk.CTkFont(size=10))
        socket_label.pack(fill="x", pady=(4, 0), anchor="w")

        # Botão Ver Socket
        self.btn_socket = ctk.CTkButton(
            frame,
            text="Ver Socket",
            command=self._check_socket,
            state="disabled",
            fg_color="gray"
        )
        self.btn_socket.pack(pady=6, ipadx=10, ipady=6)

        status_label = ctk.CTkLabel(frame, textvariable=self.status_var, anchor="w")
        status_label.pack(fill="x", pady=(12, 0), padx=10)
        self._update_status()

    def _update_socket_button(self, pid):
        # monta caminho do socket (usa TEMP com fallback)
        temp_dir = os.getenv("TEMP") or os.getcwd()
        socket_path = os.path.join(temp_dir, "openmsx-default", f"socket.{pid}")

        # armazena e mostra o caminho no campo "Socket" da interface
        self.current_socket_path = socket_path
        self.socket_var.set(f"Socket: {socket_path}")

        # atualiza estado/cor do botão conforme existência do arquivo
        if hasattr(self, "btn_socket"):
            if os.path.exists(socket_path):
                self.btn_socket.configure(state="normal", fg_color="green")
            else:
                self.btn_socket.configure(state="normal", fg_color="red")

    def open_config_window(self, initial=False):
        win = ctk.CTkToplevel(self.root)
        win.title("Configuração")
        win.geometry("560x180")
        win.grab_set()

        cur_dir = self.db.get("openmsx_dir") or ""
        entry_var = ctk.StringVar(value=cur_dir)

        lbl = ctk.CTkLabel(win, text="Diretório do openMSX (pasta contendo openmsx.exe):")
        lbl.pack(padx=12, pady=(12, 6), anchor="w")

        # Frame para entrada + botão selecionar
        entry_row = ctk.CTkFrame(win, corner_radius=0)
        entry_row.pack(fill="x", padx=12, pady=(0, 6))

        entry = ctk.CTkEntry(entry_row, textvariable=entry_var)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        def choose_dir():
            d = filedialog.askdirectory(title="Selecione o diretório do openMSX")
            if d:
                entry_var.set(d)

        btn_choose = ctk.CTkButton(entry_row, text="Selecionar", width=120, command=choose_dir)
        btn_choose.pack(side="right")

        # Buttons frame
        btn_frame = ctk.CTkFrame(win, corner_radius=0)
        btn_frame.pack(fill="x", padx=12, pady=(8, 12))

        def reset():
            entry_var.set("")

        def save():
            path = entry_var.get().strip()
            if not path:
                messagebox.showwarning("Validação", "O diretório não pode ficar vazio.")
                return
            exe_path = os.path.join(path, "openmsx.exe")
            if not os.path.isfile(exe_path):
                messagebox.showerror("Arquivo não encontrado", f"Não foi encontrado `openmsx.exe` em:\n{exe_path}")
                return
            self.db.set("openmsx_dir", path)
            self._update_status()
            win.destroy()

        def cancel():
            win.destroy()
            if initial and not self.db.get("openmsx_dir"):
                messagebox.showinfo("Configuração necessária",
                                    "Configuração não concluída. Você pode configurar depois via botão `Configuração`.")

        # Botão Reset à esquerda
        btn_reset = ctk.CTkButton(btn_frame, text="Resetar", command=reset)
        btn_reset.pack(side="left", padx=6, pady=6)

        # Espaçador que expande para empurrar os botões finais para a direita
        spacer = ctk.CTkLabel(btn_frame, text="")
        spacer.pack(side="left", expand=True)

        # Botões Cancelar e Salvar à direita
        btn_cancel = ctk.CTkButton(btn_frame, text="Cancelar", command=cancel)
        btn_cancel.pack(side="right", padx=6, pady=6)

        btn_save = ctk.CTkButton(btn_frame, text="Salvar", command=save)
        btn_save.pack(side="right", padx=6, pady=6)

    def start_openmsx(self):
        dir_path = self.db.get("openmsx_dir")
        if not dir_path:
            messagebox.showerror("Erro", "Nenhum diretório do openMSX configurado.")
            return
        exe_path = os.path.join(dir_path, "openmsx.exe")
        if not os.path.isfile(exe_path):
            messagebox.showerror("Erro", f"`openmsx.exe` não encontrado em:\n{exe_path}")
            return
        try:
            proc = subprocess.Popen([exe_path], cwd=dir_path)
            initial_pid = proc.pid
            real_pid = initial_pid

            # Tenta usar psutil para localizar o processo real `openmsx.exe`
            try:
                import time
                import psutil
                time.sleep(0.5)  # aguarda o processo principal/filhos aparecerem

                matches = []
                for p in psutil.process_iter(['pid', 'exe', 'ppid']):
                    try:
                        exe = p.info.get('exe')
                        if exe and os.path.normcase(os.path.normpath(exe)) == os.path.normcase(
                                os.path.normpath(exe_path)):
                            matches.append((p.pid, p))
                    except Exception:
                        continue

                if matches:
                    # escolhe o último match (mais provável ser o processo ativo)
                    real_pid = matches[-1][0]
                else:
                    # se não encontrou por exe, procura filhos do processo inicial
                    try:
                        parent = psutil.Process(initial_pid)
                        children = parent.children(recursive=True)
                        if children:
                            # pega o primeiro filho (ou escolha outra heurística se necessário)
                            real_pid = children[0].pid
                    except Exception:
                        pass

            except ImportError:
                messagebox.showwarning("Aviso",
                                       "Para detectar o PID real do `openmsx.exe` instale o pacote `psutil` (pip install psutil). Usando PID inicial do processo.")
            except Exception:
                # falha silenciosa: mantém initial_pid
                pass

            # Atualiza DB e UI com o PID encontrado
            self.db.set("openmsx_pid", str(real_pid))
            self.pid_var.set(f"PID: {real_pid}")
            self.status_var.set(f"openMSX iniciado: {exe_path}")

            # monta e mostra o caminho do socket imediatamente
            temp_dir = os.getenv("TEMP") or os.getcwd()
            socket_path = os.path.join(temp_dir, "openmsx-default", f"socket.{real_pid}")
            self.current_socket_path = socket_path
            self.socket_var.set(f"Socket: {socket_path}")

            # atualiza botão (cor/estado) conforme existência do arquivo
            if hasattr(self, "_update_socket_button"):
                try:
                    self._update_socket_button(real_pid)
                except Exception:
                    # garantir que botão não quebre
                    if hasattr(self, "btn_socket"):
                        self.btn_socket.configure(state="normal", fg_color="red")
        except Exception as e:
            messagebox.showerror("Erro ao iniciar", f"Falha ao iniciar openMSX:\n{e}")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        self.db.close()
        self.root.destroy()

    def _update_socket_button(self, pid):
        temp_dir = os.getenv("TEMP")
        socket_path = os.path.join(temp_dir, "openmsx-default", f"socket.{pid}")

        if os.path.exists(socket_path):
            self.btn_socket.configure(state="normal", fg_color="green")
        else:
            self.btn_socket.configure(state="normal", fg_color="red")

    def _check_socket(self):
        path = getattr(self, "current_socket_path", None)
        if not path:
            messagebox.showwarning("Socket não definido", "Nenhum caminho de socket disponível.")
            return

        if os.path.exists(path):
            subprocess.Popen(f'explorer /select,"{path}"')
        else:
            messagebox.showwarning("Socket não encontrado", f"O socket não existe em:\n{path}")

    def _update_status(self):
        # Atualiza status do diretório openMSX
        openmsx_dir = self.db.get("openmsx_dir")
        if openmsx_dir:
            self.status_var.set(f"openMSX: {openmsx_dir}")
        else:
            self.status_var.set("openMSX não configurado")

        # Recupera PID salvo no DB e atualiza UI
        pid = self.db.get("openmsx_pid")
        if pid:
            try:
                pid_int = int(pid)
                self.pid_var.set(f"PID: {pid_int}")
                # Atualiza caminho do socket e botão
                self._update_socket_button(pid_int)
            except (ValueError, TypeError):
                self.pid_var.set("PID: inválido")
                self.socket_var.set("Socket: -")
                self.current_socket_path = None
                if hasattr(self, "btn_socket"):
                    self.btn_socket.configure(state="disabled", fg_color="gray")
        else:
            self.pid_var.set("PID: Não iniciado")
            self.socket_var.set("Socket: -")
            self.current_socket_path = None
            if hasattr(self, "btn_socket"):
                self.btn_socket.configure(state="disabled", fg_color="gray")

if __name__ == "__main__":
    app = OpenMSXFrontend()
    app.run()
