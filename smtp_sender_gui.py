import base64
import json
import os
import smtplib
import ssl
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

try:
    import msal
except Exception:
    msal = None


APP_NAME = "Exchange Online SMTP Sender"
APP_VERSION = "v0.1.0"


@dataclass
class CloudProfile:
    name: str
    authority_host: str


CLOUDS = {
    "Global": CloudProfile("Global", "https://login.microsoftonline.com"),
    "21V (China)": CloudProfile("21V (China)", "https://login.chinacloudapi.cn"),
}


def documents_dir() -> str:
    path = os.path.join(os.path.expanduser("~"), "Documents", "ExchangeOnlineSmtpSender")
    os.makedirs(path, exist_ok=True)
    return path


def today_log_path() -> str:
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(documents_dir(), f"app_{date_str}.log")


def log_file_only(text: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {text}"
    try:
        with open(today_log_path(), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def build_xoauth2_string(user_email: str, access_token: str) -> str:
    # https://developers.google.com/gmail/imap/xoauth2-protocol (format is commonly used)
    auth_string = f"user={user_email}\x01auth=Bearer {access_token}\x01\x01"
    return base64.b64encode(auth_string.encode("utf-8")).decode("utf-8")


def redact_authorization(value: str) -> str:
    if not value:
        return ""
    if value.lower().startswith("bearer "):
        return "Bearer ***"
    return "***"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.minsize(860, 640)

        self.config_path = os.path.join(documents_dir(), "config.json")

        self.cloud_var = tk.StringVar(value="Global")
        self.auth_mode_var = tk.StringVar(value="OAuth2")  # Anonymous / Basic / OAuth2

        self.smtp_host_var = tk.StringVar(value="smtp.office365.com")
        self.smtp_port_var = tk.StringVar(value="587")
        self.use_starttls_var = tk.BooleanVar(value=True)

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()

        # OAuth2 settings
        self.tenant_id_var = tk.StringVar()
        self.client_id_var = tk.StringVar()
        self.client_secret_var = tk.StringVar()
        self.use_device_code_var = tk.BooleanVar(value=True)
        self.resource_scope_var = tk.StringVar(value="https://outlook.office365.com/.default")
        self.save_token_in_log_var = tk.BooleanVar(value=False)

        # Mail fields
        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.cc_var = tk.StringVar()
        self.bcc_var = tk.StringVar()
        self.subject_var = tk.StringVar()

        self._build_menu()
        self._build_ui()
        self._load_config()
        self._sync_auth_ui()

    def _build_menu(self):
        menubar = tk.Menu(self)

        tools = tk.Menu(menubar, tearoff=0)
        tools.add_command(label="打开日志目录 (Open Log Folder)", command=self._open_log_dir)
        tools.add_separator()
        tools.add_command(label="版本历史 (Version History)", command=self._show_history)
        tools.add_command(label="关于 (About)", command=self._show_about)

        menubar.add_cascade(label="工具 (Tools)", menu=tools)
        self.config(menu=menubar)

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill=tk.BOTH, expand=True)

        top = ttk.LabelFrame(root, text="SMTP Settings", padding=10)
        top.pack(fill=tk.X)

        r = 0
        ttk.Label(top, text="Cloud").grid(row=r, column=0, sticky="w")
        ttk.Combobox(top, textvariable=self.cloud_var, values=list(CLOUDS.keys()), width=18, state="readonly").grid(row=r, column=1, sticky="w")
        ttk.Label(top, text="Auth Mode").grid(row=r, column=2, sticky="w", padx=(18, 0))
        ttk.Combobox(top, textvariable=self.auth_mode_var, values=["Anonymous", "Basic", "OAuth2"], width=14, state="readonly").grid(row=r, column=3, sticky="w")
        r += 1

        ttk.Label(top, text="SMTP Host").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(top, textvariable=self.smtp_host_var, width=26).grid(row=r, column=1, sticky="w", pady=(6, 0))
        ttk.Label(top, text="Port").grid(row=r, column=2, sticky="w", padx=(18, 0), pady=(6, 0))
        ttk.Entry(top, textvariable=self.smtp_port_var, width=8).grid(row=r, column=3, sticky="w", pady=(6, 0))
        r += 1

        ttk.Checkbutton(top, text="Use STARTTLS", variable=self.use_starttls_var).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        auth = ttk.LabelFrame(root, text="Authentication", padding=10)
        auth.pack(fill=tk.X, pady=(10, 0))

        r = 0
        ttk.Label(auth, text="Username (for Basic / OAuth2 user)").grid(row=r, column=0, sticky="w")
        ttk.Entry(auth, textvariable=self.username_var, width=40).grid(row=r, column=1, sticky="w")
        r += 1

        ttk.Label(auth, text="Password (Basic)").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(auth, textvariable=self.password_var, width=40, show="*").grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Separator(auth).grid(row=r, column=0, columnspan=3, sticky="ew", pady=10)
        r += 1

        ttk.Label(auth, text="Tenant ID").grid(row=r, column=0, sticky="w")
        ttk.Entry(auth, textvariable=self.tenant_id_var, width=40).grid(row=r, column=1, sticky="w")
        r += 1

        ttk.Label(auth, text="Client ID").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(auth, textvariable=self.client_id_var, width=40).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(auth, text="Client Secret (optional)").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(auth, textvariable=self.client_secret_var, width=40, show="*").grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Checkbutton(auth, text="Use Device Code Flow (recommended)", variable=self.use_device_code_var).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(auth, text="Scope/Resource").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(auth, textvariable=self.resource_scope_var, width=40).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Checkbutton(auth, text="Expert: Save Authorization token in log (dangerous)", variable=self.save_token_in_log_var, command=self._on_toggle_save_token).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        mail = ttk.LabelFrame(root, text="Email", padding=10)
        mail.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        r = 0
        ttk.Label(mail, text="From").grid(row=r, column=0, sticky="w")
        ttk.Entry(mail, textvariable=self.from_var, width=60).grid(row=r, column=1, sticky="w")
        r += 1

        ttk.Label(mail, text="To (comma-separated)").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.to_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Cc").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.cc_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Bcc").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.bcc_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Subject").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.subject_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Body").grid(row=r, column=0, sticky="nw", pady=(6, 0))
        self.body_text = tk.Text(mail, height=10, wrap="word")
        self.body_text.grid(row=r, column=1, sticky="nsew", pady=(6, 0))
        mail.grid_columnconfigure(1, weight=1)
        mail.grid_rowconfigure(r, weight=1)
        r += 1

        actions = ttk.Frame(root)
        actions.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(actions, text="Send Test Email", command=self.send_clicked).pack(side=tk.LEFT)
        ttk.Button(actions, text="Save Config", command=self._save_config).pack(side=tk.LEFT, padx=(8, 0))

        log_frame = ttk.LabelFrame(root, text="Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=False, pady=(10, 0))

        self.log_text = tk.Text(log_frame, height=10, state="disabled")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.auth_mode_var.trace_add("write", lambda *_: self._sync_auth_ui())
        self.cloud_var.trace_add("write", lambda *_: self._sync_cloud_defaults())

    def _log(self, msg: str):
        log_file_only(msg)
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, line + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _open_log_dir(self):
        path = documents_dir()
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _show_about(self):
        messagebox.showinfo("About", f"{APP_NAME}\n{APP_VERSION}\n\nSMTP test sender for Exchange Online")

    def _show_history(self):
        try:
            with open(os.path.join(os.path.dirname(__file__), "CHANGELOG.md"), "r", encoding="utf-8") as f:
                content = f.read()
        except Exception:
            content = "(No changelog found)"

        win = tk.Toplevel(self)
        win.title("Version History")
        win.geometry("760x520")

        txt = tk.Text(win, wrap="word")
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert(tk.END, content)
        txt.configure(state="disabled")

    def _sync_cloud_defaults(self):
        # Only set default host if user didn't customize.
        cloud = self.cloud_var.get()
        if cloud == "Global":
            if self.smtp_host_var.get().strip() == "":
                self.smtp_host_var.set("smtp.office365.com")
        elif cloud == "21V (China)":
            # In practice, tenants commonly still use smtp.office365.com, but keep this configurable.
            if self.smtp_host_var.get().strip() == "":
                self.smtp_host_var.set("smtp.office365.com")

    def _sync_auth_ui(self):
        mode = self.auth_mode_var.get()

        is_basic = mode == "Basic"
        is_oauth = mode == "OAuth2"

        # Basic fields
        # (We keep them visible; user can ignore based on mode)

        # Token logging option only meaningful for OAuth2
        if not is_oauth:
            self.save_token_in_log_var.set(False)

    def _on_toggle_save_token(self):
        if not self.save_token_in_log_var.get():
            return
        confirm = messagebox.askyesno(
            "High Risk",
            "This will write OAuth Authorization tokens into log files.\n"
            "Anyone with access to the log can potentially reuse the token.\n\n"
            "Are you sure?"
        )
        if not confirm:
            self.save_token_in_log_var.set(False)

    def _load_config(self):
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return

        def g(k, default=None):
            return cfg.get(k, default)

        self.cloud_var.set(g("cloud", self.cloud_var.get()))
        self.auth_mode_var.set(g("auth_mode", self.auth_mode_var.get()))
        self.smtp_host_var.set(g("smtp_host", self.smtp_host_var.get()))
        self.smtp_port_var.set(str(g("smtp_port", self.smtp_port_var.get())))
        self.use_starttls_var.set(bool(g("use_starttls", self.use_starttls_var.get())))

        self.username_var.set(g("username", ""))
        self.password_var.set(g("password", ""))

        self.tenant_id_var.set(g("tenant_id", ""))
        self.client_id_var.set(g("client_id", ""))
        self.client_secret_var.set(g("client_secret", ""))
        self.use_device_code_var.set(bool(g("use_device_code", self.use_device_code_var.get())))
        self.resource_scope_var.set(g("resource_scope", self.resource_scope_var.get()))

        self.from_var.set(g("from", ""))
        self.to_var.set(g("to", ""))
        self.cc_var.set(g("cc", ""))
        self.bcc_var.set(g("bcc", ""))
        self.subject_var.set(g("subject", ""))

        body = g("body", "")
        try:
            self.body_text.delete("1.0", tk.END)
            self.body_text.insert(tk.END, body)
        except Exception:
            pass

    def _save_config(self):
        cfg = {
            "cloud": self.cloud_var.get(),
            "auth_mode": self.auth_mode_var.get(),
            "smtp_host": self.smtp_host_var.get(),
            "smtp_port": int(self.smtp_port_var.get() or 587),
            "use_starttls": bool(self.use_starttls_var.get()),
            "username": self.username_var.get(),
            "password": self.password_var.get(),
            "tenant_id": self.tenant_id_var.get(),
            "client_id": self.client_id_var.get(),
            "client_secret": self.client_secret_var.get(),
            "use_device_code": bool(self.use_device_code_var.get()),
            "resource_scope": self.resource_scope_var.get(),
            "from": self.from_var.get(),
            "to": self.to_var.get(),
            "cc": self.cc_var.get(),
            "bcc": self.bcc_var.get(),
            "subject": self.subject_var.get(),
            "body": self.body_text.get("1.0", tk.END).rstrip("\n"),
        }
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self._log(f"Config saved: {self.config_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def send_clicked(self):
        t = threading.Thread(target=self._send_worker, daemon=True)
        t.start()

    def _send_worker(self):
        try:
            self._save_config()
            mode = self.auth_mode_var.get()
            host = self.smtp_host_var.get().strip()
            port = int(self.smtp_port_var.get() or 587)
            use_starttls = bool(self.use_starttls_var.get())

            msg = EmailMessage()
            msg["From"] = self.from_var.get().strip() or self.username_var.get().strip()
            msg["To"] = self.to_var.get().strip()
            if self.cc_var.get().strip():
                msg["Cc"] = self.cc_var.get().strip()
            if self.subject_var.get().strip():
                msg["Subject"] = self.subject_var.get().strip()
            msg.set_content(self.body_text.get("1.0", tk.END))

            to_addrs = []
            for field in [self.to_var.get(), self.cc_var.get(), self.bcc_var.get()]:
                if field:
                    for a in field.split(","):
                        a = a.strip()
                        if a:
                            to_addrs.append(a)

            if not msg["To"]:
                raise Exception("To is required")

            self._log(f"Connecting SMTP {host}:{port} (STARTTLS={use_starttls})")
            server = smtplib.SMTP(host, port, timeout=30)
            server.ehlo()
            if use_starttls:
                context = ssl.create_default_context()
                server.starttls(context=context)
                server.ehlo()

            if mode == "Anonymous":
                self._log("Auth: Anonymous (no login)")
            elif mode == "Basic":
                user = self.username_var.get().strip()
                pwd = self.password_var.get()
                self._log(f"Auth: Basic user={user}")
                server.login(user, pwd)
            else:
                if msal is None:
                    raise Exception("msal is not installed. Run: pip install -r requirements.txt")

                user = self.username_var.get().strip()
                token = self._get_oauth_token()
                token_preview = token if self.save_token_in_log_var.get() else redact_authorization("Bearer " + token)
                self._log(f"Auth: OAuth2 XOAUTH2 user={user} token={token_preview}")

                xoauth2_b64 = build_xoauth2_string(user, token)
                code, resp = server.docmd("AUTH", "XOAUTH2 " + xoauth2_b64)
                if code != 235:
                    raise Exception(f"XOAUTH2 AUTH failed: {code} {resp}")

            self._log(f"Sending mail From={msg['From']} To={to_addrs}")
            server.send_message(msg, from_addr=msg["From"], to_addrs=to_addrs)
            server.quit()
            self._log("Send success")
            messagebox.showinfo("Success", "Email sent successfully")
        except Exception as e:
            self._log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    def _get_oauth_token(self) -> str:
        tenant = self.tenant_id_var.get().strip()
        client_id = self.client_id_var.get().strip()
        client_secret = self.client_secret_var.get().strip()
        scope_or_resource = self.resource_scope_var.get().strip()

        if not tenant or not client_id:
            raise Exception("Tenant ID and Client ID are required for OAuth2")

        cloud = CLOUDS.get(self.cloud_var.get(), CLOUDS["Global"])
        authority = f"{cloud.authority_host}/{tenant}"

        # Accept either a single scope or a space-separated scope list.
        scopes = [s for s in scope_or_resource.split() if s]

        if self.use_device_code_var.get() or not client_secret:
            app = msal.PublicClientApplication(client_id=client_id, authority=authority)
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                raise Exception("Failed to create device flow")

            self._log("Device code flow started. Follow instructions shown in the log.")
            self._log(flow.get("message") or "")

            result = app.acquire_token_by_device_flow(flow)
        else:
            app = msal.ConfidentialClientApplication(client_id=client_id, client_credential=client_secret, authority=authority)
            result = app.acquire_token_for_client(scopes=scopes)

        if not isinstance(result, dict) or "access_token" not in result:
            raise Exception(f"Failed to get token: {result}")

        return result["access_token"]


if __name__ == "__main__":
    App().mainloop()
