import base64
import ctypes
import json
import os
import smtplib
import ssl
import sys
import threading
import time
import re
import mimetypes
import html
from html.parser import HTMLParser
import webbrowser
import urllib.error
import urllib.request
import subprocess
from urllib.parse import urlparse
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
import textwrap

import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinter import colorchooser

APP_NAME = "Exchange Online SMTP Sender"
APP_VERSION = "1.0.0"

WINDOWS_APP_ID = "andylu1988.ExchangeOnlineSmtpSender"


def enable_windows_dpi_awareness() -> None:
    """Enable DPI awareness on Windows for correct scaling on high-DPI displays."""
    if os.name != "nt":
        return
    try:
        # 2 = Per-monitor DPI aware
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_EMAIL_RE = re.compile(r"\b([A-Za-z0-9_.+\-]+)@([A-Za-z0-9\-]+\.[A-Za-z0-9.\-]+)\b")


def redact_email_addresses(text: str) -> str:
    if not text:
        return ""
    return _EMAIL_RE.sub(r"***@\2", text)


def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


def set_windows_appusermodelid(app_id: str) -> None:
    if not app_id:
        return
    try:
        if os.name == "nt":
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


@dataclass
class CloudProfile:
    name: str
    authority_host: str


CLOUDS = {
    "Global": CloudProfile("Global", "https://login.microsoftonline.com"),
    "21V (China)": CloudProfile("21V (China)", "https://login.chinacloudapi.cn"),
}


DEFAULT_SMTP_HOST = {
    "Global": "smtp.office365.com",
    "21V (China)": "smtp.partner.outlook.cn",
}


EOP_DOMAIN_SUFFIX = {
    "Global": "mail.protection.outlook.com",
    "21V (China)": "mail.protection.partner.outlook.cn",
}


DEFAULT_SCOPE = {
    "Global": "https://outlook.office365.com/.default",
    "21V (China)": "https://partner.outlook.cn/.default",
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


def today_smtp_debug_log_path() -> str:
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(documents_dir(), f"smtp_debug_{date_str}.log")


_SMTP_AUTH_RE = re.compile(r"\bAUTH\b\s+(XOAUTH2|PLAIN|LOGIN)\b", re.IGNORECASE)


def redact_smtp_protocol_line(line: str) -> str:
    """Redacts sensitive SMTP AUTH payloads from smtplib debug output."""
    if not line:
        return ""

    m = _SMTP_AUTH_RE.search(line)
    if not m:
        return line

    mech = m.group(1).upper()
    # Keep the mechanism, but redact the payload which may contain secrets/tokens.
    prefix = line[: m.end()]
    suffix = line[m.end() :]
    # Try to estimate payload length (best-effort) for debugging.
    payload_len = 0
    try:
        # Common patterns include: "AUTH XOAUTH2 <b64>\r\n" or quoted variants.
        suffix_clean = suffix
        suffix_clean = suffix_clean.replace("\\r", " ").replace("\\n", " ")
        suffix_clean = suffix_clean.strip(" '\"\t")
        # Take first token-ish chunk.
        payload = suffix_clean.strip().split()[0] if suffix_clean.strip() else ""
        payload_len = len(payload)
    except Exception:
        payload_len = 0

    return f"{prefix} [REDACTED payload_len={payload_len} mech={mech}]"


class FileDebugSMTP(smtplib.SMTP):
    def __init__(
        self,
        host: str,
        port: int,
        *,
        timeout: int,
        debug_log_path: str,
        redact_data: bool = True,
        redact_addresses: bool = True,
    ):
        self._debug_log_path = debug_log_path
        self._redact_data = bool(redact_data)
        self._redact_addresses = bool(redact_addresses)
        self._dedupe_reply_retcode = False
        self._dedupe_reply_data_tuple = False
        super().__init__(host, port, timeout=timeout)

    def _maybe_dedupe_debug_line(self, text: str) -> str | None:
        """Collapse redundant smtplib debug lines into a single line.

        smtplib often emits three lines for the same server response:
        - reply: b'250 ...\r\n'
        - reply: retcode (250); Msg: b'...'
        - data: (250, b'...')

        We keep the first and suppress the redundant ones.
        """
        if not text:
            return text

        t = text.lstrip()
        tl = t.lower()

        # Suppress redundant parsed reply line.
        if self._dedupe_reply_retcode and tl.startswith("reply:") and "retcode" in tl:
            self._dedupe_reply_retcode = False
            return None

        # Suppress redundant tuple line for DATA return value (not message DATA itself).
        if self._dedupe_reply_data_tuple and tl.startswith("data:"):
            # Only suppress the tuple summary: data: (250, b'...')
            if re.match(r"^data:\s*\(\s*\d+\s*,\s*b'", tl):
                self._dedupe_reply_data_tuple = False
                return None

        # If this is the raw reply bytes line, keep it but mark subsequent lines to be suppressed.
        if tl.startswith("reply:") and "b'" in tl:
            self._dedupe_reply_retcode = True
            self._dedupe_reply_data_tuple = True

            # Normalize: strip leading "reply:" and unwrap b'...\r\n'
            # Keep original text if parsing fails.
            try:
                # Example: reply: b'250 2.0.0 OK ...\r\n'
                raw = t.split("reply:", 1)[1].strip()
                if raw.startswith("b'") and raw.endswith("'"):
                    inner = raw[2:-1]
                    inner = inner.replace("\\r", "").replace("\\n", "").strip()
                    return f"reply: {inner}"
            except Exception:
                return text

        return text

    def _print_debug(self, *args):
        try:
            text = " ".join(str(a) for a in args)
            # Dedupe/normalize before any redaction so patterns still match.
            text = self._maybe_dedupe_debug_line(text)
            if text is None:
                return

            kind = text.lstrip().lower()
            direction = "LOG"
            if kind.startswith("send:"):
                direction = "C->S"
            elif kind.startswith("reply:") or kind.startswith("data:"):
                direction = "S->C"

            if self._redact_data and kind.startswith("data:"):
                # smtplib debug output includes message DATA (headers/body/attachments).
                # Redact it by default to avoid leaking sensitive content and attachment raw data.
                text = "data: [REDACTED]"

            if self._redact_addresses:
                # smtplib debug output includes SMTP envelope recipients (MAIL FROM / RCPT TO).
                # Redact email addresses by default to avoid logging recipient lists.
                text = redact_email_addresses(text)

            text = redact_smtp_protocol_line(text)
            text = f"{direction}: {text}"
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            os.makedirs(os.path.dirname(self._debug_log_path) or ".", exist_ok=True)
            with open(self._debug_log_path, "a", encoding="utf-8") as f:
                f.write(f"[{ts}] {text}\n")
        except Exception:
            # Never break SMTP flow due to logging issues.
            pass


def prompt_text(parent: tk.Tk, title: str, label: str, *, initial: str = "") -> str:
    value_var = tk.StringVar(value=initial)
    result = {"value": ""}

    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("520x140")
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm, text=label).pack(anchor="w")
    entry = ttk.Entry(frm, textvariable=value_var)
    entry.pack(fill=tk.X, pady=(8, 10))
    entry.focus_set()

    btns = ttk.Frame(frm)
    btns.pack(fill=tk.X)

    def ok():
        result["value"] = (value_var.get() or "").strip()
        win.destroy()

    def cancel():
        result["value"] = ""
        win.destroy()

    ttk.Button(btns, text="OK", command=ok).pack(side=tk.RIGHT)
    ttk.Button(btns, text="Cancel", command=cancel).pack(side=tk.RIGHT, padx=(0, 8))

    win.bind("<Return>", lambda _e: ok())
    win.bind("<Escape>", lambda _e: cancel())

    parent.wait_window(win)
    return result["value"]


class _BasicHtmlToTkParser(HTMLParser):
    """Minimal HTML -> Tk Text parser for our editor.

    Supports: <b>/<strong>, <i>/<em>, <u>, <a href>, <br>, <p>/<div>.
    """

    def __init__(
        self,
        text_widget: tk.Text,
        *,
        link_tag_factory,
        font_family_tag_factory=None,
        font_size_tag_factory=None,
        color_tag_factory=None,
    ):
        super().__init__(convert_charrefs=True)
        self.text = text_widget
        self.active_tags: list[str] = []
        self.link_tag_factory = link_tag_factory
        self.font_family_tag_factory = font_family_tag_factory
        self.font_size_tag_factory = font_size_tag_factory
        self.color_tag_factory = color_tag_factory
        self._block_started = False

    def handle_starttag(self, tag, attrs):
        tag = (tag or "").lower()
        attrs = dict(attrs or [])

        if tag in ("b", "strong"):
            self.active_tags.append("rt_bold")
        elif tag in ("i", "em"):
            self.active_tags.append("rt_italic")
        elif tag == "u":
            self.active_tags.append("rt_underline")
        elif tag == "a":
            href = (attrs.get("href") or "").strip()
            if href:
                self.active_tags.append(self.link_tag_factory(href))
        elif tag == "span":
            style = (attrs.get("style") or "").strip()
            if style:
                style_dict = {}
                for part in style.split(";"):
                    part = part.strip()
                    if not part or ":" not in part:
                        continue
                    k, v = part.split(":", 1)
                    style_dict[k.strip().lower()] = v.strip()

                ff = style_dict.get("font-family")
                if ff and self.font_family_tag_factory:
                    # take first family in list
                    ff = ff.split(",", 1)[0].strip().strip("\"'")
                    if ff:
                        self.active_tags.append(self.font_family_tag_factory(ff))

                fs = style_dict.get("font-size")
                if fs and self.font_size_tag_factory:
                    fs = fs.lower().strip()
                    size_pt = None
                    try:
                        if fs.endswith("pt"):
                            size_pt = int(float(fs[:-2].strip()))
                        elif fs.endswith("px"):
                            px = float(fs[:-2].strip())
                            size_pt = int(round(px * 0.75))
                        else:
                            size_pt = int(float(fs))
                    except Exception:
                        size_pt = None
                    if size_pt:
                        self.active_tags.append(self.font_size_tag_factory(size_pt))

                col = style_dict.get("color")
                if col and self.color_tag_factory:
                    col = col.strip()
                    if col:
                        self.active_tags.append(self.color_tag_factory(col))
        elif tag == "br":
            self.text.insert(tk.END, "\n", tuple(self.active_tags))
        elif tag in ("p", "div"):
            if self._block_started:
                self.text.insert(tk.END, "\n", tuple(self.active_tags))
            self._block_started = True

    def handle_endtag(self, tag):
        tag = (tag or "").lower()
        if tag in ("b", "strong"):
            self._remove_last("rt_bold")
        elif tag in ("i", "em"):
            self._remove_last("rt_italic")
        elif tag == "u":
            self._remove_last("rt_underline")
        elif tag == "a":
            for i in range(len(self.active_tags) - 1, -1, -1):
                if self.active_tags[i].startswith("rt_link_"):
                    self.active_tags.pop(i)
                    break
        elif tag in ("p", "div"):
            self.text.insert(tk.END, "\n", tuple(self.active_tags))

    def handle_data(self, data):
        if not data:
            return
        self.text.insert(tk.END, data, tuple(self.active_tags))

    def _remove_last(self, tag_name: str):
        for i in range(len(self.active_tags) - 1, -1, -1):
            if self.active_tags[i] == tag_name:
                self.active_tags.pop(i)
                break


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
        set_windows_appusermodelid(WINDOWS_APP_ID)
        try:
            icon_path = resource_path(os.path.join("assets", "smtp_tool.ico"))
            if os.path.exists(icon_path):
                self.iconbitmap(default=icon_path)
        except Exception:
            pass

        self._apply_compact_theme()
        self.title(APP_NAME)
        self.minsize(860, 640)

        self.config_path = os.path.join(documents_dir(), "config.json")

        self._current_profile_cloud = None
        self._is_switching_cloud_profile = False

        self.cloud_var = tk.StringVar(value="Global")
        self.auth_mode_var = tk.StringVar(value="OAuth2")  # Anonymous / Basic / OAuth2

        self.smtp_host_var = tk.StringVar(value=DEFAULT_SMTP_HOST["Global"])
        self.smtp_port_var = tk.StringVar(value="587")
        self.use_starttls_var = tk.BooleanVar(value=True)

        # Remember per-auth-mode SMTP settings so switching modes restores expected defaults.
        self._smtp_settings_by_mode: dict[str, dict[str, str | bool]] = {}
        self._last_auth_mode = (self.auth_mode_var.get() or "").strip() or "OAuth2"

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()

        # OAuth2 settings
        self.tenant_id_var = tk.StringVar()
        self.client_id_var = tk.StringVar()
        self.client_secret_var = tk.StringVar()
        self.resource_scope_var = tk.StringVar(value=DEFAULT_SCOPE["Global"])
        self.save_token_in_log_var = tk.BooleanVar(value=False)

        # SMTP protocol debug logging (writes smtplib debug output to a dedicated file)
        self.smtp_debug_enabled_var = tk.BooleanVar(value=False)
        self.smtp_debug_log_path_var = tk.StringVar(value=today_smtp_debug_log_path())
        self.smtp_debug_include_headers_var = tk.BooleanVar(value=True)
        self.smtp_debug_include_data_var = tk.BooleanVar(value=False)
        self.smtp_debug_redact_addresses_var = tk.BooleanVar(value=True)

        # Scheduled sending
        self.schedule_interval_var = tk.StringVar(value="60")
        self.schedule_times_var = tk.StringVar(value="5")
        self._schedule_cancel_event = threading.Event()
        self._schedule_thread = None
        self._schedule_running = False
        self._schedule_start_btn = None
        self._schedule_stop_btn = None

        # Mail fields
        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.cc_var = tk.StringVar()
        self.bcc_var = tk.StringVar()
        self.subject_var = tk.StringVar()

        # Body editors
        self.body_mode_var = tk.StringVar(value="plain")  # plain/html
        self._plain_body_text = None
        self._html_body_text = None  # Rich-text editor widget
        self._rt_font_normal = None
        self._rt_font_bold = None
        self._rt_font_italic = None
        self._rt_link_map: dict[str, str] = {}
        self._rt_link_counter = 0
        self._rt_family_var = tk.StringVar(value="Calibri")
        self._rt_size_var = tk.StringVar(value="11")
        self._rt_family_tag_map: dict[str, str] = {}
        self._rt_family_cache: dict[str, str] = {}
        self._rt_size_cache: dict[int, str] = {}
        self._rt_color_cache: dict[str, str] = {}

        # Anonymous MX lookup cache (avoid UI lag)
        self._anon_mx_cache: dict[str, str] = {}
        self._anon_mx_inflight: set[str] = set()

        # Notebook refs (switch to Log tab without popups)
        self._main_notebook = None
        self._log_tab = None

        # Attachments
        self._attachments: list[str] = []
        self._attachments_listbox = None

        self._build_menu()
        self._build_ui()
        self._load_config()
        self._sync_auth_ui()

    def _apply_compact_theme(self) -> None:
        try:
            style = ttk.Style(self)
        except Exception:
            return

        # Keep native Windows look when available, but tighten spacing via style paddings.
        try:
            preferred = "vista" if "vista" in style.theme_names() else None
            if preferred:
                style.theme_use(preferred)
        except Exception:
            pass

        try:
            base_font = tkfont.nametofont("TkDefaultFont")
            base_font.configure(size=9)
            self.option_add("*Font", base_font)
        except Exception:
            pass

        try:
            style.configure("TNotebook.Tab", padding=(10, 4))
            style.configure("TButton", padding=(10, 4))
            style.configure("TCheckbutton", padding=(6, 2))
            style.configure("TLabelframe", padding=(10, 8))
            style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"))
        except Exception:
            pass

    def _build_menu(self):
        menubar = tk.Menu(self)

        tools = tk.Menu(menubar, tearoff=0)
        tools.add_command(label="打开日志目录 (Open Log Folder)", command=self._open_log_dir)
        tools.add_separator()
        tools.add_command(label="关于 (About)", command=self._show_about)

        menubar.add_cascade(label="工具 (Tools)", menu=tools)
        self.config(menu=menubar)

    def _build_ui(self):
        root = ttk.Frame(self, padding=8)
        root.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True)

        send_tab = ttk.Frame(notebook, padding=0)
        log_tab = ttk.Frame(notebook, padding=0)
        notebook.add(send_tab, text="Send")
        notebook.add(log_tab, text="Log")

        self._main_notebook = notebook
        self._log_tab = log_tab

        top = ttk.LabelFrame(send_tab, text="SMTP Settings", padding=8)
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

        auth = ttk.LabelFrame(send_tab, text="Authentication", padding=8)
        auth.pack(fill=tk.X, pady=(6, 0))

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

        ttk.Label(auth, text="Client Secret").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(auth, textvariable=self.client_secret_var, width=40, show="*").grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(auth, text="Scope").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(auth, textvariable=self.resource_scope_var, width=40).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Checkbutton(auth, text="Expert: Save Authorization token in log (dangerous)", variable=self.save_token_in_log_var, command=self._on_toggle_save_token).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        mail = ttk.LabelFrame(send_tab, text="Email", padding=8)
        mail.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        r = 0
        ttk.Label(mail, text="From").grid(row=r, column=0, sticky="w")
        ttk.Entry(mail, textvariable=self.from_var, width=60).grid(row=r, column=1, sticky="w")

        sched = ttk.LabelFrame(mail, text="Scheduled Send", padding=6)
        sched.grid(row=r, column=2, rowspan=2, sticky="ne", padx=(12, 0))
        ttk.Label(sched, text="Every (seconds)").grid(row=0, column=0, sticky="w")
        ttk.Entry(sched, textvariable=self.schedule_interval_var, width=10).grid(row=0, column=1, sticky="w", padx=(6, 18))
        ttk.Label(sched, text="Times").grid(row=0, column=2, sticky="w")
        ttk.Entry(sched, textvariable=self.schedule_times_var, width=8).grid(row=0, column=3, sticky="w", padx=(6, 18))
        self._schedule_start_btn = ttk.Button(sched, text="Start", command=self._start_scheduled_send)
        self._schedule_start_btn.grid(row=0, column=4, sticky="w")
        self._schedule_stop_btn = ttk.Button(sched, text="Stop", command=self._stop_scheduled_send)
        self._schedule_stop_btn.grid(row=0, column=5, sticky="w", padx=(8, 0))
        sched.grid_columnconfigure(6, weight=1)
        self._sync_schedule_buttons()
        r += 1

        ttk.Label(mail, text="To (comma/semicolon-separated)").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.to_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Cc (comma/semicolon-separated)").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.cc_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Bcc (comma/semicolon-separated)").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.bcc_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Subject").grid(row=r, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(mail, textvariable=self.subject_var, width=60).grid(row=r, column=1, sticky="w", pady=(6, 0))
        r += 1

        ttk.Label(mail, text="Body").grid(row=r, column=0, sticky="nw", pady=(6, 0))
        body_notebook = ttk.Notebook(mail)
        body_notebook.grid(row=r, column=1, columnspan=2, sticky="nsew", pady=(6, 0))

        plain_body_tab = ttk.Frame(body_notebook)
        html_body_tab = ttk.Frame(body_notebook)
        body_notebook.add(plain_body_tab, text="Plain")
        body_notebook.add(html_body_tab, text="HTML")

        self._plain_body_text = tk.Text(plain_body_tab, height=10, wrap="word")
        self._plain_body_text.pack(fill=tk.BOTH, expand=True)

        # Rich-text HTML editor (toolbar + WYSIWYG)
        rt_toolbar = ttk.Frame(html_body_tab)
        rt_toolbar.pack(fill=tk.X)

        ttk.Button(rt_toolbar, text="B", width=3, command=lambda: self._rt_toggle_tag("rt_bold")).pack(side=tk.LEFT)
        ttk.Button(rt_toolbar, text="I", width=3, command=lambda: self._rt_toggle_tag("rt_italic")).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(rt_toolbar, text="U", width=3, command=lambda: self._rt_toggle_tag("rt_underline")).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Separator(rt_toolbar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8)

        ttk.Label(rt_toolbar, text="Font").pack(side=tk.LEFT)
        font_box = ttk.Combobox(
            rt_toolbar,
            textvariable=self._rt_family_var,
            values=["Calibri", "Arial", "Segoe UI", "Times New Roman", "Courier New"],
            width=16,
            state="readonly",
        )
        font_box.pack(side=tk.LEFT, padx=(6, 10))
        font_box.bind("<<ComboboxSelected>>", lambda _e: self._rt_apply_font_family(self._rt_family_var.get()))

        ttk.Label(rt_toolbar, text="Size").pack(side=tk.LEFT)
        size_box = ttk.Combobox(
            rt_toolbar,
            textvariable=self._rt_size_var,
            values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"],
            width=5,
            state="readonly",
        )
        size_box.pack(side=tk.LEFT, padx=(6, 10))
        size_box.bind("<<ComboboxSelected>>", lambda _e: self._rt_apply_font_size(self._rt_size_var.get()))

        ttk.Button(rt_toolbar, text="Color", width=6, command=self._rt_choose_color).pack(side=tk.LEFT)

        ttk.Button(rt_toolbar, text="•", width=3, command=self._rt_bullet_list).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(rt_toolbar, text="1.", width=3, command=self._rt_numbered_list).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(rt_toolbar, text="Link", width=6, command=self._rt_insert_link).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(rt_toolbar, text="Clear", width=6, command=self._rt_clear_formatting).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(rt_toolbar, text="View HTML", width=9, command=self._rt_show_html_preview).pack(side=tk.RIGHT)

        self._html_body_text = tk.Text(html_body_tab, height=10, wrap="word", undo=True)
        self._html_body_text.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self._rt_init_tags()

        def _on_body_tab_changed(_evt=None):
            try:
                idx = body_notebook.index(body_notebook.select())
                self.body_mode_var.set("plain" if idx == 0 else "html")
            except Exception:
                pass

        body_notebook.bind("<<NotebookTabChanged>>", _on_body_tab_changed)
        _on_body_tab_changed()
        mail.grid_columnconfigure(1, weight=1)
        mail.grid_columnconfigure(2, weight=0)
        mail.grid_rowconfigure(r, weight=1)
        r += 1

        # Attachments
        attach_frame = ttk.LabelFrame(mail, text="Attachments", padding=6)
        attach_frame.grid(row=r, column=1, columnspan=2, sticky="nsew", pady=(8, 0))

        self._attachments_listbox = tk.Listbox(attach_frame, height=4)
        self._attachments_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        attach_buttons = ttk.Frame(attach_frame)
        attach_buttons.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))
        ttk.Button(attach_buttons, text="Add...", command=self._add_attachments).pack(fill=tk.X)
        ttk.Button(attach_buttons, text="Remove", command=self._remove_selected_attachment).pack(fill=tk.X, pady=(6, 0))
        ttk.Button(attach_buttons, text="Clear", command=self._clear_attachments).pack(fill=tk.X, pady=(6, 0))
        r += 1

        actions = ttk.Frame(send_tab)
        actions.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(actions, text="Send Test Email", command=self.send_clicked).pack(side=tk.LEFT)
        ttk.Button(actions, text="Save Config", command=self._save_config).pack(side=tk.LEFT, padx=(8, 0))

        log_frame = ttk.LabelFrame(log_tab, text="Log", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        dbg = ttk.LabelFrame(log_frame, text="SMTP Protocol Debug", padding=6)
        dbg.pack(fill=tk.X, pady=(0, 8))
        ttk.Checkbutton(
            dbg,
            text="Enable SMTP protocol debug log (AUTH payload redacted)",
            variable=self.smtp_debug_enabled_var,
        ).pack(anchor="w")

        ttk.Checkbutton(
            dbg,
            text="Include message headers in SMTP debug log",
            variable=self.smtp_debug_include_headers_var,
        ).pack(anchor="w", pady=(4, 0))

        ttk.Checkbutton(
            dbg,
            text="Include SMTP DATA (message content/attachments) in debug log (dangerous)",
            variable=self.smtp_debug_include_data_var,
        ).pack(anchor="w", pady=(4, 0))

        ttk.Checkbutton(
            dbg,
            text="Redact recipient addresses in debug log (recommended)",
            variable=self.smtp_debug_redact_addresses_var,
        ).pack(anchor="w", pady=(4, 0))

        path_row = ttk.Frame(dbg)
        path_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(path_row, text="SMTP debug log file").pack(side=tk.LEFT)
        ttk.Entry(path_row, textvariable=self.smtp_debug_log_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        ttk.Button(path_row, text="Browse", command=self._browse_smtp_debug_log).pack(side=tk.LEFT)

        self.log_text = tk.Text(log_frame, height=16, state="disabled")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.auth_mode_var.trace_add("write", lambda *_: self._sync_auth_ui())
        self.cloud_var.trace_add("write", lambda *_: self._on_cloud_changed())
        self.to_var.trace_add("write", lambda *_: self._on_to_changed())

    def _show_log_tab(self):
        try:
            if self._main_notebook is None or self._log_tab is None:
                return

            if threading.current_thread() is threading.main_thread():
                self._main_notebook.select(self._log_tab)
            else:
                self.after(0, lambda: self._main_notebook.select(self._log_tab))
        except Exception:
            pass

    def _sync_schedule_buttons(self):
        try:
            running = bool(self._schedule_running)
            if self._schedule_start_btn is not None:
                self._schedule_start_btn.configure(state=("disabled" if running else "normal"))
            if self._schedule_stop_btn is not None:
                self._schedule_stop_btn.configure(state=("normal" if running else "disabled"))
        except Exception:
            pass

    def _start_scheduled_send(self):
        if self._schedule_running:
            return

        try:
            interval = float((self.schedule_interval_var.get() or "").strip() or "0")
            times = int((self.schedule_times_var.get() or "").strip() or "0")
        except Exception:
            self._log("ERROR: Scheduled Send: invalid interval/times")
            self._show_log_tab()
            return

        if interval <= 0 or times <= 0:
            self._log("ERROR: Scheduled Send: interval and times must be > 0")
            self._show_log_tab()
            return

        # Snapshot current email/settings once (thread-safe).
        try:
            self._save_config()
        except Exception:
            pass
        snapshot = self._collect_send_snapshot()

        self._schedule_cancel_event.clear()
        self._schedule_running = True
        self._sync_schedule_buttons()

        self._schedule_thread = threading.Thread(
            target=self._scheduled_send_worker,
            args=(snapshot, interval, times),
            daemon=True,
        )
        self._schedule_thread.start()
        self._log(f"Scheduled Send started: every {interval}s x{times}")

    def _stop_scheduled_send(self):
        if not self._schedule_running:
            return
        self._schedule_cancel_event.set()
        self._log("Scheduled Send stopping...")

    def _scheduled_send_worker(self, snapshot: dict, interval: float, times: int):
        try:
            for i in range(1, times + 1):
                if self._schedule_cancel_event.is_set():
                    break
                self._log(f"Scheduled Send: attempt {i}/{times}")
                ok = self._send_worker(snapshot=snapshot, job_label=f"Scheduled {i}/{times}")
                if not ok:
                    self._log("Scheduled Send stopped due to error")
                    break
                if i < times:
                    self._schedule_cancel_event.wait(interval)
        finally:
            self._schedule_running = False
            try:
                self.after(0, self._sync_schedule_buttons)
            except Exception:
                pass
            self._log("Scheduled Send finished")

    def _collect_send_snapshot(self) -> dict:
        # Must be called on the UI thread.
        mode = (self.auth_mode_var.get() or "").strip()
        cloud = (self.cloud_var.get() or "").strip()
        host = (self.smtp_host_var.get() or "").strip().strip(".")
        port = int(self.smtp_port_var.get() or 587)
        use_starttls = bool(self.use_starttls_var.get())

        smtp_debug_enabled = bool(self.smtp_debug_enabled_var.get())
        smtp_debug_path = (self.smtp_debug_log_path_var.get() or "").strip() or today_smtp_debug_log_path()
        smtp_debug_include_headers = bool(self.smtp_debug_include_headers_var.get())
        smtp_debug_include_data = bool(self.smtp_debug_include_data_var.get())
        smtp_debug_redact_addresses = bool(self.smtp_debug_redact_addresses_var.get())

        from_addr = (self.from_var.get() or "").strip()
        to_addr = (self.to_var.get() or "").strip()
        cc_addr = (self.cc_var.get() or "").strip()
        bcc_addr = (self.bcc_var.get() or "").strip()
        subject = (self.subject_var.get() or "").strip()

        plain_body = ""
        html_body = ""
        try:
            if self._plain_body_text is not None:
                plain_body = self._plain_body_text.get("1.0", tk.END)
        except Exception:
            plain_body = ""
        try:
            if self._html_body_text is not None:
                html_body = self._rt_export_html()
        except Exception:
            html_body = ""

        return {
            "mode": mode,
            "cloud": cloud,
            "host": host,
            "port": port,
            "use_starttls": use_starttls,
            "smtp_debug_enabled": smtp_debug_enabled,
            "smtp_debug_path": smtp_debug_path,
            "smtp_debug_include_headers": smtp_debug_include_headers,
            "smtp_debug_include_data": smtp_debug_include_data,
            "smtp_debug_redact_addresses": smtp_debug_redact_addresses,
            "username": (self.username_var.get() or "").strip(),
            "password": self.password_var.get() or "",
            "tenant_id": (self.tenant_id_var.get() or "").strip(),
            "client_id": (self.client_id_var.get() or "").strip(),
            "client_secret": (self.client_secret_var.get() or "").strip(),
            "scope": (self.resource_scope_var.get() or "").strip(),
            "save_token_in_log": bool(self.save_token_in_log_var.get()),
            "from": from_addr,
            "to": to_addr,
            "cc": cc_addr,
            "bcc": bcc_addr,
            "subject": subject,
            "plain_body": plain_body,
            "html_body": html_body,
            "attachments": list(self._attachments or []),
        }

    def _profile_keys(self):
        return [
            "auth_mode",
            "smtp_host",
            "smtp_port",
            "use_starttls",
            "username",
            "password",
            "tenant_id",
            "client_id",
            "client_secret",
            "resource_scope",
            "from",
            "to",
            "cc",
            "bcc",
            "subject",
            "body_plain",
            "body_html",
            "body_mode",
        ]

    def _extract_domain_from_email(self, value: str) -> str:
        value = (value or "").strip()
        if "@" not in value:
            return ""
        domain = value.split("@", 1)[1].strip().lower()
        # Avoid obviously invalid domains
        if not domain or " " in domain or "/" in domain:
            return ""
        return domain

    def _lookup_mx_host(self, domain: str) -> str:
        domain = (domain or "").strip().strip(".").lower()
        if not domain:
            return ""

        # On Windows, launching powershell/nslookup can flash a console window.
        # Use CREATE_NO_WINDOW to keep it silent.
        creationflags = 0
        try:
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        except Exception:
            creationflags = 0

        # Prefer PowerShell's Resolve-DnsName (structured output), fallback to nslookup.
        try:
            ps = (
                "(Resolve-DnsName -Type MX -Name '" + domain + "' -ErrorAction Stop | "
                "Sort-Object -Property Preference | Select-Object -First 1 -ExpandProperty NameExchange)"
            )
            r = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps],
                capture_output=True,
                text=True,
                timeout=2.5,
                creationflags=creationflags,
            )
            if r.returncode == 0:
                host = (r.stdout or "").strip().strip(".")
                if host:
                    return host
        except Exception:
            pass

        try:
            r = subprocess.run(
                ["nslookup", "-type=mx", domain],
                capture_output=True,
                text=True,
                timeout=2.5,
                creationflags=creationflags,
            )
            out = (r.stdout or "") + "\n" + (r.stderr or "")
            # Lines commonly contain: "mail exchanger = <host>"
            for line in out.splitlines():
                if "mail exchanger" in line.lower() and "=" in line:
                    host = line.split("=", 1)[1].strip().strip(".")
                    if host:
                        return host
        except Exception:
            pass

        return ""

    def _kickoff_anonymous_mx_lookup(self, domain: str):
        domain = (domain or "").strip().strip(".").lower()
        if not domain:
            return
        if domain in self._anon_mx_cache:
            return
        if domain in self._anon_mx_inflight:
            return

        self._anon_mx_inflight.add(domain)

        def worker():
            mx = ""
            try:
                mx = (self._lookup_mx_host(domain) or "").strip().strip(".")
            except Exception:
                mx = ""

            def apply_result():
                try:
                    self._anon_mx_inflight.discard(domain)
                    if mx:
                        self._anon_mx_cache[domain] = mx
                    self._log(f"Anonymous MX lookup: {domain} -> {mx or '(not found)'}")

                    # Update host only if still in Anonymous and user hasn't overridden it.
                    if (self.auth_mode_var.get() or "") != "Anonymous":
                        return

                    to_addr = (self.to_var.get() or "").strip()
                    first = re.split(r"[;,]", to_addr, maxsplit=1)[0].strip() if to_addr else ""
                    current_domain = (self._extract_domain_from_email(first) or "").lower()
                    if current_domain != domain:
                        return

                    cloud_name = self.cloud_var.get() or "Global"
                    suffix = EOP_DOMAIN_SUFFIX.get(cloud_name, EOP_DOMAIN_SUFFIX["Global"])
                    fallback = f"{domain.replace('.', '-')}.{suffix}" if domain else ""

                    current_host = (self.smtp_host_var.get() or "").strip().strip(".")
                    if mx and (not current_host or current_host.lower() == fallback.lower()):
                        self.smtp_host_var.set(mx)
                except Exception:
                    pass

            try:
                self.after(0, apply_result)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _compute_eop_host_for_anonymous(self) -> str:
        cloud_name = self.cloud_var.get() or "Global"
        suffix = EOP_DOMAIN_SUFFIX.get(cloud_name, EOP_DOMAIN_SUFFIX["Global"])
        # For inbound EOP delivery (simulating external sender), route based on RECIPIENT domain.
        to_addr = (self.to_var.get() or "").strip()
        # Take first recipient if comma/semicolon-separated.
        first = re.split(r"[;,]", to_addr, maxsplit=1)[0].strip() if to_addr else ""
        domain = self._extract_domain_from_email(first)
        if domain:
            domain_l = domain.lower()
            cached = (self._anon_mx_cache.get(domain_l) or "").strip().strip(".")
            if cached:
                return cached

            # Fast fallback (no DNS lookup): contoso.com -> contoso-com.mail.protection...
            label = domain_l.replace(".", "-")
            fallback = f"{label}.{suffix}"

            # Kick off background MX lookup (no UI blocking).
            self._kickoff_anonymous_mx_lookup(domain_l)
            return fallback
        return ""

    def _apply_anonymous_defaults(self):
        # SMTP relay to EOP is typically port 25.
        self.smtp_port_var.set("25")
        eop_host = self._compute_eop_host_for_anonymous()
        if eop_host:
            self.smtp_host_var.set(eop_host)
        else:
            # Can't derive without a valid To address.
            self.smtp_host_var.set("")

    def _on_to_changed(self):
        # Keep EOP host in sync while user edits recipient.
        if (self.auth_mode_var.get() or "") == "Anonymous":
            eop_host = self._compute_eop_host_for_anonymous()
            if eop_host:
                self.smtp_host_var.set(eop_host)

    def _refresh_attachments_listbox(self):
        if not self._attachments_listbox:
            return
        try:
            self._attachments_listbox.delete(0, tk.END)
            for p in self._attachments:
                self._attachments_listbox.insert(tk.END, p)
        except Exception:
            pass

    def _add_attachments(self):
        try:
            paths = filedialog.askopenfilenames(title="Select attachments")
            if not paths:
                return
            for p in paths:
                if p and p not in self._attachments:
                    self._attachments.append(p)
            self._refresh_attachments_listbox()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _remove_selected_attachment(self):
        if not self._attachments_listbox:
            return
        try:
            sel = list(self._attachments_listbox.curselection())
            if not sel:
                return
            # remove from end to start
            for idx in sorted(sel, reverse=True):
                if 0 <= idx < len(self._attachments):
                    self._attachments.pop(idx)
            self._refresh_attachments_listbox()
        except Exception:
            pass

    def _clear_attachments(self):
        self._attachments = []
        self._refresh_attachments_listbox()

    def _browse_smtp_debug_log(self):
        initial = (self.smtp_debug_log_path_var.get() or "").strip() or documents_dir()
        try:
            filename = filedialog.asksaveasfilename(
                title="Select SMTP Debug Log File",
                initialdir=os.path.dirname(initial) if os.path.splitext(initial)[1] else initial,
                initialfile=os.path.basename(initial) if os.path.splitext(initial)[1] else "smtp_debug.log",
                defaultextension=".log",
                filetypes=[("Log files", "*.log"), ("All files", "*.*")],
            )
            if filename:
                self.smtp_debug_log_path_var.set(filename)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _rt_init_tags(self):
        if self._html_body_text is None:
            return
        base = tkfont.nametofont("TkDefaultFont")
        self._rt_font_normal = tkfont.Font(self._html_body_text, base)
        self._rt_font_bold = tkfont.Font(self._html_body_text, base)
        self._rt_font_bold.configure(weight="bold")
        self._rt_font_italic = tkfont.Font(self._html_body_text, base)
        self._rt_font_italic.configure(slant="italic")

        self._html_body_text.configure(font=self._rt_font_normal)
        self._html_body_text.tag_configure("rt_bold", font=self._rt_font_bold)
        self._html_body_text.tag_configure("rt_italic", font=self._rt_font_italic)
        self._html_body_text.tag_configure("rt_underline", underline=True)

    def _rt_safe_slug(self, value: str) -> str:
        v = (value or "").strip().lower()
        out = []
        for ch in v:
            if ch.isalnum():
                out.append(ch)
            else:
                out.append("_")
        s = "".join(out).strip("_")
        return s or "font"

    def _rt_family_tag(self, family: str) -> str:
        family = (family or "").strip() or "Calibri"
        if family in self._rt_family_cache:
            return self._rt_family_cache[family]
        tag = f"rt_ff_{self._rt_safe_slug(family)}"
        self._rt_family_cache[family] = tag
        self._rt_family_tag_map[tag] = family
        if self._html_body_text is not None:
            fnt = tkfont.Font(self._html_body_text, tkfont.nametofont("TkDefaultFont"))
            try:
                fnt.configure(family=family)
            except Exception:
                pass
            self._html_body_text.tag_configure(tag, font=fnt)
        return tag

    def _rt_size_tag(self, size_pt: int) -> str:
        try:
            size_pt = int(size_pt)
        except Exception:
            size_pt = 11
        size_pt = max(6, min(72, size_pt))
        if size_pt in self._rt_size_cache:
            return self._rt_size_cache[size_pt]
        tag = f"rt_fs_{size_pt}"
        self._rt_size_cache[size_pt] = tag
        if self._html_body_text is not None:
            fnt = tkfont.Font(self._html_body_text, tkfont.nametofont("TkDefaultFont"))
            try:
                fnt.configure(size=size_pt)
            except Exception:
                pass
            self._html_body_text.tag_configure(tag, font=fnt)
        return tag

    def _rt_color_tag(self, color_value: str) -> str:
        c = (color_value or "").strip()
        if not c:
            c = "#000000"
        # normalize common 'rgb(r,g,b)'
        if c.lower().startswith("rgb(") and c.endswith(")"):
            try:
                inner = c[4:-1]
                parts = [int(p.strip()) for p in inner.split(",")]
                if len(parts) == 3:
                    c = f"#{parts[0]:02X}{parts[1]:02X}{parts[2]:02X}"
            except Exception:
                c = "#000000"
        if not c.startswith("#") and len(c) == 6:
            c = "#" + c
        c = c.upper()
        if c in self._rt_color_cache:
            return self._rt_color_cache[c]
        tag = f"rt_color_{c.lstrip('#')}"
        self._rt_color_cache[c] = tag
        if self._html_body_text is not None:
            try:
                self._html_body_text.tag_configure(tag, foreground=c)
            except Exception:
                pass
        return tag

    def _rt_apply_font_family(self, family: str):
        if self._html_body_text is None or not self._rt_has_selection():
            return
        start = self._html_body_text.index(tk.SEL_FIRST)
        end = self._html_body_text.index(tk.SEL_LAST)
        # Remove existing family tags in selection
        for t in self._html_body_text.tag_names():
            if str(t).startswith("rt_ff_"):
                self._html_body_text.tag_remove(t, start, end)
        self._html_body_text.tag_add(self._rt_family_tag(family), start, end)

    def _rt_apply_font_size(self, size_text: str):
        if self._html_body_text is None or not self._rt_has_selection():
            return
        try:
            size_pt = int(float((size_text or "11").strip()))
        except Exception:
            size_pt = 11
        start = self._html_body_text.index(tk.SEL_FIRST)
        end = self._html_body_text.index(tk.SEL_LAST)
        for t in self._html_body_text.tag_names():
            if str(t).startswith("rt_fs_"):
                self._html_body_text.tag_remove(t, start, end)
        self._html_body_text.tag_add(self._rt_size_tag(size_pt), start, end)

    def _rt_choose_color(self):
        if self._html_body_text is None or not self._rt_has_selection():
            return
        rgb, hx = colorchooser.askcolor(title="Choose text color")
        if not hx:
            return
        start = self._html_body_text.index(tk.SEL_FIRST)
        end = self._html_body_text.index(tk.SEL_LAST)
        for t in self._html_body_text.tag_names():
            if str(t).startswith("rt_color_"):
                self._html_body_text.tag_remove(t, start, end)
        self._html_body_text.tag_add(self._rt_color_tag(hx), start, end)

    def _rt_selected_line_range(self):
        if self._html_body_text is None:
            return None
        try:
            start = self._html_body_text.index(tk.SEL_FIRST)
            end = self._html_body_text.index(tk.SEL_LAST)
        except Exception:
            start = self._html_body_text.index(tk.INSERT)
            end = start
        line_start = f"{start.split('.')[0]}.0"
        line_end = f"{end.split('.')[0]}.end"
        return line_start, line_end

    def _rt_bullet_list(self):
        if self._html_body_text is None:
            return
        rng = self._rt_selected_line_range()
        if not rng:
            return
        line_start, line_end = rng
        start_line = int(line_start.split(".")[0])
        end_line = int(line_end.split(".")[0])
        for ln in range(start_line, end_line + 1):
            idx = f"{ln}.0"
            line_text = self._html_body_text.get(idx, f"{ln}.end")
            if line_text.strip() == "":
                continue
            if line_text.lstrip().startswith("• "):
                continue
            self._html_body_text.insert(idx, "• ")

    def _rt_numbered_list(self):
        if self._html_body_text is None:
            return
        rng = self._rt_selected_line_range()
        if not rng:
            return
        line_start, line_end = rng
        start_line = int(line_start.split(".")[0])
        end_line = int(line_end.split(".")[0])
        n = 1
        for ln in range(start_line, end_line + 1):
            idx = f"{ln}.0"
            line_text = self._html_body_text.get(idx, f"{ln}.end")
            if line_text.strip() == "":
                continue
            if re.match(r"^\s*\d+\.\s+", line_text):
                continue
            self._html_body_text.insert(idx, f"{n}. ")
            n += 1

    def _rt_has_selection(self) -> bool:
        try:
            self._html_body_text.index(tk.SEL_FIRST)
            self._html_body_text.index(tk.SEL_LAST)
            return True
        except Exception:
            return False

    def _rt_toggle_tag(self, tag_name: str):
        if self._html_body_text is None:
            return
        if not self._rt_has_selection():
            return
        start = self._html_body_text.index(tk.SEL_FIRST)
        end = self._html_body_text.index(tk.SEL_LAST)

        # If selection already has this tag anywhere, remove it; else add.
        has_any = False
        ranges = self._html_body_text.tag_ranges(tag_name)
        for i in range(0, len(ranges), 2):
            if self._html_body_text.compare(ranges[i], "<", end) and self._html_body_text.compare(ranges[i + 1], ">", start):
                has_any = True
                break
        if has_any:
            self._html_body_text.tag_remove(tag_name, start, end)
        else:
            self._html_body_text.tag_add(tag_name, start, end)

    def _rt_new_link_tag(self, href: str) -> str:
        self._rt_link_counter += 1
        tag = f"rt_link_{self._rt_link_counter}"
        self._rt_link_map[tag] = href
        if self._html_body_text is not None:
            self._html_body_text.tag_configure(tag, foreground="blue", underline=True)
            self._html_body_text.tag_bind(tag, "<Button-1>", lambda _e, t=tag: self._rt_open_link_tag(t))
        return tag

    def _rt_open_link_tag(self, tag: str):
        href = (self._rt_link_map.get(tag) or "").strip()
        if not href:
            return
        try:
            webbrowser.open(href)
        except Exception:
            pass

    def _rt_insert_link(self):
        if self._html_body_text is None:
            return
        if not self._rt_has_selection():
            messagebox.showwarning("Link", "Select text first, then click Link.")
            return
        href = prompt_text(self, "Insert Link", "Enter URL (https://...)")
        if not href:
            return
        href = href.strip()
        if not (href.lower().startswith("http://") or href.lower().startswith("https://")):
            messagebox.showwarning("Link", "URL should start with http:// or https://")
            return
        tag = self._rt_new_link_tag(href)
        start = self._html_body_text.index(tk.SEL_FIRST)
        end = self._html_body_text.index(tk.SEL_LAST)
        self._html_body_text.tag_add(tag, start, end)

    def _rt_clear_formatting(self):
        if self._html_body_text is None:
            return
        if not self._rt_has_selection():
            return
        start = self._html_body_text.index(tk.SEL_FIRST)
        end = self._html_body_text.index(tk.SEL_LAST)
        for t in ["rt_bold", "rt_italic", "rt_underline"]:
            self._html_body_text.tag_remove(t, start, end)
        # Remove font/size/color tags
        for t in self._html_body_text.tag_names():
            t = str(t)
            if t.startswith("rt_ff_") or t.startswith("rt_fs_") or t.startswith("rt_color_"):
                self._html_body_text.tag_remove(t, start, end)
        # Remove any link tags
        for tag in list(self._rt_link_map.keys()):
            self._html_body_text.tag_remove(tag, start, end)

    def _rt_show_html_preview(self):
        html_body = self._rt_export_html().strip()
        win = tk.Toplevel(self)
        win.title("HTML")
        win.geometry("760x520")
        txt = tk.Text(win, wrap="word")
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert(tk.END, html_body)
        txt.configure(state="disabled")

    def _rt_export_html(self) -> str:
        if self._html_body_text is None:
            return ""
        events = self._html_body_text.dump("1.0", "end-1c", tag=True, text=True)

        def open_tag(t: str) -> str:
            if t == "rt_bold":
                return "<b>"
            if t == "rt_italic":
                return "<i>"
            if t == "rt_underline":
                return "<u>"
            if t.startswith("rt_ff_"):
                fam = self._rt_family_tag_map.get(t) or ""
                fam = html.escape(fam, quote=True)
                return f"<span style=\"font-family:{fam};\">"
            if t.startswith("rt_fs_"):
                try:
                    size_pt = int(t.split("_", 2)[2])
                except Exception:
                    size_pt = 11
                return f"<span style=\"font-size:{size_pt}pt;\">"
            if t.startswith("rt_color_"):
                hexv = t.split("_", 1)[1]
                return f"<span style=\"color:#{hexv};\">"
            if t.startswith("rt_link_"):
                href = html.escape(self._rt_link_map.get(t, ""), quote=True)
                return f"<a href=\"{href}\">"
            return ""

        def close_tag(t: str) -> str:
            if t == "rt_bold":
                return "</b>"
            if t == "rt_italic":
                return "</i>"
            if t == "rt_underline":
                return "</u>"
            if t.startswith("rt_ff_") or t.startswith("rt_fs_") or t.startswith("rt_color_"):
                return "</span>"
            if t.startswith("rt_link_"):
                return "</a>"
            return ""

        parts: list[str] = []
        open_stack: list[str] = []

        for kind, value, _index in events:
            if kind == "text":
                chunk = html.escape(value, quote=False)
                chunk = chunk.replace("\n", "<br>\n")
                parts.append(chunk)
            elif kind == "tagon":
                if (
                    value in ("rt_bold", "rt_italic", "rt_underline")
                    or value.startswith("rt_link_")
                    or value.startswith("rt_ff_")
                    or value.startswith("rt_fs_")
                    or value.startswith("rt_color_")
                ):
                    parts.append(open_tag(value))
                    open_stack.append(value)
            elif kind == "tagoff":
                if value in open_stack:
                    # Close in reverse until this tag, then re-open any inner tags.
                    to_reopen = []
                    while open_stack:
                        t = open_stack.pop()
                        parts.append(close_tag(t))
                        if t == value:
                            break
                        to_reopen.append(t)
                    for t in reversed(to_reopen):
                        parts.append(open_tag(t))
                        open_stack.append(t)

        while open_stack:
            parts.append(close_tag(open_stack.pop()))
        return "".join(parts).strip()

    def _rt_import_html(self, html_string: str):
        if self._html_body_text is None:
            return
        self._html_body_text.delete("1.0", tk.END)
        html_string = (html_string or "").strip()
        if not html_string:
            return

        # Reset links for fresh import.
        self._rt_link_map = {}
        self._rt_link_counter = 0
        parser = _BasicHtmlToTkParser(
            self._html_body_text,
            link_tag_factory=self._rt_new_link_tag,
            font_family_tag_factory=self._rt_family_tag,
            font_size_tag_factory=self._rt_size_tag,
            color_tag_factory=self._rt_color_tag,
        )
        try:
            parser.feed(html_string)
        except Exception:
            self._html_body_text.insert(tk.END, html_string)

    def _get_profile_from_ui(self) -> dict:
        plain_body = ""
        html_body = ""
        try:
            if self._plain_body_text is not None:
                plain_body = self._plain_body_text.get("1.0", tk.END).rstrip("\n")
        except Exception:
            plain_body = ""
        try:
            if self._html_body_text is not None:
                html_body = self._rt_export_html().rstrip("\n")
        except Exception:
            html_body = ""

        return {
            "auth_mode": self.auth_mode_var.get(),
            "smtp_host": self.smtp_host_var.get(),
            "smtp_port": int(self.smtp_port_var.get() or 587),
            "use_starttls": bool(self.use_starttls_var.get()),
            "username": self.username_var.get(),
            "password": self.password_var.get(),
            "tenant_id": self.tenant_id_var.get(),
            "client_id": self.client_id_var.get(),
            "client_secret": self.client_secret_var.get(),
            "resource_scope": self.resource_scope_var.get(),
            "from": self.from_var.get(),
            "to": self.to_var.get(),
            "cc": self.cc_var.get(),
            "bcc": self.bcc_var.get(),
            "subject": self.subject_var.get(),
            "body_plain": plain_body,
            "body_html": html_body,
            "body_mode": (self.body_mode_var.get() or "plain"),
        }

    def _apply_profile_to_ui(self, profile: dict, *, cloud: str):
        profile = dict(profile or {})

        self.auth_mode_var.set(profile.get("auth_mode") or self.auth_mode_var.get() or "OAuth2")

        self.smtp_host_var.set(profile.get("smtp_host") or "")
        self.smtp_port_var.set(str(profile.get("smtp_port") or "587"))
        self.use_starttls_var.set(bool(profile.get("use_starttls", True)))

        self.username_var.set(profile.get("username") or "")
        self.password_var.set(profile.get("password") or "")

        self.tenant_id_var.set(profile.get("tenant_id") or "")
        self.client_id_var.set(profile.get("client_id") or "")
        self.client_secret_var.set(profile.get("client_secret") or "")
        self.resource_scope_var.set(profile.get("resource_scope") or "")

        self.from_var.set(profile.get("from") or "")
        self.to_var.set(profile.get("to") or "")
        self.cc_var.set(profile.get("cc") or "")
        self.bcc_var.set(profile.get("bcc") or "")
        self.subject_var.set(profile.get("subject") or "")

        body_plain = profile.get("body_plain")
        body_html = profile.get("body_html")
        legacy_body = profile.get("body")
        if (body_plain is None) and (body_html is None) and (legacy_body is not None):
            body_plain = legacy_body
            body_html = ""

        try:
            if self._plain_body_text is not None:
                self._plain_body_text.delete("1.0", tk.END)
                self._plain_body_text.insert(tk.END, body_plain or "")
        except Exception:
            pass
        try:
            if self._html_body_text is not None:
                # body_html is stored as HTML; render it into the rich editor.
                self._rt_import_html(body_html or "")
        except Exception:
            pass

        self.body_mode_var.set((profile.get("body_mode") or "plain").lower())

        self._sync_cloud_defaults(cloud)

    def _read_config_file(self) -> dict:
        if not os.path.exists(self.config_path):
            return {}
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                return json.load(f) or {}
        except Exception:
            return {}

    def _write_config_file(self, cfg: dict) -> None:
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _migrate_config_to_profiles(self, cfg: dict) -> dict:
        cfg = dict(cfg or {})
        if isinstance(cfg.get("profiles"), dict):
            return cfg

        # Migrate legacy flat config to per-cloud profiles.
        cloud = cfg.get("cloud") or "Global"
        profile = {}
        for k in self._profile_keys():
            if k == "body":
                profile[k] = cfg.get(k, "")
            else:
                profile[k] = cfg.get(k)

        migrated = {
            "cloud": cloud,
            "profiles": {
                cloud: profile,
            },
        }
        # Preserve any global keys that may exist in legacy configs.
        for k in ["smtp_debug_enabled", "smtp_debug_log_path"]:
            if k in cfg:
                migrated[k] = cfg.get(k)
        return migrated

    def _log(self, msg: str):
        log_file_only(msg)
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"

        def _ui_append():
            try:
                self.log_text.configure(state="normal")
                self.log_text.insert(tk.END, line + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state="disabled")
            except Exception:
                pass

        if threading.current_thread() is threading.main_thread():
            _ui_append()
        else:
            try:
                self.after(0, _ui_append)
            except Exception:
                pass

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

    def _sync_cloud_defaults(self, cloud: str):
        # If the current value is empty OR still at the other cloud's default, update it.
        cloud = cloud or "Global"
        default_host = DEFAULT_SMTP_HOST.get(cloud, DEFAULT_SMTP_HOST["Global"])
        default_scope = DEFAULT_SCOPE.get(cloud, DEFAULT_SCOPE["Global"])

        current_host = (self.smtp_host_var.get() or "").strip()
        if (not current_host) or (current_host in DEFAULT_SMTP_HOST.values()):
            self.smtp_host_var.set(default_host)

        current_scope = (self.resource_scope_var.get() or "").strip()
        if (not current_scope) or (current_scope in DEFAULT_SCOPE.values()):
            self.resource_scope_var.set(default_scope)

    def _sync_auth_ui(self):
        mode = (self.auth_mode_var.get() or "").strip()
        if not mode:
            mode = "OAuth2"

        # Persist previous mode's SMTP settings before applying new mode defaults.
        try:
            prev_mode = (self._last_auth_mode or "").strip()
            if prev_mode and prev_mode != mode:
                self._smtp_settings_by_mode[prev_mode] = {
                    "host": (self.smtp_host_var.get() or "").strip(),
                    "port": (self.smtp_port_var.get() or "").strip(),
                    "starttls": bool(self.use_starttls_var.get()),
                }
        except Exception:
            pass

        is_basic = mode == "Basic"
        is_oauth = mode == "OAuth2"

        # Basic fields
        # (We keep them visible; user can ignore based on mode)

        # Token logging option only meaningful for OAuth2
        if not is_oauth:
            self.save_token_in_log_var.set(False)

        if mode == "Anonymous":
            self._apply_anonymous_defaults()
        elif mode in ("Basic", "OAuth2"):
            cloud = (self.cloud_var.get() or "Global").strip() or "Global"
            default_host = DEFAULT_SMTP_HOST.get(cloud, DEFAULT_SMTP_HOST["Global"])

            # Restore last settings for this mode if present, otherwise apply safe defaults.
            saved = None
            try:
                saved = self._smtp_settings_by_mode.get(mode)
            except Exception:
                saved = None

            if saved and (saved.get("host") or saved.get("port")):
                if saved.get("host"):
                    self.smtp_host_var.set(str(saved.get("host") or "").strip())
                if saved.get("port"):
                    self.smtp_port_var.set(str(saved.get("port") or "").strip())
                self.use_starttls_var.set(bool(saved.get("starttls")))
            else:
                # EXO submission defaults
                self.smtp_host_var.set(default_host)
                self.smtp_port_var.set("587")
                self.use_starttls_var.set(True)

        # Track last mode after applying.
        self._last_auth_mode = mode

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
        cfg = self._migrate_config_to_profiles(self._read_config_file())
        cloud = cfg.get("cloud") or self.cloud_var.get() or "Global"

        self.smtp_debug_enabled_var.set(bool(cfg.get("smtp_debug_enabled", False)))
        self.smtp_debug_log_path_var.set((cfg.get("smtp_debug_log_path") or "").strip() or today_smtp_debug_log_path())
        self.smtp_debug_include_headers_var.set(bool(cfg.get("smtp_debug_include_headers", True)))
        self.smtp_debug_include_data_var.set(bool(cfg.get("smtp_debug_include_data", False)))
        self.smtp_debug_redact_addresses_var.set(bool(cfg.get("smtp_debug_redact_addresses", True)))

        self._is_switching_cloud_profile = True
        try:
            self.cloud_var.set(cloud)
        finally:
            self._is_switching_cloud_profile = False

        self._current_profile_cloud = cloud
        profiles = cfg.get("profiles") or {}
        self._apply_profile_to_ui(profiles.get(cloud) or {}, cloud=cloud)

        self._refresh_attachments_listbox()

    def _save_config(self):
        cloud = self.cloud_var.get() or "Global"
        cfg = self._migrate_config_to_profiles(self._read_config_file())
        cfg.setdefault("profiles", {})
        cfg["cloud"] = cloud
        cfg["profiles"][cloud] = self._get_profile_from_ui()
        cfg["smtp_debug_enabled"] = bool(self.smtp_debug_enabled_var.get())
        cfg["smtp_debug_log_path"] = (self.smtp_debug_log_path_var.get() or "").strip()
        cfg["smtp_debug_include_headers"] = bool(self.smtp_debug_include_headers_var.get())
        cfg["smtp_debug_include_data"] = bool(self.smtp_debug_include_data_var.get())
        cfg["smtp_debug_redact_addresses"] = bool(self.smtp_debug_redact_addresses_var.get())
        self._write_config_file(cfg)
        self._log(f"Config saved: {self.config_path}")

    def _on_cloud_changed(self):
        if self._is_switching_cloud_profile:
            return

        new_cloud = self.cloud_var.get() or "Global"
        old_cloud = self._current_profile_cloud or new_cloud
        if new_cloud == old_cloud:
            self._sync_cloud_defaults(new_cloud)
            return

        cfg = self._migrate_config_to_profiles(self._read_config_file())
        cfg.setdefault("profiles", {})

        # Save old profile
        cfg["profiles"][old_cloud] = self._get_profile_from_ui()

        # Switch + load new profile
        cfg["cloud"] = new_cloud
        self._write_config_file(cfg)

        self._current_profile_cloud = new_cloud
        profiles = cfg.get("profiles") or {}
        self._apply_profile_to_ui(profiles.get(new_cloud) or {}, cloud=new_cloud)

    def send_clicked(self):
        try:
            self._save_config()
        except Exception:
            pass

        snapshot = self._collect_send_snapshot()
        t = threading.Thread(target=self._send_worker, kwargs={"snapshot": snapshot, "job_label": "Manual"}, daemon=True)
        t.start()

    def _send_worker(self, *, snapshot: dict | None = None, job_label: str = "Send") -> bool:
        try:
            snapshot = snapshot or {}
            mode = (snapshot.get("mode") or "").strip()
            host = (snapshot.get("host") or "").strip().strip(".")
            port = int(snapshot.get("port") or 587)
            use_starttls = bool(snapshot.get("use_starttls"))

            smtp_debug_enabled = bool(snapshot.get("smtp_debug_enabled"))
            smtp_debug_path = (snapshot.get("smtp_debug_path") or "").strip() or today_smtp_debug_log_path()
            smtp_debug_include_headers = bool(snapshot.get("smtp_debug_include_headers"))
            smtp_debug_include_data = bool(snapshot.get("smtp_debug_include_data"))
            smtp_debug_redact_addresses = bool(snapshot.get("smtp_debug_redact_addresses"))

            msg = EmailMessage()
            msg["From"] = (snapshot.get("from") or "").strip() or (snapshot.get("username") or "").strip()
            msg["To"] = (snapshot.get("to") or "").strip()
            if (snapshot.get("cc") or "").strip():
                msg["Cc"] = (snapshot.get("cc") or "").strip()
            if (snapshot.get("subject") or "").strip():
                msg["Subject"] = (snapshot.get("subject") or "").strip()

            plain_body = snapshot.get("plain_body") or ""
            html_body = snapshot.get("html_body") or ""

            msg.set_content(plain_body or "")
            if (html_body or "").strip():
                msg.add_alternative(html_body, subtype="html")

            to_addrs = []
            for field in [snapshot.get("to") or "", snapshot.get("cc") or "", snapshot.get("bcc") or ""]:
                if field:
                    for a in re.split(r"[;,]", field):
                        a = a.strip()
                        if a:
                            to_addrs.append(a)

            if not msg["To"]:
                raise Exception("To is required")

            if not host:
                if mode == "Anonymous":
                    raise Exception(
                        "SMTP Host is empty. In Anonymous mode, enter a valid To address first so the EOP host can be derived (e.g. user@contoso.com)."
                    )
                raise Exception("SMTP Host is required")

            # Guard against using only the EOP suffix (not a real inbound host)
            if mode == "Anonymous":
                if host.lower() in (s.lower() for s in EOP_DOMAIN_SUFFIX.values()):
                    raise Exception(
                        "Anonymous mode requires a tenant-specific EOP host (e.g. contoso-com.mail.protection.outlook.com). Fill To with a recipient in that domain so it can be auto-derived."
                    )

                # If an async MX lookup is running, we might still be using the fast fallback host.
                to_addr = (snapshot.get("to") or "").strip()
                first = re.split(r"[;,]", to_addr, maxsplit=1)[0].strip() if to_addr else ""
                domain = (self._extract_domain_from_email(first) or "").strip().lower()
                if domain and (domain not in self._anon_mx_cache) and (domain in self._anon_mx_inflight):
                    self._log("Anonymous: MX lookup in progress; current SMTP Host may be a fallback. Wait 1-2s if connect fails.")

            # Attach files (if any)
            attachments = list(snapshot.get("attachments") or [])
            if attachments:
                for path in attachments:
                    if not path:
                        continue
                    try:
                        with open(path, "rb") as f:
                            data = f.read()
                        ctype, _enc = mimetypes.guess_type(path)
                        if not ctype:
                            maintype, subtype = "application", "octet-stream"
                        else:
                            maintype, subtype = ctype.split("/", 1)
                        filename = os.path.basename(path)
                        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
                        self._log(f"Attachment added: {filename} ({len(data)} bytes)")
                    except Exception as ae:
                        raise Exception(f"Failed to attach file: {path}. Error: {ae}")

            self._log(f"Connecting SMTP {host}:{port} (STARTTLS={use_starttls})")
            if smtp_debug_enabled:
                self._log(f"SMTP protocol debug log enabled: {smtp_debug_path}")
                try:
                    os.makedirs(os.path.dirname(smtp_debug_path) or ".", exist_ok=True)
                    with open(smtp_debug_path, "a", encoding="utf-8") as f:
                        ts = time.strftime("%Y-%m-%d %H:%M:%S")
                        cloud_name = (snapshot.get("cloud") or "").strip()
                        f.write(f"[{ts}] ===== SMTP session start =====\n")
                        f.write(f"[{ts}] Host={host} Port={port} STARTTLS={use_starttls} Cloud={cloud_name}\n")

                        if smtp_debug_include_headers:
                            f.write(f"[{ts}] ===== Message headers =====\n")
                            for k, v in msg.items():
                                # Headers only. (Bcc is envelope-only and is not a header.)
                                hv = str(v)
                                if smtp_debug_redact_addresses:
                                    hv = redact_email_addresses(hv)
                                f.write(f"[{ts}] {k}: {hv}\n")
                            f.write(f"[{ts}] ===== End headers =====\n")

                        if not smtp_debug_include_data:
                            f.write(f"[{ts}] NOTE: SMTP DATA is redacted (enable 'Include SMTP DATA' to log message content/attachments)\n")
                except Exception:
                    pass

                server = FileDebugSMTP(
                    host,
                    port,
                    timeout=30,
                    debug_log_path=smtp_debug_path,
                    redact_data=(not smtp_debug_include_data),
                    redact_addresses=smtp_debug_redact_addresses,
                )
                server.set_debuglevel(1)
            else:
                server = smtplib.SMTP(host, port, timeout=30)
            server.ehlo()
            if use_starttls:
                context = ssl.create_default_context()
                server.starttls(context=context)
                server.ehlo()

            if mode == "Anonymous":
                self._log("Auth: Anonymous (no login)")
            elif mode == "Basic":
                user = (snapshot.get("username") or "").strip()
                pwd = snapshot.get("password") or ""
                self._log(f"Auth: Basic user={user}")
                server.login(user, pwd)
            else:
                # XOAUTH2 requires a user identity (mailbox UPN/email).
                user = (snapshot.get("username") or "").strip() or (snapshot.get("from") or "").strip()
                if not user:
                    raise Exception("OAuth2 requires Username (or From) to build XOAUTH2 user=")

                token = self._get_oauth_token(snapshot)
                # Always keep SMTP AUTH logging safe; use the checkbox for the explicit token-acquired line.
                token_preview = redact_authorization("Bearer " + token)

                xoauth2_b64 = build_xoauth2_string(user, token)
                xoauth2_preview = (xoauth2_b64[:16] + "...") if xoauth2_b64 else ""
                self._log(
                    f"Auth: OAuth2 XOAUTH2 user={user} token={token_preview} xoauth2_b64_len={len(xoauth2_b64)} xoauth2_b64_preview={xoauth2_preview}"
                )
                code, resp = server.docmd("AUTH", "XOAUTH2 " + xoauth2_b64)
                if code != 235:
                    raise Exception(f"XOAUTH2 AUTH failed: {code} {resp}")

            to_field = (snapshot.get("to") or "").strip()
            cc_field = (snapshot.get("cc") or "").strip()
            bcc_field = (snapshot.get("bcc") or "").strip()
            self._log(
                f"{job_label}: Sending mail From={msg['From']} To={to_field} Cc={cc_field} Bcc={bcc_field} RcptTo={to_addrs}"
            )
            server.send_message(msg, from_addr=msg["From"], to_addrs=to_addrs)
            server.quit()

            if smtp_debug_enabled:
                try:
                    with open(smtp_debug_path, "a", encoding="utf-8") as f:
                        ts = time.strftime("%Y-%m-%d %H:%M:%S")
                        f.write(f"[{ts}] ===== SMTP session end: success =====\n")
                except Exception:
                    pass
            self._log("Send success")
            # No popups; success is logged in Log tab.
            return True
        except Exception as e:
            msg = str(e)
            # Provide actionable hints for common EXO submission failures.
            if isinstance(e, smtplib.SMTPResponseException):
                code = getattr(e, "smtp_code", None)
                err_bytes = getattr(e, "smtp_error", b"") or b""
                err_text = ""
                try:
                    err_text = err_bytes.decode("utf-8", errors="ignore")
                except Exception:
                    err_text = str(err_bytes)

                if code == 430 and ("STOREDRV" in err_text or "Cannot open mailbox" in err_text):
                    cloud = self.cloud_var.get() or "Global"
                    from_addr = (self.from_var.get() or "").strip()
                    smtp_host = (self.smtp_host_var.get() or "").strip()
                    msg += (
                        "\n\nTroubleshooting tips for 'Cannot open mailbox' (STOREDRV):\n"
                        f"- Verify the FROM mailbox exists and is an Exchange Online mailbox: {from_addr or '<empty>'}\n"
                        "- The mailbox must be in the SAME cloud/tenant you are authenticating against.\n"
                        f"  Current Cloud={cloud}, SMTP Host={smtp_host}\n"
                        "- If FROM is a 21V mailbox, use Cloud='21V (China)' and SMTP host 'smtp.partner.outlook.cn'.\n"
                        "- Ensure the app/service principal has SendAs permission for that mailbox (and any required application access policy).\n"
                        "- If the mailbox is newly created/changed, wait for propagation and retry."
                    )

            self._log(f"ERROR: {msg}")
            self._show_log_tab()
            return False

    def _get_oauth_token(self, snapshot: dict | None = None) -> str:
        snapshot = snapshot or {}
        tenant = (snapshot.get("tenant_id") or (self.tenant_id_var.get() or "")).strip()
        client_id = (snapshot.get("client_id") or (self.client_id_var.get() or "")).strip()
        client_secret = (snapshot.get("client_secret") or (self.client_secret_var.get() or "")).strip()
        scope_or_resource = (snapshot.get("scope") or (self.resource_scope_var.get() or "")).strip()

        if tenant.lower().startswith("http://") or tenant.lower().startswith("https://"):
            # User pasted an authority URL; extract the tenant segment.
            try:
                parsed = urlparse(tenant)
                tenant = (parsed.path or "").strip("/").split("/")[0]
            except Exception:
                pass

        if not tenant or not client_id:
            raise Exception("Tenant ID and Client ID are required for OAuth2")

        cloud_name = (snapshot.get("cloud") or (self.cloud_var.get() or "Global")).strip() or "Global"
        cloud = CLOUDS.get(cloud_name, CLOUDS["Global"])
        authority = f"{cloud.authority_host.rstrip('/')}/{tenant}"
        self._log(f"OAuth2 authority: {authority} (Cloud={cloud.name})")

        try:
            import msal  # lazy import for faster startup
        except Exception:
            msal = None
        if msal is None:
            raise Exception("msal is not installed. Run: pip install -r requirements.txt")

        if not client_secret:
            raise Exception("Client Secret is required for OAuth2 (Client Credentials)")

        # Accept either a single scope or a space-separated scope list.
        scopes = [s for s in scope_or_resource.split() if s]
        if not scopes:
            raise Exception("Scope is required for OAuth2")

        try:
            app = msal.ConfidentialClientApplication(
                client_id=client_id,
                client_credential=client_secret,
                authority=authority,
            )
            result = app.acquire_token_for_client(scopes=scopes)
        except Exception as e:
            msg = str(e)
            if "Unable to get authority configuration" in msg:
                well_known = authority.rstrip("/") + "/v2.0/.well-known/openid-configuration"
                probe_detail = None
                try:
                    req = urllib.request.Request(well_known, headers={"User-Agent": "ExchangeOnlineSmtpSender"})
                    with urllib.request.urlopen(req, timeout=10) as resp:
                        probe_detail = f"HTTP {resp.status}"
                except urllib.error.HTTPError as he:
                    probe_detail = f"HTTPError {he.code} {he.reason}"
                except Exception as pe:
                    probe_detail = str(pe)

                hint = (
                    "\n\nTroubleshooting tips:\n"
                    "- Verify the Tenant ID is correct (GUID or tenant domain).\n"
                    "- If this is a 21V (China) tenant, switch Cloud to '21V (China)' (it cannot authenticate via login.microsoftonline.com).\n"
                    "- If this is a Global tenant but you're on a restricted network/proxy, ensure access to login.microsoftonline.com.\n"
                    f"- Test in browser: {well_known}\n"
                    f"- Well-known probe result: {probe_detail}"
                )
                raise Exception(msg + hint) from e
            raise

        if not isinstance(result, dict) or "access_token" not in result:
            raise Exception(f"Failed to get token: {result}")

        access_token = result["access_token"]
        save_token = bool(snapshot.get("save_token_in_log")) if snapshot else bool(self.save_token_in_log_var.get())
        token_preview = access_token if save_token else redact_authorization("Bearer " + access_token)
        self._log(f"Token acquired (MSAL): {token_preview}")
        return access_token


if __name__ == "__main__":
    enable_windows_dpi_awareness()
    App().mainloop()
