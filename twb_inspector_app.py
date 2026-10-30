"""
TWB Inspector - desktop app.  Run:  python twb_inspector_app.py

Drop one or more Tableau workbooks (.twb / .twbx) on the window and get:
  * an Excel audit (data sources & tables, all fields, unused fields,
    duplicate calculations, review flags, dependencies, worksheets)
  * an interactive HTML lineage diagram (works offline)
plus a short summary of what is worth a look, right in the window.

The analysis lives in the twb_inspector package; this file only builds the window.

Needs: customtkinter (modern look). Optional: tkinterdnd2 (drag and drop).
"""

import json
import os
import queue
import subprocess
import sys
import threading
import webbrowser
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ImportError:  # pragma: no cover - shown to people running from source
    sys.exit("TWB Inspector needs customtkinter. Install it with:  pip install -r requirements.txt")

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

from twb_inspector import APP_NAME, __version__ as APP_VERSION
from twb_inspector.core import STATUS_UNUSED
from twb_inspector.report import process_workbook, resource_path

# Colours (shared with the icon and the lineage diagram)
NAVY = "#14213D"
ACCENT = "#0B6BCB"
ACCENT_HOVER = "#0A5DB0"
BG = "#F6F8FB"
CARD = "#FFFFFF"
BORDER = "#E3E8EF"
MUTED = "#6B7785"
TEXT = "#1F2933"
WARN_BG, WARN_BORDER, WARN_TEXT = "#FFF5F5", "#F5C2C7", "#C53030"
OK_TEXT, RUN_TEXT, WAIT_TEXT = "#2F855A", "#B7791F", "#9AA5B1"

STATUS_ICON = {"waiting": ("○", WAIT_TEXT), "running": ("●", RUN_TEXT),
               "done": ("✔", OK_TEXT), "failed": ("✘", WARN_TEXT)}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def _settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, APP_NAME, "settings.json")


def load_settings() -> dict:
    try:
        with open(_settings_path(), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_settings(data: dict) -> None:
    try:
        os.makedirs(os.path.dirname(_settings_path()), exist_ok=True)
        with open(_settings_path(), "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
    except Exception:
        pass  # settings are a convenience only


def default_output_dir() -> str:
    base = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "outputs")


def open_path(path: str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 - local file chosen by the user
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def short_path(path: str, keep: int = 2, max_len: int = 38) -> str:
    parts = os.path.normpath(path).split(os.sep)
    text = path if len(parts) <= keep + 1 else "…" + os.sep + os.sep.join(parts[-keep:])
    if len(text) > max_len:
        text = "…" + os.sep + parts[-1]
    return text if len(text) <= max_len else "…" + text[-(max_len - 1):]


def _soft_lower(text: str) -> str:
    """'Long formula' -> 'long formula' but keep acronyms such as 'LOD expression' or 'COUNTD'."""
    first = text.split(" ")[0]
    return text if first.isupper() else text[:1].lower() + text[1:]


def findings(lineage, limit: int = 8):
    """Most useful things to look at first, as (bold part, rest) pairs."""
    fields = lineage.fields.values()
    out = []
    unused_calcs = sorted((f for f in fields if f.kind == "Calculated" and f.status == STATUS_UNUSED
                           and f.defined_in != "Published data source"),
                          key=lambda f: (-len(f.flags), f.name.lower()))
    for f in unused_calcs[:2]:
        extra = [_soft_lower(fl.split(" (")[0]) for fl in f.flags if not fl.startswith("Same formula")]
        out.append((f.name, "unused" + (", " + ", ".join(extra[:2]) if extra else "")))
    if len(unused_calcs) > 2:
        out.append((f"{len(unused_calcs) - 2} more unused calculations", "see the Unused sheet"))
    from_pub = [f for f in fields if f.kind == "Calculated" and f.status == STATUS_UNUSED
                and f.defined_in == "Published data source"]
    if from_pub:
        out.append((f"{len(from_pub)} unused in the published source", "other workbooks may use them"))
    params = [f for f in fields if f.kind == "Parameter" and f.status == STATUS_UNUSED]
    if params:
        out.append((f"{len(params)} unused parameter{'s' if len(params) > 1 else ''}",
                    ", ".join(p.name for p in params[:3])))
    for gid, keys in list(lineage.duplicate_groups.items())[:2]:
        names = [lineage.fields[k].name for k in keys]
        out.append(("Same formula", " = ".join(names[:3]) + (" …" if len(names) > 3 else "")))
    raw = [f for f in fields if f.kind == "Raw" and f.status == STATUS_UNUSED and not f.hidden]
    if raw:
        out.append((f"{len(raw)} column{'s' if len(raw) > 1 else ''} never used and not hidden",
                    "hide them to tidy the Data pane" if len(raw) > 1 else "hide it to tidy the Data pane"))
    hidden = [f for f in fields if f.hidden]
    if hidden:
        hidden_unused = sum(1 for f in hidden if f.status == STATUS_UNUSED)
        out.append((f"{len(hidden)} hidden field{'s' if len(hidden) > 1 else ''}",
                    f"{hidden_unused} unused, {len(hidden) - hidden_unused} still used in the workbook"))
    sql = [t for t in lineage.tables if t.table_type == "Custom SQL"]
    if sql:
        out.append((f"{len(sql)} Custom SQL quer{'ies' if len(sql) > 1 else 'y'}", ", ".join(t.name for t in sql[:3])))
    deep = [f for f in fields if f.kind == "Calculated" and any(fl.startswith("Deep") for fl in f.flags)]
    if deep:
        out.append((f"{len(deep)} deep calculation chain{'s' if len(deep) > 1 else ''}", deep[0].name))
    if not out:
        out.append(("Nothing stands out", "no unused calculations, duplicates or Custom SQL found"))
    return out[:limit]


# --------------------------------------------------------------------------- #
# Window
# --------------------------------------------------------------------------- #

_Base = (type("AppBase", (ctk.CTk, TkinterDnD.DnDWrapper), {}) if HAS_DND else ctk.CTk)


class TWBInspector(_Base):
    def __init__(self):
        super().__init__()
        if HAS_DND:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self.dnd_ready = True
            except Exception:
                self.dnd_ready = False
        else:
            self.dnd_ready = False

        ctk.set_appearance_mode("light")
        self.title(APP_NAME)
        self.geometry("1080x760")
        self.minsize(940, 680)
        self.configure(fg_color=BG)
        self._set_icon()

        self.settings = load_settings()
        self.files = []            # list of dicts: path, status, row widgets
        self.results = {}          # path -> result dict or error string
        self.queue = queue.Queue()
        self.running = False

        self.f_title = ctk.CTkFont(family="Segoe UI", size=20, weight="bold")
        self.f_sub = ctk.CTkFont(family="Segoe UI", size=13)
        self.f_h3 = ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
        self.f_body = ctk.CTkFont(family="Segoe UI", size=13)
        self.f_bold = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        self.f_small = ctk.CTkFont(family="Segoe UI", size=12)
        self.f_big = ctk.CTkFont(family="Segoe UI", size=24, weight="bold")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_header()
        self._build_body()
        self._build_footer()
        self._show_placeholder()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------- layout
    def _set_icon(self):
        ico = resource_path(os.path.join("assets", "twb_inspector.ico"))
        png = resource_path(os.path.join("assets", "twb_inspector.png"))

        def apply():
            try:
                if sys.platform.startswith("win") and os.path.exists(ico):
                    self.iconbitmap(ico)
                elif os.path.exists(png):
                    from PIL import Image, ImageTk
                    self._icon_img = ImageTk.PhotoImage(Image.open(png))
                    self.iconphoto(True, self._icon_img)
            except Exception:
                pass
        apply()
        # customtkinter sets its own icon shortly after start; set ours again after that
        self.after(300, apply)

    def _card(self, parent, title):
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=10, border_width=1, border_color=BORDER)
        ctk.CTkLabel(card, text=title.upper(), font=self.f_h3, text_color=MUTED).pack(anchor="w", padx=16, pady=(14, 8))
        return card

    def _build_header(self):
        hd = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0, height=86)
        hd.grid(row=0, column=0, sticky="ew")
        hd.grid_columnconfigure(1, weight=1)
        try:
            from PIL import Image
            img = Image.open(resource_path(os.path.join("assets", "twb_inspector_header.png")))
            self._hd_img = ctk.CTkImage(light_image=img, dark_image=img, size=(48, 48))
            ctk.CTkLabel(hd, text="", image=self._hd_img).grid(row=0, column=0, rowspan=2, padx=(22, 14), pady=18)
        except Exception:
            ctk.CTkLabel(hd, text="").grid(row=0, column=0, rowspan=2, padx=(22, 0))
        ctk.CTkLabel(hd, text=APP_NAME, font=self.f_title, text_color="white").grid(row=0, column=1, sticky="sw", pady=(18, 0))
        ctk.CTkLabel(hd, text="Unused fields, lineage and data sources for any .twb / .twbx",
                     font=self.f_sub, text_color="#B8C4DC").grid(row=1, column=1, sticky="nw", pady=(0, 18))
        ctk.CTkLabel(hd, text=f" v{APP_VERSION} ", font=self.f_small, text_color="#8FA0C0", fg_color="#22335A",
                     corner_radius=10).grid(row=0, column=2, rowspan=2, padx=22)

    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=22, pady=(18, 8))
        body.grid_columnconfigure(0, weight=0, minsize=386)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="transparent", width=370)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        left.grid_propagate(False)          # keep a fixed width, whatever the results contain
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)

        # 1 · Workbooks
        wb = self._card(left, "1 · Workbooks")
        wb.grid(row=0, column=0, sticky="nsew")
        drop = ctk.CTkFrame(wb, fg_color="#F9FBFF", corner_radius=8, border_width=2, border_color="#9FB3D1", height=96)
        drop.pack(fill="x", padx=16)
        self.drop_label = ctk.CTkLabel(drop, text="Drag workbooks here" if self.dnd_ready else "Add your workbooks",
                                       font=self.f_bold, text_color=TEXT)
        self.drop_label.pack(pady=(14, 4))
        row = ctk.CTkFrame(drop, fg_color="transparent")
        row.pack(pady=(0, 14))
        if self.dnd_ready:
            ctk.CTkLabel(row, text="or", font=self.f_body, text_color=MUTED).pack(side="left", padx=(0, 8))
        self.browse_btn = ctk.CTkButton(row, text="Browse…", width=100, command=self.browse_input, fg_color=CARD,
                                        text_color=TEXT, border_width=1, border_color="#C9D3E0", hover_color="#EEF2F7")
        self.browse_btn.pack(side="left")
        self.file_list = ctk.CTkScrollableFrame(wb, fg_color=CARD, height=150)
        self.file_list.pack(fill="both", expand=True, padx=10, pady=(8, 6))
        self.empty_files = ctk.CTkLabel(self.file_list, text="No workbooks yet", font=self.f_small, text_color=WAIT_TEXT)
        self.empty_files.pack(anchor="w", padx=6, pady=6)
        self.clear_btn = ctk.CTkButton(wb, text="Clear list", width=80, height=24, command=self.clear_files,
                                       fg_color="transparent", text_color=MUTED, hover_color="#EEF2F7", font=self.f_small)
        self.clear_btn.pack(anchor="e", padx=12, pady=(0, 10))

        if self.dnd_ready:
            for w in (self, drop, self.drop_label):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

        # 2 · Output
        out = self._card(left, "2 · Output")
        out.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.excel_var = ctk.BooleanVar(value=self.settings.get("excel", True))
        self.html_var = ctk.BooleanVar(value=self.settings.get("html", True))
        self.open_var = ctk.BooleanVar(value=self.settings.get("open", True))
        for text, var in (("Excel audit", self.excel_var), ("Lineage diagram", self.html_var),
                          ("Open diagram when done", self.open_var)):
            ctk.CTkCheckBox(out, text=text, variable=var, font=self.f_body, text_color=TEXT,
                            fg_color=ACCENT, hover_color=ACCENT_HOVER, checkbox_width=18,
                            checkbox_height=18).pack(anchor="w", padx=16, pady=3)
        self.output_dir = self.settings.get("output_dir") or default_output_dir()
        path_row = ctk.CTkFrame(out, fg_color="#F3F5F8", corner_radius=6, border_width=1, border_color=BORDER)
        path_row.pack(fill="x", padx=16, pady=(8, 14))
        ctk.CTkButton(path_row, text="Change", width=60, height=24, command=self.browse_output_dir,
                      fg_color="transparent", text_color=ACCENT, hover_color="#E6EEF8",
                      font=self.f_small).pack(side="right", padx=4)
        self.path_label = ctk.CTkLabel(path_row, text="Folder:  " + short_path(self.output_dir), font=self.f_small,
                                       text_color="#4A5A70", anchor="w")
        self.path_label.pack(side="left", padx=8, pady=4, fill="x", expand=True)

        # Analyze
        run = ctk.CTkFrame(left, fg_color="transparent")
        run.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        self.run_btn = ctk.CTkButton(run, text="Analyze", height=44, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
                                     fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self.process_workbooks)
        self.run_btn.pack(fill="x")
        self.progress = ctk.CTkProgressBar(run, height=6, progress_color=ACCENT, fg_color=BORDER)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(10, 2))
        self.status_label = ctk.CTkLabel(run, text="Add one or more workbooks to start.", font=self.f_small,
                                         text_color=MUTED, anchor="w")
        self.status_label.pack(fill="x")

        # Results
        self.results_card = ctk.CTkFrame(body, fg_color=CARD, corner_radius=10, border_width=1, border_color=BORDER)
        self.results_card.grid(row=0, column=1, sticky="nsew")
        self.results_card.grid_columnconfigure(0, weight=1)
        self.results_card.grid_rowconfigure(0, weight=1)
        self.tabs = None
        self._update_run_button()

    def _build_footer(self):
        ft = ctk.CTkFrame(self, fg_color="transparent")
        ft.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 10))
        ctk.CTkLabel(ft, text="Runs on your computer. Workbooks are never changed or uploaded.", font=self.f_small,
                     text_color=MUTED).pack(side="left")
        ctk.CTkButton(ft, text="About", width=50, height=22, fg_color="transparent", text_color=MUTED,
                      hover_color="#E6EAF0", font=self.f_small, command=self.show_about).pack(side="right")

    def _clear_results(self):
        for w in self.results_card.winfo_children():
            w.destroy()
        self.tabs = None

    def _show_placeholder(self):
        self._clear_results()
        box = ctk.CTkFrame(self.results_card, fg_color="transparent")
        box.grid(row=0, column=0)
        ctk.CTkLabel(box, text="Results appear here", font=self.f_title, text_color=TEXT).pack(pady=(0, 10))
        for line in ("1.  Add .twb or .twbx workbooks on the left",
                     "2.  Choose what to create and where",
                     "3.  Click Analyze. You get a summary here, an Excel audit",
                     "     and an interactive lineage diagram"):
            ctk.CTkLabel(box, text=line, font=self.f_body, text_color=MUTED, anchor="w", justify="left").pack(anchor="w")

    # ------------------------------------------------------------- inputs
    def _add_files(self, paths):
        added = 0
        known = {f["path"] for f in self.files}
        for p in paths:
            p = p.strip().strip("{}")
            if os.path.isdir(p):
                inner = [os.path.join(p, n) for n in sorted(os.listdir(p)) if n.lower().endswith((".twb", ".twbx"))]
                self._add_files(inner)
                continue
            if not p.lower().endswith((".twb", ".twbx")) or p in known or not os.path.isfile(p):
                continue
            self.files.append({"path": p, "status": "waiting"})
            known.add(p)
            added += 1
        self._render_files()
        self._update_run_button()
        if added:
            self.status_label.configure(text=f"{len(self.files)} workbook(s) ready.")
        return added

    def _on_drop(self, event):
        if self.running:
            return
        paths = self.tk.splitlist(event.data)
        if not self._add_files(paths):
            self.status_label.configure(text="Only .twb and .twbx files can be added.")

    def browse_input(self):
        names = filedialog.askopenfilenames(
            title="Select Tableau workbooks",
            initialdir=self.settings.get("last_input_dir") or os.path.expanduser("~"),
            filetypes=[("Tableau workbooks", "*.twb *.twbx"), ("All files", "*")])
        if names:
            self.settings["last_input_dir"] = os.path.dirname(names[0])
            self._add_files(names)

    def clear_files(self):
        if self.running:
            return
        self.files = []
        self._render_files()
        self._update_run_button()
        self.status_label.configure(text="Add one or more workbooks to start.")

    def _remove_file(self, path):
        if self.running:
            return
        self.files = [f for f in self.files if f["path"] != path]
        self._render_files()
        self._update_run_button()

    def _render_files(self):
        for w in self.file_list.winfo_children():
            w.destroy()
        if not self.files:
            self.empty_files = ctk.CTkLabel(self.file_list, text="No workbooks yet", font=self.f_small, text_color=WAIT_TEXT)
            self.empty_files.pack(anchor="w", padx=6, pady=6)
            return
        for f in self.files:
            row = ctk.CTkFrame(self.file_list, fg_color="transparent")
            row.pack(fill="x", pady=1)
            sym, col = STATUS_ICON[f["status"]]
            f["icon"] = ctk.CTkLabel(row, text=sym, width=18, font=self.f_bold, text_color=col)
            f["icon"].pack(side="left", padx=(4, 6))
            ctk.CTkLabel(row, text=os.path.basename(f["path"]), font=self.f_body, text_color=TEXT,
                         anchor="w").pack(side="left", fill="x", expand=True)
            ctk.CTkButton(row, text="✕", width=24, height=22, fg_color="transparent", text_color=WAIT_TEXT,
                          hover_color="#EEF2F7", command=lambda p=f["path"]: self._remove_file(p)).pack(side="right")

    def _set_file_status(self, path, status):
        for f in self.files:
            if f["path"] == path:
                f["status"] = status
                sym, col = STATUS_ICON[status]
                if f.get("icon"):
                    f["icon"].configure(text=sym, text_color=col)

    def _update_run_button(self):
        n = len(self.files)
        self.run_btn.configure(text=f"Analyze {n} workbook{'s' if n != 1 else ''}" if n else "Analyze",
                               state="normal" if n and not self.running else "disabled")

    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="Select output folder", initialdir=self.output_dir)
        if directory:
            self.output_dir = directory
            self.path_label.configure(text="Folder:  " + short_path(directory))

    # ------------------------------------------------------------- processing
    def process_workbooks(self):
        if not self.files or self.running:
            return
        if not (self.excel_var.get() or self.html_var.get()):
            messagebox.showerror(APP_NAME, "Tick Excel audit and/or Lineage diagram.")
            return
        self.running = True
        self.results = {}
        self._persist()
        for f in self.files:
            self._set_file_status(f["path"], "waiting")
        self.progress.set(0)
        self._update_run_button()
        self.browse_btn.configure(state="disabled")
        self._clear_results()
        self.tabs = ctk.CTkTabview(self.results_card, fg_color=CARD, segmented_button_selected_color=NAVY,
                                   segmented_button_selected_hover_color=NAVY,
                                   segmented_button_unselected_color="#8492AB",
                                   segmented_button_unselected_hover_color="#6F7E99",
                                   segmented_button_fg_color="#8492AB", text_color="#FFFFFF")
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=10, pady=(4, 10))
        self._tab_names = set()
        self._tab_order = []
        args = ([f["path"] for f in self.files], self.output_dir, self.excel_var.get(), self.html_var.get())
        threading.Thread(target=self._worker, args=args, daemon=True).start()
        self.after(100, self._poll)

    def _worker(self, files, output_dir, excel, html):
        for i, path in enumerate(files, start=1):
            self.queue.put(("start", path, i, len(files)))
            try:
                self.queue.put(("done", path, process_workbook(path, output_dir, excel=excel, html=html)))
            except Exception as exc:
                self.queue.put(("error", path, f"{type(exc).__name__}: {exc}"))
        self.queue.put(("finished",))

    def _poll(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "start":
                    _, path, i, n = msg
                    self._set_file_status(path, "running")
                    self.status_label.configure(text=f"Analyzing {os.path.basename(path)} ({i} of {n})…")
                    self.progress.set((i - 1) / n)
                elif msg[0] == "done":
                    _, path, res = msg
                    self.results[path] = res
                    self._set_file_status(path, "done")
                    self._add_result_tab(path, res)
                elif msg[0] == "error":
                    _, path, err = msg
                    self.results[path] = err
                    self._set_file_status(path, "failed")
                    self._add_error_tab(path, err)
                elif msg[0] == "finished":
                    self._finish()
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _finish(self):
        self.running = False
        self.progress.set(1)
        self.browse_btn.configure(state="normal")
        self._update_run_button()
        ok = sum(1 for r in self.results.values() if isinstance(r, dict))
        failed = len(self.results) - ok
        self.status_label.configure(text=f"Done: {ok} analyzed" + (f", {failed} failed" if failed else "") + ".")
        # Select one tab, once: customtkinter hides tabs ~100 ms after each switch,
        # so switching repeatedly in quick succession can leave every tab hidden.
        first_ok = next((lbl for p, lbl in self._tab_order if isinstance(self.results.get(p), dict)),
                        self._tab_order[0][1] if self._tab_order else None)
        if first_ok:
            self.after(150, lambda: self.tabs.set(first_ok))
        if self.open_var.get():
            for r in self.results.values():
                if isinstance(r, dict) and r.get("html"):
                    webbrowser.open("file:///" + os.path.realpath(r["html"]).replace("\\", "/"))

    # ------------------------------------------------------------- results
    def _new_tab(self, path):
        name = os.path.splitext(os.path.basename(path))[0]
        label = name if len(name) <= 22 else name[:20] + "…"
        base, i = label, 2
        while label in self._tab_names:
            label = f"{base} ({i})"
            i += 1
        self._tab_names.add(label)
        tab = self.tabs.add(label)
        frame = ctk.CTkFrame(tab, fg_color=CARD)
        frame.pack(fill="both", expand=True, padx=6, pady=(4, 0))
        self._tab_order.append((path, label))
        return frame

    def _add_result_tab(self, path, res):
        L = res["lineage"]
        s = L.summary()
        fr = self._new_tab(path)
        ctk.CTkLabel(fr, text=os.path.basename(path), font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
                     text_color=TEXT, anchor="w").pack(anchor="w", padx=6)
        version = (L.tableau_version or "").split(" (")[0]
        pub = s.get("Published data sources", 0)
        sub = (f"{s['Datasources']} data source{'s' if s['Datasources'] != 1 else ''}"
               + (f" ({pub} published)" if pub else "") + f" · {s['Tables']} table"
               f"{'s' if s['Tables'] != 1 else ''} · {s['Worksheets']} worksheets · {s['Dashboards']} dashboard"
               f"{'s' if s['Dashboards'] != 1 else ''}" + (f" · Tableau {version}" if version else ""))
        ctk.CTkLabel(fr, text=sub, font=self.f_small, text_color=MUTED, anchor="w").pack(anchor="w", padx=6, pady=(0, 12))

        tiles = ctk.CTkFrame(fr, fg_color="transparent")
        tiles.pack(fill="x", padx=4)
        for i, (label, value, warn) in enumerate([
                ("Calculations", s["Calculated fields"], False),
                ("Unused (workbook)" if s.get("Published data sources") else "Unused calcs",
                 s["Unused workbook calculations"], True),
                ("Duplicates", s["Duplicate calculation groups"], True),
                ("Review flags", s["Calculations with flags"], False)]):
            tiles.grid_columnconfigure(i, weight=1, uniform="t")
            hot = warn and value > 0
            t = ctk.CTkFrame(tiles, fg_color=WARN_BG if hot else CARD, corner_radius=8, border_width=1,
                             border_color=WARN_BORDER if hot else BORDER)
            t.grid(row=0, column=i, sticky="ew", padx=4)
            ctk.CTkLabel(t, text=str(value), font=self.f_big, text_color=WARN_TEXT if hot else TEXT,
                         anchor="w").pack(anchor="w", padx=12, pady=(8, 0))
            ctk.CTkLabel(t, text=label, font=self.f_small, text_color=MUTED, anchor="w").pack(anchor="w", padx=12, pady=(0, 8))

        ctk.CTkLabel(fr, text="WORTH A LOOK", font=self.f_h3, text_color=MUTED).pack(anchor="w", padx=6, pady=(16, 6))
        for bold, rest in findings(L):
            line = ctk.CTkFrame(fr, fg_color="transparent")
            line.pack(anchor="w", fill="x", padx=6, pady=1)
            ctk.CTkLabel(line, text="•  " + bold, font=self.f_bold, text_color=TEXT).pack(side="left")
            ctk.CTkLabel(line, text=": " + rest, font=self.f_body, text_color=TEXT, anchor="w",
                         wraplength=420, justify="left").pack(side="left", fill="x")
        for w in L.warnings:
            ctk.CTkLabel(fr, text="⚠ " + w, font=self.f_small, text_color=RUN_TEXT, anchor="w").pack(anchor="w", padx=6)

        acts = ctk.CTkFrame(fr, fg_color="transparent")
        acts.pack(anchor="w", padx=6, pady=(16, 6))
        if res.get("html"):
            ctk.CTkButton(acts, text="Open diagram", width=130, height=34, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                          command=lambda p=res["html"]: webbrowser.open("file:///" + os.path.realpath(p).replace("\\", "/"))
                          ).pack(side="left", padx=(0, 8))
        if res.get("excel"):
            ctk.CTkButton(acts, text="Open Excel", width=110, height=34, fg_color=CARD, text_color=TEXT, border_width=1,
                          border_color="#C9D3E0", hover_color="#EEF2F7",
                          command=lambda p=res["excel"]: open_path(p)).pack(side="left", padx=(0, 8))
        out_dir = os.path.dirname(res.get("excel") or res.get("html") or self.output_dir)
        ctk.CTkButton(acts, text="Open folder", width=110, height=34, fg_color=CARD, text_color=TEXT, border_width=1,
                      border_color="#C9D3E0", hover_color="#EEF2F7",
                      command=lambda p=out_dir: open_path(p)).pack(side="left")

    def _add_error_tab(self, path, err):
        fr = self._new_tab(path)
        ctk.CTkLabel(fr, text=os.path.basename(path), font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
                     text_color=TEXT).pack(anchor="w", padx=6)
        ctk.CTkLabel(fr, text="This workbook could not be analyzed.", font=self.f_bold,
                     text_color=WARN_TEXT).pack(anchor="w", padx=6, pady=(10, 4))
        ctk.CTkLabel(fr, text=err, font=self.f_small, text_color=TEXT, wraplength=520, justify="left",
                     anchor="w").pack(anchor="w", padx=6)
        ctk.CTkLabel(fr, text="Check that the file opens in Tableau and is a .twb or .twbx.", font=self.f_small,
                     text_color=MUTED).pack(anchor="w", padx=6, pady=(8, 0))

    # ------------------------------------------------------------- misc
    def show_about(self):
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME} {APP_VERSION}\n\n"
            "Finds unused fields, duplicate calculations and lineage in Tableau workbooks, "
            "and lists the tables each data source reads.\n\n"
            "Everything runs on this computer. Workbooks are read, never changed or uploaded.\n\n"
            "Free and open source:\ngithub.com/Saina2405/tableau-twb-inspector")

    def _persist(self):
        self.settings.update({"excel": self.excel_var.get(), "html": self.html_var.get(),
                              "open": self.open_var.get(), "output_dir": self.output_dir})
        save_settings(self.settings)

    def _on_close(self):
        self._persist()
        self.destroy()


def main():
    if sys.platform.startswith("win"):
        try:  # show our icon (not Python's) in the taskbar when run as a script
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("TWBInspector.App")
        except Exception:
            pass
    app = TWBInspector()
    app.mainloop()


if __name__ == "__main__":
    main()
