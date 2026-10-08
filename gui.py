"""
gui.py
======
Tkinter-Oberflaeche fuer den Leepa Order & CompMap Finder.

Ermoeglicht die Eingabe eines einzelnen Sensors oder das Einlesen einer
Sensorliste aus einer Excel-Datei, fragt beim Start nach einem Speicherort
und fuehrt die Analyse (core.run_analysis) in einem Hintergrundthread aus,
damit die GUI nicht einfriert.
"""

import os
import queue
import re
import sys
import threading
import getpass
import platform
import json
import tkinter as tk
import traceback
from datetime import datetime

from tkinter import (
    DISABLED, NORMAL, StringVar, Tk, filedialog, messagebox,
)
from tkinter import ttk

import core

# Spalten der Ergebnis-Uebersicht in der GUI (Reihenfolge + Anzeigename)
RESULT_COLUMNS = [
    ("Sensor", "Sensor"),
    ("Leepa_Auftrag", "Leepa-Auftrag"),
    ("Messfahrt", "Messfahrt"),
    ("Abweichung", "Abweichung"),
    ("FE20_MapID", "FE20"),
    ("FE23_MapID", "FE23"),
    ("FE3A_MapID", "FE3A"),
    ("FEA3_MapID", "FEA3"),
    ("Hinweis", "Hinweis"),
]

# Spalten der Sensor-Zusammenfassung (ein Eintrag je Sensor)
SENSOR_SUMMARY_COLUMNS = [
    ("Sensor", "Sensor"),
    ("Leepa_Auftraege", "Leepa-Auftraege"),
    ("FE20_Werte", "FE20 - haeufigste reale Werte"),
    ("FE23_Werte", "FE23 - haeufigste reale Werte"),
    ("FE3A_Werte", "FE3A - haeufigste reale Werte"),
    ("FEA3_Werte", "FEA3 - haeufigste reale Werte"),
]

ORDER_COLUMNS = [
    ("Leepa_Auftrag", "Leepa-Auftrag"),
    ("Anzahl_Sensoren", "Sensoren"),
    ("Anzahl_Messfahrten", "Messfahrten"),
    ("Max_LTC_h", "Max LTC [h]"),
]

SOURCE_OPTIONS = [
    (key, value["label"]) for key, value in core.DATA_SOURCES.items()
]
SOURCE_LABEL_TO_KEY = {label: key for key, label in SOURCE_OPTIONS}

APP_NAME = "SensorMeasureTrace"
APP_VERSION = "1.1.0"
APP_SHORT = "SMT"
APP_SUBTITLE = "Sensor, Order & Compensation Map Trace"
APP_AUTHOR = "by Gabor Herke"
USAGE_LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), APP_NAME)
USAGE_LOG_PATH = os.path.join(USAGE_LOG_DIR, "usage_log.jsonl")
THEME = {
    "bg": "#121a2b",
    "panel": "#1a2440",
    "panel_alt": "#243457",
    "field": "#31456f",
    "border": "#28d7e5",
    "border_soft": "#4f6799",
    "text": "#f7fbff",
    "muted": "#d7e3ff",
    "accent": "#17d4e0",
    "accent_2": "#8b5cf6",
    "danger": "#ff9aa2",
    "selection": "#24345b",
}
INFO_TEXT = (
    "Wo finde ich was?\n\n"
    "Datenquelle:\n"
    "- Oben kann gewaehlt werden, ob in CLNX/clnxSCUDataSet oder in CLMG/clmgSCUDataSet gesucht wird.\n"
    "- Die Wahl gilt fuer Auftragssuche und Map-Suche gemeinsam.\n\n"
    "GUI:\n"
    "- Verlauf: zeigt den aktuellen Suchlauf und Hinweise/Fehler.\n"
    "- Sensor-Zusammenfassung: pro Sensor die gefundenen Leepa-Auftraege und die haeufigsten Map-Werte.\n"
    "- Ergebnis-Uebersicht (Detail): einzelne Zeilen je Sensor und Messfahrt mit Map-ID und Abweichungsmarkierung.\n\n"
    "Excel-Datei:\n"
    "- Blatt INFO: kurze Erklaerung der Blaetter, Map-IDs und der verwendeten Datenquelle.\n"
    "- Blatt Uebersicht: alle Treffer je Sensor und Messfahrt mit FE20/FE23/FE3A/FEA3-Map-ID.\n"
    "- Blatt Leepa_Auftraege: Zuordnung der Sensoren zu Leepa-Auftraegen inkl. Messfahrten/LTC.\n"
    "- Blatt Leepa_unique: jeder gefundene Leepa-Auftrag genau einmal, daneben je gesuchtem Sensor der LTC-Wert.\n"
    "- Blatt Abweichler: alle Faelle, in denen eine Map nicht der haeufigsten Variante entspricht (inkl. Leepa-Auftrag).\n"
    "- Blaetter Rohdaten_<Leepa-Auftrag>: nur vorhanden, wenn oben 'Ausgabe aller FE-Inhalte als Tabelle' angehakt war. "
    "Vollstaendiger Zeilen-Auszug der SCU-Tabelle, ein Blatt je Leepa-Auftrag, je Sensor genau eine Zeile.\n"
    "- Weitere Blaetter wie FE20_A, FE23_B usw.: Detailblaetter je eindeutiger Map. Dort stehen Physical- und Real-Werte sowie, von welchen Sensoren/Messfahrten die Map verwendet wird.\n\n"
    "Map-IDs kurz erklaert:\n"
    "- FE20 / FE23 / FE3A / FEA3: Map-Typ.\n"
    "- Der Buchstabe zeigt die Haeufigkeit im aktuellen Ergebnis: A = haeufigste Variante, B = zweithaeufigste, C = dritthaeufigste usw.\n"
    "- Beispiel: FE20_A bedeutet die im aktuellen Suchlauf am haeufigsten gefundene FE20-Map."
)


def _parse_list_text(text):
    """Zerlegt eingefuegten/eingetippten Text in eine Liste von Eintraegen.
    Unterstuetzt sowohl eine Eintragung pro Zeile (z.B. aus Excel kopierte
    Spalte) als auch komma-/semikolon-/tab-getrennte Werte."""
    if not text:
        return []
    parts = re.split(r"[,;\n\t]+", text)
    return [p.strip() for p in parts if p.strip()]


def resource_path(*parts):
    """Liefert den Pfad zu einer mitgelieferten Ressource - funktioniert
    sowohl im normalen Python-Lauf als auch in der von PyInstaller
    gebauten EXE (dort liegen die Daten in sys._MEIPASS)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, *parts)


def current_exe_path():
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    candidate = sys.argv[0] if sys.argv and sys.argv[0] else __file__
    if candidate in ("-c", "<string>"):
        candidate = __file__
    return os.path.abspath(candidate)


def record_usage_event(event_type, source_key=None):
    os.makedirs(USAGE_LOG_DIR, exist_ok=True)
    payload = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "event": event_type,
        "source_key": source_key or "",
        "user": getpass.getuser(),
        "computer": platform.node(),
        "exe_path": current_exe_path(),
        "exe_name": os.path.basename(current_exe_path()),
    }
    with open(USAGE_LOG_PATH, "a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False) + "\n")


def load_usage_summary():
    if not os.path.exists(USAGE_LOG_PATH):
        return [], None

    rows = []
    try:
        with open(USAGE_LOG_PATH, "r", encoding="utf-8") as handle:
            for line in handle:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except OSError:
        return [], None

    grouped = {}
    for row in rows:
        key = row.get("exe_path") or row.get("exe_name") or "unbekannt"
        bucket = grouped.setdefault(key, {
            "exe_name": row.get("exe_name", "unbekannt"),
            "exe_path": row.get("exe_path", ""),
            "user": row.get("user", ""),
            "computer": row.get("computer", ""),
            "launch_count": 0,
            "last_seen": "",
        })
        bucket["launch_count"] += 1
        bucket["last_seen"] = max(bucket["last_seen"], row.get("timestamp", ""))

    current_path = current_exe_path()
    current = grouped.get(current_path)
    ordered = sorted(grouped.values(), key=lambda item: (item["launch_count"], item["last_seen"]), reverse=True)
    return ordered, current


class SensorCalFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()
        self.root.title(f"{APP_NAME} v{APP_VERSION} - {APP_SUBTITLE}")
        self.root.geometry("760x680")
        self.root.minsize(680, 560)

        self.mode = StringVar(value="single")
        self.single_sensor = StringVar()
        self.source_label = StringVar(value=SOURCE_OPTIONS[0][1])
        self.export_raw_scu_var = tk.BooleanVar(value=False)

        self.msg_queue = queue.Queue()
        self.worker_thread = None
        self._logo_images = []  # Referenzen halten, sonst greift der GC zu frueh
        self._dashboard_window = None
        self._usage_rows = []
        self._usage_current = None

        record_usage_event("launch")
        self._apply_theme()
        self._set_window_icon()
        self._build_ui()
        self._show_start_splash()
        self._poll_queue()

    # ------------------------------------------------------------------
    # Aufbau der Oberflaeche
    # ------------------------------------------------------------------
    def _set_window_icon(self):
        try:
            from PIL import Image, ImageTk

            logo_path = resource_path("resources", "logo.png")
            img = Image.open(logo_path)
            icon_img = ImageTk.PhotoImage(img)
            self._logo_images.append(icon_img)
            self.root.iconphoto(True, icon_img)
        except Exception:
            pass  # Kein Logo vorhanden -> Standard-Icon verwenden

    def _apply_theme(self):
        self.root.configure(bg=THEME["bg"])
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("TFrame", background=THEME["bg"])
        style.configure("Card.TFrame", background=THEME["panel"])
        style.configure("TLabel", background=THEME["bg"], foreground=THEME["text"])
        style.configure("Muted.TLabel", background=THEME["bg"], foreground=THEME["muted"])
        style.configure("Header.TLabel", background=THEME["bg"], foreground=THEME["text"])
        style.configure("TLabelframe", background=THEME["panel"], foreground=THEME["text"], borderwidth=1, bordercolor=THEME["border_soft"])
        style.configure("TLabelframe.Label", background=THEME["panel"], foreground=THEME["text"])
        style.configure("TSeparator", background=THEME["border_soft"])
        style.configure(
            "TButton",
            background=THEME["field"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            focusthickness=0,
            padding=(10, 6),
        )
        style.map(
            "TButton",
            background=[("active", THEME["accent"]), ("disabled", THEME["field"])],
            foreground=[("active", "#0f1730"), ("disabled", THEME["muted"])],
        )
        style.configure("TEntry", fieldbackground=THEME["field"], foreground=THEME["text"], insertcolor=THEME["text"], bordercolor=THEME["border_soft"])
        style.configure("TCombobox", fieldbackground=THEME["field"], foreground=THEME["text"], arrowsize=16, bordercolor=THEME["border_soft"])
        style.map("TCombobox", fieldbackground=[("readonly", THEME["field"])], foreground=[("readonly", THEME["text"])])
        style.configure("TRadiobutton", background=THEME["panel"], foreground=THEME["text"])
        style.map("TRadiobutton", foreground=[("selected", THEME["accent"]), ("active", THEME["text"])])
        style.configure("TCheckbutton", background=THEME["panel"], foreground=THEME["text"])
        style.map("TCheckbutton", foreground=[("selected", THEME["accent"]), ("active", THEME["text"])])
        style.configure("TNotebook", background=THEME["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=THEME["panel_alt"], foreground=THEME["text"], padding=(14, 8))
        style.map("TNotebook.Tab", background=[("selected", THEME["accent_2"]), ("active", THEME["panel"])], foreground=[("selected", THEME["text"]), ("active", THEME["text"])])
        style.configure("Horizontal.TProgressbar", troughcolor=THEME["field"], background=THEME["accent"], bordercolor=THEME["border"], lightcolor=THEME["accent"], darkcolor=THEME["accent"])
        style.configure("Treeview", background=THEME["panel_alt"], fieldbackground=THEME["panel_alt"], foreground=THEME["text"], bordercolor=THEME["border_soft"], rowheight=24)
        style.configure("Treeview.Heading", background=THEME["panel"], foreground=THEME["text"], bordercolor=THEME["border"])
        style.map("Treeview", background=[("selected", THEME["selection"])], foreground=[("selected", THEME["text"])])

    def _center_window(self, window, width, height):
        window.update_idletasks()
        x_pos = max((window.winfo_screenwidth() - width) // 2, 0)
        y_pos = max((window.winfo_screenheight() - height) // 2, 0)
        window.geometry(f"{width}x{height}+{x_pos}+{y_pos}")

    def _show_start_splash(self):
        try:
            from PIL import Image, ImageTk

            splash = tk.Toplevel(self.root)
            splash.overrideredirect(True)
            splash.configure(bg=THEME["bg"])
            self._center_window(splash, 460, 420)

            logo_path = resource_path("resources", "logo.png")
            logo_img = Image.open(logo_path)
            logo_img.thumbnail((260, 260))
            logo_photo = ImageTk.PhotoImage(logo_img)
            self._logo_images.append(logo_photo)

            frame = tk.Frame(splash, bg=THEME["bg"], padx=28, pady=24)
            frame.pack(fill="both", expand=True)
            tk.Label(frame, image=logo_photo, bg=THEME["bg"]).pack(pady=(0, 16))
            tk.Label(
                frame,
                text=f"{APP_NAME} ({APP_SHORT})",
                font=("Segoe UI", 18, "bold"),
                bg=THEME["bg"],
                fg=THEME["text"],
            ).pack()
            tk.Label(
                frame,
                text="Loading...",
                font=("Segoe UI", 10),
                bg=THEME["bg"],
                fg=THEME["muted"],
            ).pack(pady=(8, 0))

            def close_splash():
                if splash.winfo_exists():
                    splash.destroy()
                self.root.deiconify()

            self.root.after(1400, close_splash)
        except Exception:
            self.root.deiconify()

    def _build_ui(self):
        header = ttk.Frame(self.root, padding=(16, 16, 16, 8), style="Card.TFrame")
        header.pack(fill="x")

        try:
            from PIL import Image, ImageTk

            logo_path = resource_path("resources", "logo.png")
            logo_img = Image.open(logo_path)
            logo_img.thumbnail((72, 72))
            logo_photo = ImageTk.PhotoImage(logo_img)
            self._logo_images.append(logo_photo)
            ttk.Label(header, image=logo_photo).pack(side="left", padx=(0, 12))
        except Exception:
            pass

        title_frame = ttk.Frame(header, style="Card.TFrame")
        title_frame.pack(side="left", fill="both", expand=True)
        ttk.Label(title_frame, text=f"{APP_NAME} ({APP_SHORT})", font=("Segoe UI", 16, "bold"), style="Header.TLabel").pack(anchor="w")
        ttk.Label(title_frame, text=APP_SUBTITLE, font=("Segoe UI", 9), style="Muted.TLabel").pack(anchor="w")
        ttk.Label(title_frame, text=APP_AUTHOR, font=("Segoe UI", 8), style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=16)

        body = ttk.Frame(self.root, padding=16)
        body.pack(fill="both", expand=True)

        source_frame = ttk.LabelFrame(body, text="Datenquelle", padding=12)
        source_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(
            source_frame,
            text="In welchen Tabellen soll gesucht werden?",
        ).pack(side="left")
        self.source_combo = ttk.Combobox(
            source_frame,
            textvariable=self.source_label,
            values=[label for _, label in SOURCE_OPTIONS],
            state="readonly",
            width=28,
        )
        self.source_combo.pack(side="left", padx=(10, 0))
        ttk.Checkbutton(
            source_frame,
            text="Ausgabe aller FE-Inhalte als Tabelle (Rohdaten je Leepa-Auftrag)",
            variable=self.export_raw_scu_var,
        ).pack(side="left", padx=(20, 0))

        # -------------------- Eingabe --------------------
        mode_frame = ttk.LabelFrame(body, text="Eingabe", padding=12)
        mode_frame.pack(fill="x", pady=(0, 12))

        # --- Einzelner Sensor ---
        single_row = ttk.Frame(mode_frame)
        single_row.pack(fill="x", pady=2)
        ttk.Radiobutton(
            single_row, text="Einzelner Sensor", variable=self.mode, value="single",
            command=self._on_mode_change,
        ).pack(side="left")
        self.single_entry = ttk.Entry(single_row, textvariable=self.single_sensor, width=30)
        self.single_entry.pack(side="left", padx=(8, 0))

        ttk.Separator(mode_frame, orient="horizontal").pack(fill="x", pady=8)

        # --- Mehrere Sensoren (Liste einfuegen oder Excel-Datei) ---
        multi_header = ttk.Frame(mode_frame)
        multi_header.pack(fill="x")
        ttk.Radiobutton(
            multi_header, text="Mehrere Sensoren", variable=self.mode, value="multi",
            command=self._on_mode_change,
        ).pack(side="left")
        self.excel_button = ttk.Button(
            multi_header, text="...aus Excel-Datei laden", command=self._choose_excel
        )
        self.excel_button.pack(side="left", padx=(8, 0))
        self.excel_label = ttk.Label(multi_header, text="", style="Muted.TLabel")
        self.excel_label.pack(side="left", padx=(8, 0))

        ttk.Label(
            mode_frame,
            text=(
                "Liste hier einfuegen (eine Bezeichnung pro Zeile - z.B. Spalte "
                "aus Excel kopieren und mit Strg+V einfuegen):"
            ),
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 2))
        self.sensor_paste_text = tk.Text(
            mode_frame,
            height=4,
            wrap="none",
            background=THEME["field"],
            foreground=THEME["text"],
            insertbackground=THEME["text"],
            highlightbackground=THEME["border_soft"],
            highlightcolor=THEME["border"],
            relief="flat",
            bd=1,
        )
        self.sensor_paste_text.pack(fill="x")

        ttk.Separator(mode_frame, orient="horizontal").pack(fill="x", pady=8)

        # --- Leepa-Auftrag(e) ---
        ttk.Radiobutton(
            mode_frame, text="Leepa-Auftrag(e)", variable=self.mode, value="auftrag",
            command=self._on_mode_change,
        ).pack(anchor="w")
        ttk.Label(
            mode_frame,
            text=(
                "Liste hier einfuegen (ein Auftrag pro Zeile oder kommagetrennt, "
                "z.B. L2400123):"
            ),
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 2))
        self.auftrag_paste_text = tk.Text(
            mode_frame,
            height=3,
            wrap="none",
            background=THEME["field"],
            foreground=THEME["text"],
            insertbackground=THEME["text"],
            highlightbackground=THEME["border_soft"],
            highlightcolor=THEME["border"],
            relief="flat",
            bd=1,
        )
        self.auftrag_paste_text.pack(fill="x")

        self._on_mode_change()

        # -------------------- Start --------------------
        action_frame = ttk.Frame(body)
        action_frame.pack(fill="x", pady=(0, 12))
        self.start_button = ttk.Button(
            action_frame, text="Datenbank-Suche starten", command=self._on_start
        )
        self.start_button.pack(side="left")
        self.open_button = ttk.Button(
            action_frame, text="Ausgabedatei oeffnen", command=self._open_output, state=DISABLED
        )
        self.open_button.pack(side="left", padx=(8, 0))

        # -------------------- Fortschritt --------------------
        progress_frame = ttk.LabelFrame(body, text="Fortschritt", padding=12)
        progress_frame.pack(fill="x", pady=(0, 12))

        self.status_var = StringVar(value="Bereit.")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(anchor="w")

        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(6, 0))

        # -------------------- Verlauf / Ergebnisse --------------------
        notebook = ttk.Notebook(body)
        notebook.pack(fill="both", expand=True)

        log_frame = ttk.Frame(notebook, padding=8)
        notebook.add(log_frame, text="Verlauf")

        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side="right", fill="y")

        self.log_text = tk.Text(
            log_frame, height=10, wrap="word", yscrollcommand=log_scroll.set,
            state=DISABLED, font=("Consolas", 9), background=THEME["panel_alt"],
            foreground=THEME["text"], insertbackground=THEME["text"], relief="flat", bd=0,
        )
        self.log_text.pack(fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)

        info_frame = ttk.Frame(notebook, padding=12)
        notebook.add(info_frame, text="INFO")

        info_scroll = ttk.Scrollbar(info_frame)
        info_scroll.pack(side="right", fill="y")
        self.info_text = tk.Text(
            info_frame,
            wrap="word",
            yscrollcommand=info_scroll.set,
            state=NORMAL,
            font=("Segoe UI", 10),
            padx=8,
            pady=8,
            bd=0,
            background=THEME["field"],
            foreground=THEME["text"],
            insertbackground=THEME["text"],
        )
        self.info_text.insert("1.0", INFO_TEXT)
        self.info_text.config(state=DISABLED)
        self.info_text.pack(fill="both", expand=True)
        info_scroll.config(command=self.info_text.yview)

        usage_frame = ttk.Frame(notebook, padding=8)
        notebook.add(usage_frame, text="Nutzung")

        usage_header = ttk.Frame(usage_frame)
        usage_header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            usage_header,
            text="Startstatistik der gefundenen EXE-Kopien",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")
        self.usage_summary_label = ttk.Label(
            usage_header,
            text="",
            style="Muted.TLabel",
        )
        self.usage_summary_label.pack(anchor="w")

        usage_scroll_y = ttk.Scrollbar(usage_frame, orient="vertical")
        usage_scroll_y.pack(side="right", fill="y")
        self.usage_tree = ttk.Treeview(
            usage_frame,
            columns=["Exe", "Benutzer", "Rechner", "Starts", "Zuletzt"],
            show="headings",
            yscrollcommand=usage_scroll_y.set,
        )
        for key, label, width in [
            ("Exe", "EXE/Kopie", 220),
            ("Benutzer", "Benutzer", 120),
            ("Rechner", "Rechner", 160),
            ("Starts", "Starts", 80),
            ("Zuletzt", "Zuletzt gesehen", 150),
        ]:
            self.usage_tree.heading(key, text=label)
            self.usage_tree.column(key, width=width, anchor="w")
        self.usage_tree.pack(fill="both", expand=True)
        usage_scroll_y.config(command=self.usage_tree.yview)

        summary_frame = ttk.Frame(notebook, padding=8)
        notebook.add(summary_frame, text="Sensor-Zusammenfassung")

        summary_scroll_y = ttk.Scrollbar(summary_frame, orient="vertical")
        summary_scroll_y.pack(side="right", fill="y")
        summary_scroll_x = ttk.Scrollbar(summary_frame, orient="horizontal")
        summary_scroll_x.pack(side="bottom", fill="x")

        self.sensor_summary_tree = ttk.Treeview(
            summary_frame,
            columns=[key for key, _ in SENSOR_SUMMARY_COLUMNS],
            show="headings",
            yscrollcommand=summary_scroll_y.set,
            xscrollcommand=summary_scroll_x.set,
        )
        for key, label in SENSOR_SUMMARY_COLUMNS:
            self.sensor_summary_tree.heading(key, text=label)
            if key == "Sensor":
                width = 110
            elif key == "Leepa_Auftraege":
                width = 160
            else:
                width = 340
            self.sensor_summary_tree.column(key, width=width, anchor="w")
        self.sensor_summary_tree.pack(fill="both", expand=True)
        summary_scroll_y.config(command=self.sensor_summary_tree.yview)
        summary_scroll_x.config(command=self.sensor_summary_tree.xview)

        results_frame = ttk.Frame(notebook, padding=8)
        notebook.add(results_frame, text="Ergebnis-Uebersicht (Detail)")

        results_scroll_y = ttk.Scrollbar(results_frame, orient="vertical")
        results_scroll_y.pack(side="right", fill="y")
        results_scroll_x = ttk.Scrollbar(results_frame, orient="horizontal")
        results_scroll_x.pack(side="bottom", fill="x")

        self.results_tree = ttk.Treeview(
            results_frame,
            columns=[key for key, _ in RESULT_COLUMNS],
            show="headings",
            yscrollcommand=results_scroll_y.set,
            xscrollcommand=results_scroll_x.set,
        )
        for key, label in RESULT_COLUMNS:
            self.results_tree.heading(key, text=label)
            width = 130 if key in ("Sensor", "Leepa_Auftrag", "Hinweis") else 80
            self.results_tree.column(key, width=width, anchor="w")
        self.results_tree.tag_configure("abweichung", background="#4c2631", foreground=THEME["text"])
        self.results_tree.pack(fill="both", expand=True)
        results_scroll_y.config(command=self.results_tree.yview)
        results_scroll_x.config(command=self.results_tree.xview)

        self._refresh_usage_tab()

        self._output_path = None

    # ------------------------------------------------------------------
    # Eingabe-Handling
    # ------------------------------------------------------------------
    def _on_mode_change(self):
        mode = self.mode.get()
        self.single_entry.config(state=NORMAL if mode == "single" else DISABLED)
        self.excel_button.config(state=NORMAL if mode == "multi" else DISABLED)
        self.sensor_paste_text.config(state=NORMAL if mode == "multi" else DISABLED)
        self.auftrag_paste_text.config(state=NORMAL if mode == "auftrag" else DISABLED)

    def _choose_excel(self):
        path = filedialog.askopenfilename(
            title="Bezeichnungsliste (Excel) waehlen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")],
        )
        if not path:
            return
        try:
            sensors = core.read_sensors_from_excel(path)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Excel-Datei konnte nicht gelesen werden:\n{exc}")
            return

        if not sensors:
            messagebox.showwarning(APP_NAME, "In der Excel-Datei wurden keine Bezeichnungen gefunden.")
            return

        self.sensor_paste_text.delete("1.0", "end")
        self.sensor_paste_text.insert("1.0", "\n".join(sensors))
        self.excel_label.config(
            text=f"{os.path.basename(path)}  ({len(sensors)} Sensor(en) geladen)"
        )

    def _collect_input(self):
        """Liefert ein Dict {"sensor_list": [...], "auftrag_list": [...]} je nach
        gewaehltem Eingabemodus, oder None wenn die Eingabe ungueltig ist
        (in diesem Fall wurde bereits eine Warnung angezeigt)."""
        mode = self.mode.get()

        if mode == "single":
            sensor = self.single_sensor.get().strip()
            if not sensor:
                messagebox.showwarning(APP_NAME, "Bitte eine Bezeichnung eingeben.")
                return None
            return {"sensor_list": [sensor], "auftrag_list": None}

        if mode == "multi":
            raw = self.sensor_paste_text.get("1.0", "end")
            sensors = _parse_list_text(raw)
            if not sensors:
                messagebox.showwarning(
                    APP_NAME,
                    "Bitte Bezeichnungen einfuegen oder eine Excel-Datei laden.",
                )
                return None
            return {"sensor_list": sensors, "auftrag_list": None}

        # mode == "auftrag"
        raw = self.auftrag_paste_text.get("1.0", "end")
        auftraege = _parse_list_text(raw)
        if not auftraege:
            messagebox.showwarning(APP_NAME, "Bitte mindestens einen Leepa-Auftrag eingeben.")
            return None
        return {"sensor_list": None, "auftrag_list": auftraege}

    # ------------------------------------------------------------------
    # Start / Hintergrundverarbeitung
    # ------------------------------------------------------------------
    def _on_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        collected = self._collect_input()
        if not collected:
            return

        default_name = f"Leepa_Order_CompMap_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Ausgabedatei speichern unter",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel-Datei", "*.xlsx")],
        )
        if not output_path:
            return

        self._output_path = output_path
        self.open_button.config(state=DISABLED)
        self.start_button.config(state=DISABLED)
        self.progress.config(value=0, maximum=1)
        self._log_clear()
        self._clear_results()
        self._clear_sensor_summary()
        self._close_result_windows()
        self._log("Starte Datenbank-Suche...")
        self.status_var.set("Datenbank-Suche laeuft...")

        self.worker_thread = threading.Thread(
            target=self._worker,
            args=(
                collected["sensor_list"],
                collected["auftrag_list"],
                output_path,
                self._get_selected_source_key(),
                self.export_raw_scu_var.get(),
            ),
            daemon=True,
        )
        self.worker_thread.start()

    def _get_selected_source_key(self):
        return SOURCE_LABEL_TO_KEY.get(self.source_label.get(), "CLNX")

    def _worker(self, sensor_list, auftrag_list, output_path, source_key, export_raw_scu):
        def progress_callback(idx, total, sensor):
            self.msg_queue.put(("progress", idx, total, sensor))

        try:
            record_usage_event("analysis", source_key=source_key)
            summary = core.run_analysis(
                sensor_list=sensor_list,
                auftrag_list=auftrag_list,
                output_path=output_path,
                progress_callback=progress_callback,
                source_key=source_key,
                export_raw_scu=export_raw_scu,
            )
            self.msg_queue.put(("done", summary))
        except Exception:
            self.msg_queue.put(("error", traceback.format_exc()))

    def _poll_queue(self):
        try:
            while True:
                message = self.msg_queue.get_nowait()
                self._handle_message(message)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_message(self, message):
        kind = message[0]

        if kind == "progress":
            _, idx, total, sensor = message
            self.progress.config(value=idx, maximum=max(total, 1))
            self.status_var.set(f"Verarbeite Sensor {idx}/{total}: {sensor}")
            self._log(f"[{idx}/{total}] {sensor}")

        elif kind == "done":
            summary = message[1]
            self.status_var.set("Fertig.")
            self.progress.config(value=self.progress["maximum"])
            self._log("Datenbank-Suche abgeschlossen.")
            self._log(
                f"  Sensoren: {summary['sensoren']}, Zeilen Uebersicht: {summary['zeilen_uebersicht']}, "
                f"Abweichler: {summary['abweichler']}"
            )
            self._log(f"  Datenquelle: {summary.get('source_label', '-')}")
            if self.export_raw_scu_var.get():
                self._log(f"  Rohdaten-Blaetter (je Leepa-Auftrag): {summary.get('raw_export_sheets', 0)}")
            self._fill_results(summary.get("overview_records", []))
            self._fill_sensor_summary(summary.get("sensor_summary_records", []))
            self._show_dashboard_window(summary)
            self._refresh_usage_tab()
            self.start_button.config(state=NORMAL)
            self.open_button.config(state=NORMAL)

            hinweis = ""
            fehlende_auftraege = summary.get("auftraege_ohne_sensoren") or []
            if fehlende_auftraege:
                hinweis = (
                    "\n\nHinweis: Zu folgenden Auftraegen wurden keine Sensoren gefunden:\n"
                    + ", ".join(fehlende_auftraege)
                )
                self._log("Hinweis: keine Sensoren gefunden fuer Auftrag(e): " + ", ".join(fehlende_auftraege))

            messagebox.showinfo(
                APP_NAME,
                f"Datenbank-Suche abgeschlossen.\n\nDatei gespeichert unter:\n{self._output_path}{hinweis}",
            )

        elif kind == "error":
            self.status_var.set("Fehler.")
            self._log("FEHLER:")
            self._log(message[1])
            self.start_button.config(state=NORMAL)
            messagebox.showerror(APP_NAME, f"Bei der Datenbank-Suche ist ein Fehler aufgetreten:\n\n{message[1]}")

    def _open_output(self):
        if not self._output_path:
            messagebox.showwarning(APP_NAME, "Es wurde noch keine Ausgabedatei erzeugt.")
            return
        if not os.path.exists(self._output_path):
            messagebox.showwarning(
                APP_NAME,
                f"Die Ausgabedatei wurde nicht gefunden:\n{self._output_path}",
            )
            return
        try:
            os.startfile(self._output_path)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Datei konnte nicht geoeffnet werden:\n{exc}")

    # ------------------------------------------------------------------
    # Ergebnis-Uebersicht (Treeview)
    # ------------------------------------------------------------------
    def _clear_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

    def _fill_results(self, records):
        self._clear_results()
        for record in records:
            values = []
            for key, _ in RESULT_COLUMNS:
                value = record.get(key, "")
                if key == "Abweichung":
                    value = "Ja" if value else ""
                values.append("" if value is None else value)
            tags = ("abweichung",) if record.get("Abweichung") else ()
            self.results_tree.insert("", "end", values=values, tags=tags)

    def _clear_sensor_summary(self):
        for item in self.sensor_summary_tree.get_children():
            self.sensor_summary_tree.delete(item)

    def _fill_sensor_summary(self, records):
        self._clear_sensor_summary()
        for record in records:
            values = [record.get(key, "") for key, _ in SENSOR_SUMMARY_COLUMNS]
            self.sensor_summary_tree.insert("", "end", values=values)

    def _close_result_windows(self):
        if self._dashboard_window is not None and self._dashboard_window.winfo_exists():
            self._dashboard_window.destroy()
        self._dashboard_window = None

    def _build_map_card(self, parent, title, variant, empty_text=None):
        card = ttk.LabelFrame(parent, text=title, padding=10)

        if variant is None:
            ttk.Label(
                card,
                text=empty_text or "Keine Daten gefunden.",
                style="Muted.TLabel",
            ).pack(anchor="w")
            return card

        meta = (
            f"Map-ID: {variant['map_id']}\n"
            f"Treffer: {variant['anzahl_verwendungen']}\n"
            f"Sensoren: {variant['anzahl_sensoren']}"
        )
        ttk.Label(card, text=meta, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            card,
            text="Realwerte der Map:",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(0, 6))

        text = tk.Text(
            card,
            height=8,
            wrap="none",
            font=("Consolas", 9),
            bd=0,
            padx=8,
            pady=8,
            background=THEME["field"],
            foreground=THEME["text"],
            insertbackground=THEME["text"],
        )
        text.insert("1.0", variant["real_pretty"])
        text.config(state=DISABLED)
        text.pack(fill="both", expand=True)
        return card

    def _show_dashboard_window(self, summary):
        if self._dashboard_window is not None and self._dashboard_window.winfo_exists():
            self._dashboard_window.destroy()

        variant_summary = summary.get("map_variant_summary", {})
        records = summary.get("auftrag_summary_records", [])

        window = tk.Toplevel(self.root)
        window.title(f"{APP_NAME} - Ergebnis-Dashboard")
        self._center_window(window, 1540, 860)
        self._dashboard_window = window

        outer = ttk.Frame(window, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="Ergebnis-Dashboard",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(0, 6))
        ttk.Label(
            outer,
            text=(
                f"Datenquelle: {summary.get('source_label', '-')}  |  "
                "oben Leepa-Auftraege, unten die haeufigste und zweithaeufigste Map-Variante je Typ."
            ),
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(0, 10))

        orders_frame = ttk.LabelFrame(outer, text="Gefundene Leepa-Auftraege", padding=10)
        orders_frame.pack(fill="x", pady=(0, 10))

        tree_frame = ttk.Frame(orders_frame)
        tree_frame.pack(fill="both", expand=True)
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        tree = ttk.Treeview(
            tree_frame,
            columns=[key for key, _ in ORDER_COLUMNS],
            show="headings",
            yscrollcommand=scroll_y.set,
            height=7,
        )
        for key, label in ORDER_COLUMNS:
            tree.heading(key, text=label)
            width = 190 if key == "Leepa_Auftrag" else 130
            tree.column(key, width=width, anchor="w")
        for record in records:
            values = []
            for key, _ in ORDER_COLUMNS:
                value = record.get(key, "")
                if key == "Max_LTC_h" and value not in (None, ""):
                    value = f"{float(value):.0f}"
                values.append("" if value is None else value)
            tree.insert("", "end", values=values)
        tree.pack(fill="x", expand=True)
        scroll_y.config(command=tree.yview)

        grid = ttk.Frame(outer)
        grid.pack(fill="both", expand=True)
        map_names = ["FE20", "FE23", "FE3A", "FEA3"]
        for col_idx in range(len(map_names)):
            grid.columnconfigure(col_idx, weight=1)
        for row_idx in range(2):
            grid.rowconfigure(row_idx, weight=1)

        labels = {1: "Haeufigste Variante", 2: "Zweithaeufigste Variante"}
        for col_idx, map_name in enumerate(map_names):
            variants = variant_summary.get(map_name, [])
            by_rank = {item.get("rank"): item for item in variants}
            for row_idx, rank in enumerate((1, 2)):
                title = f"{map_name} - {labels[rank]}"
                empty_text = "Keine Daten gefunden." if rank == 1 else "Keine zweite Variante gefunden."
                card = self._build_map_card(grid, title, by_rank.get(rank), empty_text)
                card.grid(row=row_idx, column=col_idx, sticky="nsew", padx=6, pady=6)

    def _refresh_usage_tab(self):
        self.usage_tree.delete(*self.usage_tree.get_children())
        rows, current = load_usage_summary()
        self._usage_rows = rows
        self._usage_current = current
        current_count = current["launch_count"] if current else 0
        self.usage_summary_label.config(
            text=(
                f"Aktuelle EXE: {current['exe_path'] if current else current_exe_path()} | "
                f"Starts dieser Kopie: {current_count} | "
                f"Gesamt bekannte Kopien: {len(rows)}"
            )
        )
        for row in rows:
            self.usage_tree.insert(
                "",
                "end",
                values=(
                    row.get("exe_path", row.get("exe_name", "")),
                    row.get("user", ""),
                    row.get("computer", ""),
                    row.get("launch_count", 0),
                    row.get("last_seen", ""),
                ),
            )

    # ------------------------------------------------------------------
    # Log-Hilfsfunktionen
    # ------------------------------------------------------------------
    def _log(self, text):
        self.log_text.config(state=NORMAL)
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.config(state=DISABLED)

    def _log_clear(self):
        self.log_text.config(state=NORMAL)
        self.log_text.delete("1.0", "end")
        self.log_text.config(state=DISABLED)


def main():
    root = Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    app = SensorCalFinderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
