# -*- coding: utf-8 -*-
"""
Excel Ultimate Search - search many Excel / CSV / Word / PDF files for a list of terms at once.
Python / PyQt6 rewrite of the original JavaFX tool.

Double-click this file on Windows: it opens without a console window, installs what it needs
on the first start (behind a splash screen) and looks for a newer version on GitHub at every start.
Each release is named Excel_Ultimate_Search_v<version>.pyw.

    python Excel_Ultimate_Search.pyw --console    keep the console, to read errors
    python Excel_Ultimate_Search.pyw --no-update  skip the update check once

github.com/nikkpap/Excel_Ultimate_Search  ·  Nikolaos K. Paridis (nikkpap) · ALU DEV TEAM
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "Excel Ultimate Search"
APP_VERSION = "3.0.0"
APP_FILE_STEM = "Excel_Ultimate_Search"
HERE = Path(__file__).resolve().parent
# pip name -> module to look for
PACKAGES = ["PyQt6", "openpyxl", "xlrd", "rapidfuzz", "PyMuPDF"]
MODULES = {"PyQt6": "PyQt6.QtWidgets", "openpyxl": "openpyxl", "xlrd": "xlrd",
           "rapidfuzz": "rapidfuzz", "PyMuPDF": "pymupdf"}
# The search runs in worker processes, which load this file again under the name "__mp_main__".
# They must not restart themselves, install anything or open windows.
IS_WORKER = __name__ == "__mp_main__"

# pythonw.exe has no console: sys.stdout / sys.stderr are None there. Give prints somewhere to go.
class _Null:
    def write(self, *_):
        return 0

    def flush(self):
        pass


if sys.stdout is None:
    sys.stdout = _Null()
if sys.stderr is None:
    sys.stderr = _Null()

NO_WINDOW = 0x08000000 if os.name == "nt" else 0   # CREATE_NO_WINDOW
DETACHED = 0x00000008 if os.name == "nt" else 0    # DETACHED_PROCESS


def _gui_python() -> str:
    """pythonw.exe next to the running interpreter: same Python, no console window."""
    exe = Path(sys.executable)
    if os.name == "nt" and exe.name.lower() in ("python.exe", "py.exe"):
        w = exe.with_name("pythonw.exe")
        if w.exists():
            return str(w)
    return str(exe)


def _console_python() -> str:
    """python.exe (pip needs a real interpreter, not pythonw)."""
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        c = exe.with_name("python.exe")
        if c.exists():
            return str(c)
    return str(exe)


# Never show a console: started with python.exe, restart with pythonw.exe and leave.
if not IS_WORKER and __name__ == "__main__" and os.name == "nt" and "--console" not in sys.argv \
        and Path(sys.executable).name.lower() == "python.exe" and _gui_python() != sys.executable:
    import subprocess as _sp
    _sp.Popen([_gui_python(), str(Path(__file__).resolve())] + sys.argv[1:],
              creationflags=DETACHED | NO_WINDOW, close_fds=True)
    sys.exit(0)


def _missing_packages() -> list[str]:
    import importlib
    import importlib.util
    importlib.invalidate_caches()
    out = []
    for pkg in PACKAGES:
        try:
            if importlib.util.find_spec(MODULES[pkg]) is None:
                out.append(pkg)
        except (ImportError, ValueError):
            out.append(pkg)
    return out


def _add_user_site():
    import importlib
    import site
    try:
        p = site.getusersitepackages()
        if p and p not in sys.path:
            sys.path.append(p)
    except Exception:
        pass
    importlib.invalidate_caches()


def _pip_cmd(pkgs: list[str]) -> list[str]:
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)   # --user is refused inside a venv
    return [_console_python(), "-m", "pip", "install", "--disable-pip-version-check", "--upgrade"] + \
        ([] if in_venv else ["--user"]) + pkgs


def _ensure_qt():
    """First start only: PyQt6 is needed to draw anything, so everything is installed behind a Tk splash."""
    try:
        import PyQt6.QtWidgets  # noqa: F401
        return
    except ImportError:
        pass
    import subprocess
    import threading
    try:
        import tkinter
        from tkinter import messagebox, ttk
    except Exception:
        msg = f"{APP_NAME} needs PyQt6. In a terminal run:\n\n  pip install " + " ".join(PACKAGES)
        sys.stderr.write(msg + "\n")
        if os.name == "nt":
            try:
                import ctypes
                ctypes.windll.user32.MessageBoxW(None, msg, APP_NAME, 0x10)
            except Exception:
                pass
        sys.exit(1)
    root = tkinter.Tk()
    root.overrideredirect(True)
    w, h = 480, 200
    root.geometry(f"{w}x{h}+{(root.winfo_screenwidth() - w) // 2}+{(root.winfo_screenheight() - h) // 2}")
    root.configure(bg="#0f1115")
    tkinter.Label(root, text=APP_NAME, bg="#0f1115", fg="#e6e8ee",
                  font=("Segoe UI", 18, "bold")).pack(pady=(28, 2))
    tkinter.Label(root, text="First start: installing what the app needs (one time only)", bg="#0f1115",
                  fg="#8b93a7", font=("Segoe UI", 10)).pack()
    style = ttk.Style(root)
    style.theme_use("default")
    style.configure("E.Horizontal.TProgressbar", troughcolor="#1e222b", background="#3b82f6", thickness=6,
                    bordercolor="#0f1115", lightcolor="#3b82f6", darkcolor="#3b82f6")
    bar = ttk.Progressbar(root, style="E.Horizontal.TProgressbar", mode="indeterminate", length=400)
    bar.pack(pady=(22, 8))
    bar.start(12)
    line = tkinter.Label(root, text="Starting pip…", bg="#0f1115", fg="#5d636f", font=("Segoe UI", 9))
    line.pack()
    state = {"rc": None, "out": [], "last": ""}

    def work():
        try:
            p = subprocess.Popen(_pip_cmd(PACKAGES), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                 text=True, creationflags=NO_WINDOW)
            for ln in p.stdout:
                state["out"].append(ln)
                state["last"] = ln.strip()[:72]
            state["rc"] = p.wait()
        except Exception as e:  # noqa: BLE001
            state["out"].append(str(e))
            state["rc"] = 1

    threading.Thread(target=work, daemon=True).start()

    def poll():
        if state["last"]:
            line.config(text=state["last"])
        if state["rc"] is None:
            root.after(120, poll)
        else:
            root.quit()

    root.after(120, poll)
    root.mainloop()
    root.destroy()
    if state["rc"] != 0:
        r = tkinter.Tk()
        r.withdraw()
        messagebox.showerror(APP_NAME, "Installing PyQt6 failed. Is the internet on?\n\n" +
                             "".join(state["out"])[-1200:])
        sys.exit(1)
    # start again in a fresh process, so the new packages load cleanly
    subprocess.Popen([_gui_python(), str(Path(__file__).resolve())] + sys.argv[1:],
                     creationflags=DETACHED | NO_WINDOW)
    sys.exit(0)


if not IS_WORKER:
    _ensure_qt()
os.environ["QT_LOGGING_RULES"] = ";".join(filter(None, [os.environ.get("QT_LOGGING_RULES", ""),
                                                        "qt.qpa.windowclass.warning=false"]))

# ------------------------------------------------------------------------------------------------
import csv
import datetime as _dt
import fnmatch
import hashlib
import json
import pickle
import re
import subprocess
import threading
import time
import zlib
from dataclasses import dataclass, field

from PyQt6.QtCore import (
    QAbstractTableModel, QFileSystemWatcher, QModelIndex, QObject, QPointF, QRectF, QSettings,
    QSortFilterProxyModel, Qt, QThread, QTimer, QUrl, pyqtSignal,
)
from PyQt6.QtGui import (
    QAction, QBrush, QColor, QDesktopServices, QFont, QFontMetrics, QGuiApplication, QIcon,
    QKeySequence, QLinearGradient, QPainter, QPainterPath, QPen, QPixmap, QRadialGradient, QShortcut,
)
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog, QFileDialog,
    QFrame, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QListWidget,
    QListWidgetItem, QMainWindow, QMenu, QMessageBox, QPlainTextEdit,
    QProgressBar, QPushButton, QScrollArea, QSizePolicy, QSpinBox, QSplitter, QStyle,
    QStyledItemDelegate, QStyleOptionViewItem, QTableView, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

LANG = "en"   # "el" for Greek; set from the settings in main()


def tr(s: str) -> str:
    """Greek text when the interface is in Greek (EL is at the end of the file)."""
    return EL.get(s, s) if LANG == "el" else s


SHEET_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".csv"}
DOC_EXTS = {".docx"}
PDF_EXTS = {".pdf"}
ALL_EXTS = SHEET_EXTS | DOC_EXTS | PDF_EXTS
SETTINGS_ORG, SETTINGS_APP = "nikkpap", "ExcelUltimateSearch"


# --------------------------------------------------------------------------- #
#  Theme
# --------------------------------------------------------------------------- #

THEMES = {
    "dark": dict(
        bg="#0f1115", surface="#171a21", surface2="#1e222b", border="#2a2f3a",
        text="#e6e8ee", muted="#8b93a7", accent="#3b82f6", accent_hover="#5b9bff",
        accent_text="#ffffff", hl_bg="#facc15", hl_text="#111111",
        danger="#f87171", ok="#34d399", sel="#26406b", row_alt="#1a1e26",
        hl_fuzzy="#fb923c", warn="#f59e0b",
    ),
    "light": dict(
        bg="#f4f6fa", surface="#ffffff", surface2="#f0f2f7", border="#dde1ea",
        text="#1b1f2a", muted="#667085", accent="#2563eb", accent_hover="#1d4ed8",
        accent_text="#ffffff", hl_bg="#fde047", hl_text="#111111",
        danger="#dc2626", ok="#059669", sel="#cfe0ff", row_alt="#f8f9fc",
        hl_fuzzy="#fdba74", warn="#b45309",
    ),
}


def _icon_files(t: dict) -> dict:
    """Write tiny SVGs (checkmark, chevron) to a temp dir so QSS can use them."""
    import tempfile
    d = Path(tempfile.gettempdir()) / "excel_ultimate_search_icons"
    d.mkdir(exist_ok=True)
    files = {
        "check": f'<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16">'
                 f'<path d="M3.5 8.5l3 3 6-7" fill="none" stroke="{t["accent_text"]}" stroke-width="2.2" '
                 f'stroke-linecap="round" stroke-linejoin="round"/></svg>',
        "chevron": f'<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12">'
                   f'<path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="{t["muted"]}" stroke-width="1.8" '
                   f'stroke-linecap="round" stroke-linejoin="round"/></svg>',
    }
    out = {}
    for name, svg in files.items():
        f = d / f"{name}_{t['accent_text'][1:]}_{t['muted'][1:]}.svg"
        if not f.exists():
            f.write_text(svg, encoding="utf-8")
        out[name] = f.as_posix()
    return out


def build_qss(t: dict) -> str:
    ic = _icon_files(t)
    return f"""
    * {{ font-size: 10pt; }}
    QMainWindow, QWidget#root {{ background: {t['bg']}; color: {t['text']}; }}
    QWidget {{ color: {t['text']}; }}
    QDialog, QMessageBox {{ background: {t['bg']}; }}
    QLabel a {{ color: {t['accent']}; }}
    QFrame#card {{
        background: {t['surface']}; border: 1px solid {t['border']}; border-radius: 12px;
    }}
    QLabel#title {{ font-size: 15pt; font-weight: 700; }}
    QLabel#subtitle {{ color: {t['muted']}; }}
    QLabel#section {{
        color: {t['muted']}; font-size: 8.5pt; font-weight: 700;
    }}
    QLabel#stat {{
        background: {t['surface2']}; border: 1px solid {t['border']};
        border-radius: 10px; padding: 4px 10px;
    }}
    QLabel#statWarn {{
        background: {t['surface2']}; border: 1px solid {t['border']};
        border-radius: 10px; padding: 4px 10px; color: {t['danger']};
    }}
    QLabel#dropHint {{ color: {t['muted']}; }}

    QPushButton {{
        background: {t['surface2']}; border: 1px solid {t['border']};
        border-radius: 8px; padding: 7px 14px;
    }}
    QPushButton:hover {{ border-color: {t['accent']}; }}
    QPushButton:pressed {{ background: {t['border']}; }}
    QPushButton:disabled {{ color: {t['muted']}; }}
    QPushButton#primary {{
        background: {t['accent']}; color: {t['accent_text']}; border: none;
        font-weight: 700; padding: 10px 16px; font-size: 11pt;
    }}
    QPushButton#primary:hover {{ background: {t['accent_hover']}; }}
    QPushButton#primary:disabled {{ background: {t['border']}; color: {t['muted']}; }}
    QPushButton#danger {{ color: {t['danger']}; }}
    QPushButton#ghost {{ background: transparent; border: none; padding: 6px 8px; }}
    QPushButton#ghost:hover {{ background: {t['surface2']}; }}

    QLineEdit, QPlainTextEdit, QComboBox {{
        background: {t['surface2']}; border: 1px solid {t['border']};
        border-radius: 8px; padding: 6px 8px;
        selection-background-color: {t['accent']}; selection-color: {t['accent_text']};
    }}
    QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus {{ border-color: {t['accent']}; }}
    QComboBox::drop-down {{ border: none; width: 26px; }}
    QComboBox::down-arrow {{ image: url("{ic['chevron']}"); width: 12px; height: 12px; }}
    QComboBox QAbstractItemView {{
        background: {t['surface']}; border: 1px solid {t['border']};
        selection-background-color: {t['sel']}; selection-color: {t['text']};
    }}

    QListWidget {{
        background: {t['surface2']}; border: 1px dashed {t['border']};
        border-radius: 10px; padding: 4px; outline: none;
    }}
    QListWidget#dragOver {{ border: 2px dashed {t['accent']}; }}
    QListWidget::item {{ padding: 5px 6px; border-radius: 6px; }}
    QListWidget::item:selected {{ background: {t['sel']}; color: {t['text']}; }}
    QListWidget::item:hover {{ background: {t['border']}; }}

    QCheckBox {{ spacing: 8px; }}
    QCheckBox::indicator {{
        width: 16px; height: 16px; border-radius: 4px;
        border: 1px solid {t['border']}; background: {t['surface2']};
    }}
    QCheckBox::indicator:checked {{
        background: {t['accent']}; border-color: {t['accent']}; image: url("{ic['check']}");
    }}
    QCheckBox::indicator:hover {{ border-color: {t['accent']}; }}

    QTableView {{
        background: {t['surface']}; alternate-background-color: {t['row_alt']};
        border: none; gridline-color: transparent; outline: none;
        selection-background-color: {t['sel']}; selection-color: {t['text']};
    }}
    QTableView::item {{ padding: 4px 6px; border: none; }}
    QHeaderView::section {{
        background: {t['surface']}; color: {t['muted']}; border: none;
        border-bottom: 1px solid {t['border']}; padding: 8px 6px;
        font-weight: 700; font-size: 9pt;
    }}
    QTableCornerButton::section {{ background: {t['surface']}; border: none; }}

    QProgressBar {{
        background: {t['surface2']}; border: none; border-radius: 3px; max-height: 6px;
    }}
    QProgressBar::chunk {{ background: {t['accent']}; border-radius: 3px; }}

    QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
    QScrollBar::handle:vertical {{ background: {t['border']}; border-radius: 4px; min-height: 30px; }}
    QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
    QScrollBar::handle:horizontal {{ background: {t['border']}; border-radius: 4px; min-width: 30px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
    QScrollBar::add-page, QScrollBar::sub-page {{ background: none; }}

    QSplitter::handle {{ background: transparent; width: 8px; height: 8px; }}
    QScrollArea {{ background: transparent; border: none; }}
    QWidget#scrollContent {{ background: transparent; }}
    QPushButton::menu-indicator {{ image: none; width: 0; }}
    QSpinBox {{
        background: {t['surface2']}; border: 1px solid {t['border']};
        border-radius: 8px; padding: 5px 6px;
    }}
    QSpinBox:disabled, QLineEdit:disabled {{ color: {t['muted']}; }}
    QFrame#preview {{
        background: {t['surface2']}; border: 1px solid {t['border']}; border-radius: 10px;
    }}
    QFrame#preview QTableWidget {{ background: {t['surface2']}; alternate-background-color: {t['surface']}; }}
    QLabel#hint {{ color: {t['muted']}; font-size: 8.5pt; }}
    QMenu {{
        background: {t['surface']}; border: 1px solid {t['border']}; padding: 4px;
        border-radius: 8px;
    }}
    QMenu::item {{ padding: 6px 18px; border-radius: 6px; }}
    QMenu::item:selected {{ background: {t['sel']}; }}
    QToolTip {{
        background: {t['surface']}; color: {t['text']}; border: 1px solid {t['border']};
        padding: 4px;
    }}
    """


def make_app_icon() -> QIcon:
    pm = QPixmap(64, 64)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(QRectF(2, 2, 60, 60), 14, 14)
    p.fillPath(path, QColor("#16a34a"))
    # grid
    p.setPen(QPen(QColor(255, 255, 255, 90), 2))
    for x in (20, 34):
        p.drawLine(x, 12, x, 52)
    for y in (24, 38):
        p.drawLine(10, y, 54, y)
    # magnifier
    p.setPen(QPen(QColor("#ffffff"), 5))
    p.setBrush(QColor("#16a34a"))
    p.drawEllipse(QRectF(24, 20, 22, 22))
    p.setPen(QPen(QColor("#ffffff"), 6, cap=Qt.PenCapStyle.RoundCap))
    p.drawLine(43, 39, 53, 49)
    p.end()
    return QIcon(pm)



# --------------------------------------------------------------------------- #
#  Data
# --------------------------------------------------------------------------- #

@dataclass
class Hit:
    term: str
    value: str
    row: int                 # 1-based (row, paragraph or line)
    col: int                 # 1-based column (1 for Word / PDF)
    sheet: str               # sheet name; page number for PDF
    path: str
    kind: str = "sheet"      # "sheet" | "doc" | "pdf"
    span: tuple | None = None
    score: int = 100         # 100 = exact; less = similar (fuzzy)
    row_values: tuple = ()   # whole row as text (for the preview)
    header: tuple = ()       # first row of the sheet as text

    @property
    def cell(self) -> str:
        if self.kind == "doc":
            return f"¶{self.row}"
        if self.kind == "pdf":
            return f"L{self.row}"
        return f"{col_letter(self.col)}{self.row}"

    @property
    def file_name(self) -> str:
        return os.path.basename(self.path)

    @property
    def sheet_display(self) -> str:
        if self.kind == "pdf":
            return f"{tr('Page')} {self.sheet}"
        if self.kind == "doc":
            return tr("Document")
        return self.sheet


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def to_text(v) -> str | None:
    """Any cell value as the text the user would see (roughly)."""
    if v is None:
        return None
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, _dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S") if (v.hour or v.minute or v.second) else v.strftime("%Y-%m-%d")
    if isinstance(v, (_dt.date, _dt.time)):
        return v.isoformat()
    return str(v)


# --------------------------------------------------------------------------- #
#  Readers: every file becomes [(sheet_name, kind, [(row_no, (values…)), …]), …]
#  (only rows with something in them)
# --------------------------------------------------------------------------- #

def _trim(vals) -> tuple:
    vals = list(vals)
    while vals and (vals[-1] is None or vals[-1] == ""):
        vals.pop()
    return tuple(vals)


def read_xlsx(path: str):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    try:
        for ws in wb.worksheets:
            rows = []
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                vals = _trim(row)
                if vals:
                    rows.append((r_idx, vals))
            out.append((ws.title, "sheet", rows))
    finally:
        wb.close()
    return out


def read_xls(path: str):
    import xlrd
    book = xlrd.open_workbook(path, on_demand=True)
    out = []
    try:
        for si in range(book.nsheets):
            sh = book.sheet_by_index(si)
            rows = []
            for r in range(sh.nrows):
                vals = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    t = cell.ctype
                    v = cell.value
                    if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                        v = None
                    elif t == xlrd.XL_CELL_DATE:
                        try:
                            v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                        except Exception:
                            pass
                    elif t == xlrd.XL_CELL_BOOLEAN:
                        v = bool(v)
                    vals.append(v)
                vals = _trim(vals)
                if vals:
                    rows.append((r + 1, vals))
            out.append((sh.name, "sheet", rows))
            book.unload_sheet(si)
    finally:
        book.release_resources()
    return out


def read_csv(path: str):
    raw = Path(path).read_bytes()
    text = None
    for enc in ("utf-8-sig", "cp1253", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    rows = []
    for r_idx, row in enumerate(csv.reader(text.splitlines(), dialect), start=1):
        vals = _trim(row)
        if vals:
            rows.append((r_idx, vals))
    return [(Path(path).stem, "sheet", rows)]


def read_docx(path: str):
    import xml.etree.ElementTree as ET
    import zipfile
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    rows = []
    for n, p in enumerate(root.iter(w + "p"), start=1):
        parts = []
        for el in p.iter():
            if el.tag == w + "t" and el.text:
                parts.append(el.text)
            elif el.tag == w + "tab":
                parts.append("\t")
            elif el.tag in (w + "br", w + "cr"):
                parts.append("\n")
        t = "".join(parts).strip()
        if t:
            rows.append((n, (t,)))
    return [("Document", "doc", rows)]


def read_pdf(path: str):
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    out = []
    total = 0
    with fitz.open(path) as doc:
        if doc.needs_pass:
            raise RuntimeError("password-protected PDF")
        for i, page in enumerate(doc, start=1):
            lines = [ln.strip() for ln in page.get_text("text").splitlines()]
            rows = [(j, (ln,)) for j, ln in enumerate(lines, start=1) if ln]
            total += len(rows)
            out.append((str(i), "pdf", rows))
    if not total:
        raise RuntimeError("no text in this PDF (scanned image?)")
    return out


def read_any(path: str):
    ext = Path(path).suffix.lower()
    if ext == ".xls":
        return read_xls(path)
    if ext == ".csv":
        return read_csv(path)
    if ext == ".docx":
        return read_docx(path)
    if ext == ".pdf":
        return read_pdf(path)
    return read_xlsx(path)


# ---- cache: a file that has not changed is not read again ------------------ #

CACHE_FORMAT = 3
CACHE_MAX_BYTES = 512 * 1024 * 1024


def load_file(path: str, cache_dir: str | None):
    """(sheets, from_cache)."""
    cf = None
    if cache_dir:
        st = os.stat(path)
        key = hashlib.sha1(f"{CACHE_FORMAT}|{os.path.abspath(path)}|{st.st_size}|{st.st_mtime_ns}"
                           .encode("utf-8")).hexdigest()
        cf = Path(cache_dir) / f"{key}.bin"
        if cf.exists():
            try:
                data = pickle.loads(zlib.decompress(cf.read_bytes()))
                try:
                    os.utime(cf)   # most recently used
                except OSError:
                    pass
                return data, True
            except Exception:
                pass
    data = read_any(path)
    if cf is not None:
        try:
            cf.parent.mkdir(parents=True, exist_ok=True)
            tmp = cf.with_name(cf.name + f".{os.getpid()}.part")
            tmp.write_bytes(zlib.compress(pickle.dumps(data, protocol=pickle.HIGHEST_PROTOCOL), 1))
            os.replace(tmp, cf)
        except Exception:
            pass
    return data, False


def cache_size(cache_dir: Path) -> int:
    try:
        return sum(f.stat().st_size for f in cache_dir.glob("*.bin"))
    except OSError:
        return 0


def prune_cache(cache_dir: Path, max_bytes: int = CACHE_MAX_BYTES):
    """Keep the cache under max_bytes, dropping the least recently used files first."""
    try:
        files = sorted(cache_dir.glob("*.bin"), key=lambda f: f.stat().st_mtime)
        total = sum(f.stat().st_size for f in files)
        for f in files:
            if total <= max_bytes:
                break
            total -= f.stat().st_size
            f.unlink(missing_ok=True)
        for f in cache_dir.glob("*.part"):
            if time.time() - f.stat().st_mtime > 3600:
                f.unlink(missing_ok=True)
    except OSError:
        pass


# --------------------------------------------------------------------------- #
#  Text normalizing: case, accents, Greek <-> Latin
# --------------------------------------------------------------------------- #

_ACCENT_MAP = str.maketrans(
    "άέήίόύώϊϋΐΰΆΈΉΊΌΎΏΪΫ" "àáâãäåèéêëìíîïòóôõöùúûüýÿñçÀÁÂÃÄÅÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÝÑÇ",
    "αεηιουωιυιυΑΕΗΙΟΥΩΙΥ" "aaaaaaeeeeiiiiooooouuuuyyncAAAAAAEEEEIIIIOOOOOUUUUYNC",
)


def strip_accents(s: str) -> str:
    return s.translate(_ACCENT_MAP)


def _lower(s: str) -> str:
    low = s.lower()
    if len(low) == len(s):
        return low
    return "".join(c.lower() if len(c.lower()) == 1 else c for c in s)   # keep positions 1:1


# Greeklish / sound: both Greek and Latin text are turned into the same "sound skeleton",
# so σκυροδεμα = skyrodema = skirodema, μπετον = beton, ντουλαπα = doulapa, θεμελιο = themelio.
_PHON_DIGRAPHS = {"th": "θ", "ps": "ψ", "ks": "ξ", "ch": "x", "ph": "f", "mp": "v", "nt": "d",
                  "gk": "g", "gg": "g", "μπ": "v", "ντ": "d", "γκ": "g", "γγ": "g"}
_PHON_RX = re.compile("|".join(sorted(_PHON_DIGRAPHS, key=len, reverse=True)))
_PHON_TABLE = str.maketrans(
    "αβγδεζηικλμνοπρσςτυφχω" "bhywucq",
    "avgdeziiklmnoprsstifxo" "viioikk",
)
# Wrong keyboard: a Latin code typed with the Greek layout on (Β500C, ΑΒ-12 …): key positions.
_KBD_TABLE = str.maketrans("αβγδεζηθικλμνξοπρσςτυφχψω", "abgdezhuiklmnjoprswtyfxcv")
# Look-alike letters: Greek capitals that look like Latin ones (Α Β Ε Η Ι Κ Μ Ν Ο Ρ Τ Υ Χ Ζ).
_VIS_TABLE = str.maketrans("αβεζηικμνορτυχ", "abezhikmnoptyx")


class Norm:
    """One way of normalizing text. fast() for searching, mapped() to find the match in the original."""

    def __init__(self, case: bool, accents: bool, kind: str | None = None):
        self.case, self.accents, self.kind = case, accents, kind
        self.table = {"phon": _PHON_TABLE, "kbd": _KBD_TABLE, "vis": _VIS_TABLE}.get(kind)

    def _pre(self, s: str) -> str:
        if not self.case:
            s = _lower(s)
        if self.accents:
            s = s.translate(_ACCENT_MAP)
        return s

    def fast(self, s: str) -> str:
        s = self._pre(s)
        if self.kind == "phon":
            s = _PHON_RX.sub(lambda m: _PHON_DIGRAPHS[m.group()], s)
        return s.translate(self.table) if self.table else s

    def mapped(self, s: str):
        """(normalized, spans) where spans[i] = (start, end) in s of normalized char i; None = 1:1."""
        s1 = self._pre(s)
        if self.kind != "phon":
            return (s1.translate(self.table) if self.table else s1), None
        out, spans, pos = [], [], 0
        for m in _PHON_RX.finditer(s1):
            for i in range(pos, m.start()):
                out.append(s1[i])
                spans.append((i, i + 1))
            out.append(_PHON_DIGRAPHS[m.group()])
            spans.append((m.start(), m.end()))
            pos = m.end()
        for i in range(pos, len(s1)):
            out.append(s1[i])
            spans.append((i, i + 1))
        return "".join(out).translate(self.table), spans

    def back(self, s: str, a: int, b: int) -> tuple:
        _, spans = self.mapped(s) if self.kind == "phon" else (None, None)
        if not spans or b <= a:
            return (a, b)
        a = min(a, len(spans) - 1)
        b = min(b, len(spans))
        return (spans[a][0], spans[b - 1][1])


# --------------------------------------------------------------------------- #
#  Matching
# --------------------------------------------------------------------------- #

MODES = [("contains", "Contains"), ("whole", "Whole cell"), ("starts", "Starts with"),
         ("word", "Whole word"), ("fuzzy", "Similar (typos)"), ("regex", "Regular expression")]
_OLD_MODE_LABELS = {lbl: key for key, lbl in MODES}

try:
    from rapidfuzz import fuzz as _rf_fuzz
except Exception:   # installed by the splash; this is only a fallback
    _rf_fuzz = None


def _fuzzy(tn: str, t: str, th: int):
    """(start, end, score) of the best similar part of t, or None."""
    if not t:
        return None
    if _rf_fuzz is not None:
        if len(t) <= len(tn):
            sc = _rf_fuzz.ratio(tn, t)
            return (0, len(t), int(round(sc))) if sc >= th else None
        al = _rf_fuzz.partial_ratio_alignment(tn, t, score_cutoff=th)
        if al is None or al.score < th:
            return None
        a, b = max(0, al.dest_start), min(len(t), al.dest_end)
        if b <= a:
            a, b = 0, len(t)
        return (a, b, int(round(al.score)))
    import difflib
    n = len(tn)
    best = (0, 0, 0)
    for i in range(0, max(1, len(t) - n + 1)):
        sc = difflib.SequenceMatcher(None, tn, t[i:i + n]).ratio() * 100
        if sc > best[2]:
            best = (i, min(len(t), i + n), sc)
    return (best[0], best[1], int(round(best[2]))) if best[2] >= th else None


_COMPACT_SKIP = set(" \t\n-_./,:;·")


def _compact(s: str):
    """s without spaces, dashes, dots… and where each kept char came from: 'B500 C' -> 'B500C'."""
    keep = [i for i, ch in enumerate(s) if ch not in _COMPACT_SKIP]
    return "".join(s[i] for i in keep), keep


class Matcher:
    def __init__(self, term: str, mode: str, match_case: bool, accents: bool, greek: bool, fuzzy_th: int):
        self.term, self.mode, self.th = term, mode, fuzzy_th
        if mode == "regex":
            self.norms = [Norm(True, accents)]
            self.rxs = [re.compile(term, 0 if match_case else re.IGNORECASE)]
            self.tns = [term]
            return
        if greek:
            self.norms = [Norm(False, True, k) for k in (None, "phon", "kbd", "vis")]
        else:
            self.norms = [Norm(match_case, accents)]
        self.tns, self.rxs = [], []
        for n in self.norms:
            tn = n.fast(term).strip()
            e = re.escape(tn)
            pat = {"whole": r"^\s*" + e + r"\s*$", "starts": r"^\s*" + e,
                   "word": r"(?<!\w)" + e + r"(?!\w)"}.get(mode, e)
            self.tns.append(tn)
            self.rxs.append(re.compile(pat))

    def search(self, text: str):
        """((start, end), score) or None."""
        for n, rx, tn in zip(self.norms, self.rxs, self.tns):
            t = n.fast(text)
            if self.mode == "fuzzy" and len(tn) >= 4:
                r = _fuzzy(tn, t, self.th)
                tc, keep = _compact(t)
                tnc = _compact(tn)[0]
                if len(tnc) >= 4 and keep:
                    r2 = _fuzzy(tnc, tc, self.th)       # B500C = B500 C = B-500-C
                    if r2 and (not r or r2[2] > r[2]):
                        r = (keep[r2[0]], keep[max(r2[0], r2[1] - 1)] + 1, r2[2])
                if r:
                    return n.back(text, r[0], r[1]), r[2]
                continue
            m = rx.search(t)
            if m:
                return n.back(text, *m.span()), 100
        return None


def parse_columns(spec: str):
    """'B, D:F, 3, "Qty", Περιγραφή' -> (column numbers, header names)."""
    ints, names = set(), []
    for tok in re.split(r"[,;]", spec or ""):
        t = tok.strip()
        if not t:
            continue
        if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'":
            names.append(t[1:-1].strip().casefold())
            continue
        m = re.fullmatch(r"([A-Za-z]{1,3})\s*[:\-]\s*([A-Za-z]{1,3})", t)
        if m:
            a, b = col_index(m[1]), col_index(m[2])
            ints.update(range(min(a, b), max(a, b) + 1))
            continue
        m = re.fullmatch(r"(\d+)\s*[:\-]\s*(\d+)", t)
        if m:
            a, b = int(m[1]), int(m[2])
            ints.update(range(min(a, b), max(a, b) + 1))
            continue
        if re.fullmatch(r"[A-Za-z]{1,3}", t):
            ints.add(col_index(t))
        elif t.isdigit():
            ints.add(int(t))
        else:
            names.append(t.casefold())
    return sorted(ints), names


def parse_sheets(spec: str) -> list[str]:
    return [t.strip().casefold() for t in re.split(r"[,;]", spec or "") if t.strip()]


def sheet_ok(name: str, pats: list[str]) -> bool:
    nf = name.casefold()
    for p in pats:
        if any(ch in p for ch in "*?["):
            if fnmatch.fnmatchcase(nf, p):
                return True
        elif p in nf:
            return True
    return False


def search_file(path: str, p: dict):
    """Searches one file. Runs in a worker process, so it returns plain tuples:
    (path, [hit tuples], error or None, from_cache)."""
    try:
        data, cached = load_file(path, p.get("cache_dir"))
    except Exception as e:  # noqa: BLE001
        return path, [], f"{type(e).__name__}: {e}", False
    try:
        matchers = [Matcher(t, p["mode"], p["case"], p["accents"], p["greek"], p["fuzzy"]) for t in p["terms"]]
        col_ints, col_names = p["cols"]
        sheet_pats = p["sheets"]
        out = []
        for sheet, kind, rows in data:
            if kind == "sheet" and sheet_pats and not sheet_ok(sheet, sheet_pats):
                continue
            header = tuple(to_text(v) or "" for v in rows[0][1]) if (kind == "sheet" and rows) else ()
            allowed = None
            if kind == "sheet" and (col_ints or col_names):
                allowed = set(col_ints)
                for i, h in enumerate(header, start=1):
                    hf = h.casefold().strip()
                    if hf and any(nm == hf or nm in hf for nm in col_names):
                        allowed.add(i)
            seen = set()
            for r, vals in rows:
                row_txt = None
                for c, v in enumerate(vals, start=1):
                    if v is None or v == "":
                        continue
                    if allowed is not None and c not in allowed:
                        continue
                    if p["text_only"] and not isinstance(v, str):
                        continue
                    text = to_text(v)
                    if not text:
                        continue
                    for mi, m in enumerate(matchers):
                        res = m.search(text)
                        if not res:
                            continue
                        if p["one_per_row"]:
                            key = (mi, sheet, r)
                            if key in seen:
                                continue
                            seen.add(key)
                        if row_txt is None:
                            row_txt = tuple(to_text(x) or "" for x in vals)
                        out.append((m.term, text, r, c, sheet, kind, res[0], res[1], row_txt, header))
        return path, out, None, cached
    except Exception as e:  # noqa: BLE001
        return path, [], f"{type(e).__name__}: {e}", False


class SearchWorker(QObject):
    progress = pyqtSignal(int, int, str)       # done, total, message
    hits = pyqtSignal(list)                    # batch of Hit
    file_error = pyqtSignal(str, str)          # path, error
    finished = pyqtSignal(float, bool, int, int)   # seconds, cancelled, files from cache, worker processes

    def __init__(self, files: list[str], params: dict, parallel: bool):
        super().__init__()
        self.files, self.p, self.parallel = files, params, parallel
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def _take(self, result):
        path, rows, err, cached = result
        if err:
            self.file_error.emit(path, err)
        if rows:
            self.hits.emit([Hit(term=t, value=v, row=r, col=c, sheet=s, path=path, kind=k, span=sp, score=sc,
                                row_values=rv, header=hd)
                            for (t, v, r, c, s, k, sp, sc, rv, hd) in rows])
        return 1 if cached else 0

    def run(self):
        t0 = time.perf_counter()
        total = len(self.files)
        remaining = list(self.files)
        done = cached = procs = 0
        if self.parallel and total > 1 and (os.cpu_count() or 1) > 1:
            ex = None
            try:
                import multiprocessing as mp
                from concurrent.futures import ProcessPoolExecutor, as_completed
                ctx = mp.get_context("spawn")
                if os.name == "nt":
                    ctx.set_executable(_gui_python())   # pythonw: no console windows for the workers
                procs = max(1, min(total, (os.cpu_count() or 2) - 1, 8))
                ex = ProcessPoolExecutor(max_workers=procs, mp_context=ctx)
                futs = {ex.submit(search_file, f, self.p): f for f in remaining}
                self.progress.emit(0, total, tr("Searching with {n} processes").format(n=procs))
                for fut in as_completed(futs):
                    if self._cancel:
                        break
                    res = fut.result()
                    remaining.remove(futs[fut])
                    cached += self._take(res)
                    done += 1
                    self.progress.emit(done, total, os.path.basename(futs[fut]))
            except Exception:
                procs = 0     # workers could not start: do the rest here, one by one
            finally:
                if ex is not None:
                    ex.shutdown(wait=not self._cancel, cancel_futures=True)
        for f in list(remaining):
            if self._cancel:
                break
            self.progress.emit(done, total, tr("Reading {name}").format(name=os.path.basename(f)))
            cached += self._take(search_file(f, self.p))
            done += 1
        self.progress.emit(total, total, tr("Done"))
        self.finished.emit(time.perf_counter() - t0, self._cancel, cached, procs)


# --------------------------------------------------------------------------- #
#  Results model / view
# --------------------------------------------------------------------------- #

COLUMNS = ["#", "Term", "Found value", "Match", "Cell", "Row", "Sheet / page", "File"]
SORT_ROLE = Qt.ItemDataRole.UserRole + 1
HIT_ROLE = Qt.ItemDataRole.UserRole + 2


class ResultsModel(QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self.rows: list[Hit] = []
        self.muted = QColor("#8b93a7")
        self.warn = QColor("#f59e0b")

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(COLUMNS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return tr(COLUMNS[section]).upper()
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        h = self.rows[index.row()]
        col = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            return [index.row() + 1, h.term, h.value.replace("\n", " ⏎ "), f"{h.score}%", h.cell, h.row,
                    h.sheet_display, h.file_name][col]
        if role == SORT_ROLE:
            return [index.row(), h.term.casefold(), h.value.casefold(), h.score, (h.col, h.row), h.row,
                    h.sheet_display.casefold(), h.file_name.casefold()][col]
        if role == HIT_ROLE:
            return h
        if role == Qt.ItemDataRole.ToolTipRole:
            if col == 7:
                return h.path
            if col == 2:
                return h.value
            if col == 3:
                return tr("Exact match") if h.score >= 100 else tr("Similar: {p}% alike").format(p=h.score)
        if role == Qt.ItemDataRole.ForegroundRole:
            if col in (0, 5):
                return QBrush(self.muted)
            if col == 3:
                return QBrush(self.muted if h.score >= 100 else self.warn)
        if role == Qt.ItemDataRole.TextAlignmentRole and col in (0, 3, 4, 5):
            return int(Qt.AlignmentFlag.AlignCenter)
        return None

    def add(self, hits: list[Hit]):
        if not hits:
            return
        start = len(self.rows)
        self.beginInsertRows(QModelIndex(), start, start + len(hits) - 1)
        self.rows.extend(hits)
        self.endInsertRows()

    def clear(self):
        self.beginResetModel()
        self.rows.clear()
        self.endResetModel()


class FilterProxy(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.needle = ""
        self.setSortRole(SORT_ROLE)

    def set_needle(self, s: str):
        self.needle = s.casefold().strip()
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        if not self.needle:
            return True
        h: Hit = self.sourceModel().rows[row]
        hay = f"{h.term}\t{h.value}\t{h.sheet_display}\t{h.file_name}\t{h.cell}".casefold()
        return self.needle in hay


class HighlightDelegate(QStyledItemDelegate):
    """Paints the 'Found value' column with the matched part highlighted."""

    def __init__(self, parent, theme_getter):
        super().__init__(parent)
        self.theme = theme_getter

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index):
        h: Hit = index.data(HIT_ROLE)
        if h is None or not h.span:
            return super().paint(painter, option, index)
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = opt.text
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)

        t = self.theme()
        s, e = h.span
        pre = h.value[:s].replace("\n", " ⏎ ")
        mid = h.value[s:e].replace("\n", " ⏎ ")
        post = text[len(pre) + len(mid):]

        fm = QFontMetrics(opt.font)
        rect = opt.rect.adjusted(8, 0, -6, 0)
        avail = rect.width()
        if fm.horizontalAdvance(pre + mid) > avail and len(pre) > 12:   # keep the hit in view
            keep = pre
            while len(keep) > 1 and fm.horizontalAdvance("…" + keep + mid) > avail * 0.7:
                keep = keep[1:]
            pre = "…" + keep

        painter.save()
        painter.setClipRect(opt.rect)
        x = rect.left()
        y_center = rect.center().y()
        h_text = fm.height()
        baseline = int(y_center - h_text / 2 + fm.ascent())
        text_color = QColor(t["text"])
        painter.setPen(text_color)
        painter.drawText(x, baseline, pre)
        x += fm.horizontalAdvance(pre)
        w_mid = fm.horizontalAdvance(mid)
        if x < rect.right():
            bg = QRectF(x - 1, y_center - h_text / 2 - 1, w_mid + 2, h_text + 2)
            path = QPainterPath()
            path.addRoundedRect(bg, 3, 3)
            painter.fillPath(path, QColor(t["hl_bg"] if h.score >= 100 else t["hl_fuzzy"]))
            painter.setPen(QColor(t["hl_text"]))
            painter.drawText(x, baseline, mid)
        x += w_mid
        painter.setPen(text_color)
        rest_w = rect.right() - x
        if rest_w > 0:
            painter.drawText(x, baseline, fm.elidedText(post, Qt.TextElideMode.ElideRight, rest_w))
        painter.restore()


# --------------------------------------------------------------------------- #
#  File list with drag & drop
# --------------------------------------------------------------------------- #

class FileList(QListWidget):
    files_dropped = pyqtSignal(list)
    delete_pressed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setTextElideMode(Qt.TextElideMode.ElideMiddle)

    def _set_over(self, on: bool):
        self.setObjectName("dragOver" if on else "")
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self._set_over(True)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dragLeaveEvent(self, e):
        self._set_over(False)

    def dropEvent(self, e):
        self._set_over(False)
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.files_dropped.emit(paths)
            e.acceptProposedAction()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key.Key_Delete:
            self.delete_pressed.emit()
            return
        super().keyPressEvent(e)


# --------------------------------------------------------------------------- #
#  Helpers
# --------------------------------------------------------------------------- #

def open_path(path: str):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:  # noqa: BLE001
        QMessageBox.warning(None, APP_NAME, tr("Could not open:") + f"\n{path}\n\n{e}")


def reveal_in_folder(path: str):
    if sys.platform.startswith("win"):
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        open_path(os.path.dirname(path))


_EXCEL_GOTO_PS = r"""
$ErrorActionPreference = 'Stop'
$f = $env:EUS_FILE
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { $xl = New-Object -ComObject Excel.Application }
$xl.Visible = $true
$wb = $null
foreach ($w in @($xl.Workbooks)) { if ($w.FullName -ieq $f) { $wb = $w } }
if (-not $wb) { $wb = $xl.Workbooks.Open($f) }
try { $ws = $wb.Worksheets.Item($env:EUS_SHEET) } catch { $ws = $wb.Worksheets.Item(1) }
$ws.Activate()
$rg = $ws.Range($env:EUS_CELL)
$xl.Goto($rg, $true)
try { $xl.ActiveWindow.ScrollRow = [Math]::Max(1, $rg.Row - 5) } catch {}
try { $xl.ActiveWindow.ScrollColumn = [Math]::Max(1, $rg.Column - 2) } catch {}
$rg.Select() | Out-Null
try {
  Add-Type -Namespace EUS -Name W -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h); [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int n);'
  $h = [IntPtr]$xl.Hwnd
  [EUS.W]::ShowWindow($h, 9) | Out-Null
  [EUS.W]::SetForegroundWindow($h) | Out-Null
} catch {}
"""


def open_at_cell(h: Hit, on_status):
    """Open the workbook in Excel at the sheet and cell of the hit. Falls back to opening the file."""
    if os.name != "nt" or h.kind != "sheet":
        open_path(h.path)
        return
    on_status(tr("Opening in Excel at {where}…").format(where=f"{h.sheet}!{h.cell}"))
    env = dict(os.environ, EUS_FILE=os.path.normpath(h.path), EUS_SHEET=h.sheet, EUS_CELL=h.cell)

    def work():
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                            "-WindowStyle", "Hidden", "-Command", _EXCEL_GOTO_PS],
                           env=env, capture_output=True, text=True, timeout=120, creationflags=NO_WINDOW)
        return r.returncode == 0

    def done(ok, err):
        if not ok:
            on_status(tr("Excel could not be driven; opening the file normally"))
            open_path(h.path)
        else:
            on_status(tr("Opened at {where}").format(where=f"{h.sheet}!{h.cell}"))

    run_in_background(work, done)


def card() -> QFrame:
    f = QFrame()
    f.setObjectName("card")
    return f


def section(text: str) -> QLabel:
    lb = QLabel(text.upper())
    lb.setObjectName("section")
    return lb


def hint(text: str) -> QLabel:
    lb = QLabel(text)
    lb.setObjectName("hint")
    lb.setWordWrap(True)
    return lb


def sbool(v, default=False) -> bool:
    if v is None:
        return default
    return str(v).lower() in ("true", "1")


def slist(v) -> list:
    if not v:
        return []
    return [v] if isinstance(v, str) else list(v)


def safe_sheet_title(name: str, used: set) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", name).strip("'") or "Sheet"
    base = base[:31]
    title, i = base, 2
    while title.casefold() in used:
        suffix = f" ({i})"
        title = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(title.casefold())
    return title


# --------------------------------------------------------------------------- #
#  Saved term lists
# --------------------------------------------------------------------------- #

def _lists_file() -> Path:
    return _app_data_dir() / "term_lists.json"


def load_term_lists() -> dict:
    try:
        d = json.loads(_lists_file().read_text(encoding="utf-8"))
        return {k: list(v) for k, v in d.items()} if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_term_lists(d: dict):
    f = _lists_file()
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, f)


class ColumnImportDialog(QDialog):
    """Pick a sheet and a column of a spreadsheet; its cells become the search terms."""

    def __init__(self, parent, path: str, sheets):
        super().__init__(parent)
        self.setWindowTitle(tr("Import terms from a column"))
        self.setMinimumWidth(460)
        self.sheets = [s for s in sheets if s[1] == "sheet"]
        self.result_terms: list[str] = []
        self.mode = "replace"
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(10)
        lay.addWidget(QLabel(f"<b>{os.path.basename(path)}</b>"))
        lay.addWidget(QLabel(tr("Sheet")))
        self.sheet_cb = QComboBox()
        self.sheet_cb.addItems([s[0] for s in self.sheets])
        lay.addWidget(self.sheet_cb)
        lay.addWidget(QLabel(tr("Column")))
        self.col_cb = QComboBox()
        lay.addWidget(self.col_cb)
        self.skip = QCheckBox(tr("First row is a header (skip it)"))
        self.skip.setChecked(True)
        lay.addWidget(self.skip)
        self.info = hint("")
        lay.addWidget(self.info)
        row = QHBoxLayout()
        b_rep = QPushButton(tr("Replace terms"), objectName="primary")
        b_add = QPushButton(tr("Add to terms"))
        b_can = QPushButton(tr("Cancel"))
        b_rep.clicked.connect(lambda: self._finish("replace"))
        b_add.clicked.connect(lambda: self._finish("add"))
        b_can.clicked.connect(self.reject)
        row.addWidget(b_rep)
        row.addWidget(b_add)
        row.addStretch()
        row.addWidget(b_can)
        lay.addLayout(row)
        self.sheet_cb.currentIndexChanged.connect(self._fill_cols)
        self.col_cb.currentIndexChanged.connect(self._count)
        self.skip.toggled.connect(self._count)
        self._fill_cols()

    def _rows(self):
        i = self.sheet_cb.currentIndex()
        return self.sheets[i][2] if 0 <= i < len(self.sheets) else []

    def _fill_cols(self):
        rows = self._rows()
        n = max((len(v) for _r, v in rows), default=0)
        header = rows[0][1] if rows else ()
        self.col_cb.blockSignals(True)
        self.col_cb.clear()
        best = 0
        for c in range(1, n + 1):
            h = to_text(header[c - 1]) if c <= len(header) else ""
            self.col_cb.addItem(f"{col_letter(c)}" + (f"  —  {h}" if h else ""), c)
        # pick the column with the most text in it (not the numbering one)
        filled = [sum(1 for _r, v in rows[1:] if c <= len(v) and isinstance(v[c - 1], str) and v[c - 1].strip())
                  for c in range(1, n + 1)]
        if filled:
            best = max(range(len(filled)), key=lambda i: filled[i])
        self.col_cb.setCurrentIndex(best)
        self.col_cb.blockSignals(False)
        self._count()

    def _values(self) -> list[str]:
        c = self.col_cb.currentData()
        if not c:
            return []
        rows = self._rows()[1:] if self.skip.isChecked() else self._rows()
        seen, out = set(), []
        for _r, v in rows:
            if c <= len(v):
                t = (to_text(v[c - 1]) or "").strip()
                if t and t not in seen:
                    seen.add(t)
                    out.append(t)
        return out

    def _count(self):
        vals = self._values()
        prev = ", ".join(vals[:4]) + ("…" if len(vals) > 4 else "")
        self.info.setText(tr("{n} distinct values").format(n=len(vals)) + (f":  {prev}" if vals else ""))

    def _finish(self, mode):
        self.result_terms = self._values()
        self.mode = mode
        self.accept()


# --------------------------------------------------------------------------- #
#  Main window
# --------------------------------------------------------------------------- #

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.theme_name = self.settings.value("theme", "dark")
        if self.theme_name not in THEMES:
            self.theme_name = "dark"
        self.thread: QThread | None = None
        self.worker: SearchWorker | None = None
        self.errors: list[tuple[str, str]] = []
        self.last_terms: list[str] = []
        self.folders: list[str] = []          # watched folders
        self.last_dir = str(Path.home())
        self.watcher = QFileSystemWatcher(self)
        self.watcher.directoryChanged.connect(lambda _p: self._rescan_timer.start())
        self._rescan_timer = QTimer(self, singleShot=True, interval=1500)
        self._rescan_timer.timeout.connect(lambda: self.rescan_folders(auto=True))

        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setWindowIcon(make_app_icon())
        self.resize(1360, 820)
        self.setMinimumSize(1000, 620)

        self._build_ui()
        self._restore()
        self.apply_theme()

    # ================= UI ================= #
    def _build_ui(self):
        root = QWidget(objectName="root")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 14, 18, 14)
        outer.setSpacing(12)

        hdr = QHBoxLayout()
        icon = QLabel()
        icon.setPixmap(make_app_icon().pixmap(36, 36))
        titles = QVBoxLayout()
        titles.setSpacing(0)
        titles.addWidget(QLabel(APP_NAME, objectName="title"))
        titles.addWidget(QLabel(tr("Search many spreadsheets, Word and PDF files for many terms — at once."),
                                objectName="subtitle"))
        hdr.addWidget(icon)
        hdr.addSpacing(6)
        hdr.addLayout(titles)
        hdr.addStretch()
        self.lang_btn = QPushButton("🌐  English" if LANG == "el" else "🌐  Ελληνικά", objectName="ghost")
        self.lang_btn.setToolTip(tr("Change the language (the app restarts)"))
        self.lang_btn.clicked.connect(self.toggle_lang)
        self.theme_btn = QPushButton(objectName="ghost")
        self.theme_btn.setToolTip(tr("Light / dark theme"))
        self.theme_btn.clicked.connect(self.toggle_theme)
        self.about_btn = QPushButton(f"ⓘ  v{APP_VERSION}", objectName="ghost")
        self.about_btn.setToolTip(tr("About, and check for updates"))
        self.about_btn.clicked.connect(lambda: AboutDialog(self).exec())
        for b in (self.lang_btn, self.theme_btn, self.about_btn):
            hdr.addWidget(b)
        outer.addLayout(hdr)

        split = QSplitter(Qt.Orientation.Horizontal)
        split.setChildrenCollapsible(False)
        outer.addWidget(split, 1)
        self.splitter = split
        split.addWidget(self._build_left())
        split.addWidget(self._build_right())
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([420, 980])

        QShortcut(QKeySequence("Ctrl+Return"), self, activated=self.toggle_search)
        QShortcut(QKeySequence("Ctrl+Enter"), self, activated=self.toggle_search)
        QShortcut(QKeySequence("Ctrl+F"), self,
                  activated=lambda: (self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        QShortcut(QKeySequence("Ctrl+O"), self, activated=self.pick_files)
        QShortcut(QKeySequence("Escape"), self, activated=self._escape)
        QShortcut(QKeySequence("F5"), self, activated=lambda: self.rescan_folders(auto=False))
        self.update_counts()
        self._update_empty()

    def _build_left(self) -> QWidget:
        left = card()
        left.setMinimumWidth(400)
        outer = QVBoxLayout(left)
        outer.setContentsMargins(0, 0, 0, 14)
        outer.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        content = QWidget(objectName="scrollContent")
        lv = QVBoxLayout(content)
        lv.setContentsMargins(16, 14, 14, 6)
        lv.setSpacing(8)
        scroll.setWidget(content)
        outer.addWidget(scroll, 1)

        # ---- 1. files ----
        fh = QHBoxLayout()
        fh.addWidget(section(tr("1 · Files")))
        fh.addStretch()
        self.files_count = QLabel("0", objectName="subtitle")
        fh.addWidget(self.files_count)
        lv.addLayout(fh)

        fb = QHBoxLayout()
        fb.setSpacing(6)
        b_add = QPushButton(tr("＋ Files"))
        b_add.setToolTip(tr("Excel, CSV, Word (.docx) or PDF files — Ctrl+O"))
        b_add.clicked.connect(self.pick_files)
        b_dir = QPushButton(tr("＋ Folder"))
        b_dir.setToolTip(tr("A folder with everything inside it; it is watched for new files"))
        b_dir.clicked.connect(self.pick_folder)
        self.recent_btn = QPushButton(tr("Recent ▾"))
        self.recent_menu = QMenu(self)
        self.recent_menu.aboutToShow.connect(self._fill_recent_menu)
        self.recent_btn.setMenu(self.recent_menu)
        for b in (b_add, b_dir, self.recent_btn):
            fb.addWidget(b)
        lv.addLayout(fb)

        self.file_list = FileList()
        self.file_list.files_dropped.connect(self.add_paths)
        self.file_list.delete_pressed.connect(self.remove_selected_files)
        self.file_list.setMinimumHeight(130)
        self.drop_hint = QLabel(tr("Drop Excel / CSV / Word / PDF\nfiles or folders here"), objectName="dropHint")
        self.drop_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_hint.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.drop_hint.setParent(self.file_list.viewport())
        lv.addWidget(self.file_list, 2)

        fb2 = QHBoxLayout()
        fb2.setSpacing(6)
        self.watch_btn = QPushButton("")
        self.watch_btn.setToolTip(tr("Folders watched for new or deleted files"))
        self.watch_menu = QMenu(self)
        self.watch_menu.aboutToShow.connect(self._fill_watch_menu)
        self.watch_btn.setMenu(self.watch_menu)
        b_rm = QPushButton(tr("Remove"))
        b_rm.setToolTip(tr("Remove the selected files from the list (Delete)"))
        b_rm.clicked.connect(self.remove_selected_files)
        b_clr = QPushButton(tr("Clear"), objectName="danger")
        b_clr.clicked.connect(self.clear_files)
        lv.addWidget(self.watch_btn)
        fb2.addWidget(b_rm)
        fb2.addWidget(b_clr)
        lv.addLayout(fb2)

        # ---- 2. terms ----
        lv.addSpacing(6)
        th = QHBoxLayout()
        th.addWidget(section(tr("2 · Search terms")))
        th.addStretch()
        self.terms_count = QLabel("0", objectName="subtitle")
        th.addWidget(self.terms_count)
        self.lists_btn = QPushButton(tr("Lists ▾"))
        self.lists_btn.setToolTip(tr("Save, load and import lists of terms"))
        self.lists_menu = QMenu(self)
        self.lists_menu.aboutToShow.connect(self._fill_lists_menu)
        self.lists_btn.setMenu(self.lists_menu)
        th.addWidget(self.lists_btn)
        lv.addLayout(th)
        self.terms_edit = QPlainTextEdit()
        self.terms_edit.setPlaceholderText(tr("One term per line…\n\nOr: Lists ▾ → Import from an Excel column"))
        self.terms_edit.setMinimumHeight(120)
        self.terms_edit.textChanged.connect(self.update_counts)
        lv.addWidget(self.terms_edit, 2)

        # ---- 3. options ----
        lv.addSpacing(6)
        lv.addWidget(section(tr("3 · Options")))
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel(tr("Match")))
        self.mode_cb = QComboBox()
        for key, label in MODES:
            self.mode_cb.addItem(tr(label), key)
        mode_row.addWidget(self.mode_cb, 1)
        self.fuzzy_spin = QSpinBox()
        self.fuzzy_spin.setRange(50, 100)
        self.fuzzy_spin.setSuffix(" %")
        self.fuzzy_spin.setValue(85)
        self.fuzzy_spin.setMinimumWidth(80)
        self.fuzzy_spin.setToolTip(tr("How alike a text must be to count (Similar mode). Lower = more results."))
        mode_row.addWidget(self.fuzzy_spin)
        lv.addLayout(mode_row)
        self.mode_cb.currentIndexChanged.connect(self._mode_changed)

        self.cb_case = QCheckBox(tr("Match case"))
        self.cb_accents = QCheckBox(tr("Ignore accents (ά = α, é = e)"))
        self.cb_greek = QCheckBox(tr("Greek = Latin (Greeklish, wrong keyboard)"))
        self.cb_greek.setToolTip(tr("skyrodema finds σκυρόδεμα, beton finds μπετόν, and a code typed with the "
                                    "Greek keyboard (Β500C with a Greek Β) finds B500C. Ignores case."))
        self.cb_text_only = QCheckBox(tr("Text cells only (skip numbers && dates)"))
        self.cb_one_per_row = QCheckBox(tr("One hit per row per term"))
        self.cb_accents.setChecked(True)
        for cb in (self.cb_case, self.cb_accents, self.cb_greek, self.cb_text_only, self.cb_one_per_row):
            lv.addWidget(cb)
        self.cb_greek.toggled.connect(lambda on: self.cb_case.setEnabled(not on))

        lv.addWidget(QLabel(tr("Only these columns")))
        self.cols_edit = QLineEdit()
        self.cols_edit.setPlaceholderText(tr("all  ·  e.g.  B, D:F, Περιγραφή"))
        self.cols_edit.setClearButtonEnabled(True)
        self.cols_edit.setToolTip(tr("Column letters (B), ranges (D:F), numbers, or header names from the first "
                                     "row. Put a short header name in quotes: \"Qty\". Empty = all columns."))
        lv.addWidget(self.cols_edit)
        lv.addWidget(QLabel(tr("Only these sheets")))
        self.sheets_edit = QLineEdit()
        self.sheets_edit.setPlaceholderText(tr("all  ·  e.g.  Επιμετρήσεις*, Υλικά"))
        self.sheets_edit.setClearButtonEnabled(True)
        self.sheets_edit.setToolTip(tr("Sheet names or parts of them, comma separated; * and ? as wildcards. "
                                       "Empty = all sheets. Word and PDF files are always searched whole."))
        lv.addWidget(self.sheets_edit)

        self.cb_parallel = QCheckBox(tr("Use all processor cores (faster)"))
        self.cb_cache = QCheckBox(tr("Remember unchanged files (faster repeat searches)"))
        self.cb_parallel.setChecked(True)
        self.cb_cache.setChecked(True)
        lv.addWidget(self.cb_parallel)
        lv.addWidget(self.cb_cache)
        lv.addStretch()

        # buttons and lists share the panel's width instead of pushing it wider (long Greek labels)
        for wdg in content.findChildren(QPushButton) + content.findChildren(QComboBox):
            if wdg is self.lists_btn:
                continue
            pol = wdg.sizePolicy()
            pol.setHorizontalPolicy(QSizePolicy.Policy.Ignored)
            wdg.setSizePolicy(pol)
            wdg.setMinimumWidth(60)

        # ---- 4. search (always visible) ----
        bottom = QVBoxLayout()
        bottom.setContentsMargins(16, 0, 16, 0)
        self.search_btn = QPushButton(tr("4 · Search"), objectName="primary")
        self.search_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.search_btn.setToolTip("Ctrl+Enter")
        self.search_btn.clicked.connect(self.toggle_search)
        bottom.addWidget(self.search_btn)
        outer.addLayout(bottom)
        return left

    def _build_right(self) -> QWidget:
        right = card()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(16, 14, 16, 12)
        rv.setSpacing(10)

        rh = QHBoxLayout()
        rh.addWidget(section(tr("Results")))
        rh.addSpacing(8)
        self.stat_hits = QLabel("", objectName="stat")
        self.stat_files = QLabel("", objectName="stat")
        self.stat_missing = QPushButton("", objectName="ghost")
        self.stat_missing.setVisible(False)
        self.stat_missing.clicked.connect(self.show_missing)
        self.stat_errors = QPushButton("", objectName="ghost")
        self.stat_errors.setVisible(False)
        self.stat_errors.clicked.connect(self.show_errors)
        for w in (self.stat_hits, self.stat_files, self.stat_missing, self.stat_errors):
            rh.addWidget(w)
        rh.addStretch()
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText(tr("Filter results…  (Ctrl+F)"))
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.setFixedWidth(280)
        rh.addWidget(self.filter_edit)
        rv.addLayout(rh)

        self.model = ResultsModel()
        self.proxy = FilterProxy()
        self.proxy.setSourceModel(self.model)
        self.filter_edit.textChanged.connect(self._on_filter)

        tv = QTableView()
        tv.setModel(self.proxy)
        tv.setSortingEnabled(True)
        tv.sortByColumn(0, Qt.SortOrder.AscendingOrder)
        tv.setAlternatingRowColors(True)
        tv.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        tv.setShowGrid(False)
        tv.setWordWrap(False)
        tv.verticalHeader().setVisible(False)
        tv.verticalHeader().setDefaultSectionSize(30)
        tv.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        tv.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        hh = tv.horizontalHeader()
        hh.setHighlightSections(False)
        hh.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        for i, w in enumerate([56, 130, 0, 70, 66, 60, 130, 200]):
            if w:
                tv.setColumnWidth(i, w)
        for i in (0, 1, 3, 4, 5, 6, 7):
            hh.resizeSection(i, max(hh.sectionSize(i), hh.sectionSizeHint(i) + 12))
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        tv.setItemDelegateForColumn(2, HighlightDelegate(tv, lambda: THEMES[self.theme_name]))
        tv.doubleClicked.connect(self._open_hit)
        tv.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        tv.customContextMenuRequested.connect(self._table_menu)
        tv.selectionModel().currentRowChanged.connect(lambda cur, _prev: self._show_preview(cur))
        copy_act = QAction(self)
        copy_act.setShortcut(QKeySequence.StandardKey.Copy)
        copy_act.setShortcutContext(Qt.ShortcutContext.WidgetShortcut)
        copy_act.triggered.connect(self.copy_rows)
        tv.addAction(copy_act)
        self.table = tv
        self.empty_lbl = QLabel("", objectName="dropHint")
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setParent(tv.viewport())
        self.empty_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

        # row preview
        self.preview = QFrame(objectName="preview")
        pv = QVBoxLayout(self.preview)
        pv.setContentsMargins(12, 10, 12, 10)
        pv.setSpacing(6)
        ph = QHBoxLayout()
        self.preview_title = QLabel(tr("Row preview — select a result"), objectName="subtitle")
        ph.addWidget(self.preview_title, 1)
        self.preview_open = QPushButton(tr("Open at this cell"))
        self.preview_open.clicked.connect(lambda: self._open_hit(self.table.currentIndex()))
        self.preview_open.setEnabled(False)
        ph.addWidget(self.preview_open)
        pv.addLayout(ph)
        self.preview_table = QTableWidget(0, 3)
        self.preview_table.setHorizontalHeaderLabels([tr("Column"), tr("Header"), tr("Value")])
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.preview_table.setShowGrid(False)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setWordWrap(True)
        ph2 = self.preview_table.horizontalHeader()
        ph2.setHighlightSections(False)
        ph2.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.preview_table.setColumnWidth(0, 70)
        self.preview_table.setColumnWidth(1, 200)
        ph2.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        pv.addWidget(self.preview_table, 1)

        vsplit = QSplitter(Qt.Orientation.Vertical)
        vsplit.setChildrenCollapsible(True)
        vsplit.addWidget(tv)
        vsplit.addWidget(self.preview)
        vsplit.setStretchFactor(0, 3)
        vsplit.setStretchFactor(1, 1)
        vsplit.setSizes([480, 200])
        self.vsplit = vsplit
        rv.addWidget(vsplit, 1)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setVisible(False)
        rv.addWidget(self.progress)

        bottom = QHBoxLayout()
        self.status_lbl = QLabel(tr("Ready"), objectName="subtitle")
        bottom.addWidget(self.status_lbl, 1)
        self.preview_btn = QPushButton(tr("Row preview"))
        self.preview_btn.setCheckable(True)
        self.preview_btn.setChecked(True)
        self.preview_btn.toggled.connect(self.preview.setVisible)
        b_copy = QPushButton(tr("Copy"))
        b_copy.setToolTip(tr("Copy the selected rows (or all) — Ctrl+C"))
        b_copy.clicked.connect(self.copy_rows)
        b_csv = QPushButton(tr("Export CSV"))
        b_csv.clicked.connect(lambda: self.export("csv"))
        b_xlsx = QPushButton(tr("Excel report"))
        b_xlsx.setToolTip(tr("Checklist found / not found, all results, and one sheet per term"))
        b_xlsx.clicked.connect(lambda: self.export("xlsx"))
        b_clear = QPushButton(tr("Clear results"), objectName="danger")
        b_clear.clicked.connect(self.clear_results)
        for b in (self.preview_btn, b_copy, b_csv, b_xlsx, b_clear):
            bottom.addWidget(b)
        rv.addLayout(bottom)
        return right

    # ================= theme / language / persistence ================= #
    def apply_theme(self):
        t = THEMES[self.theme_name]
        QApplication.instance().setStyleSheet(build_qss(t))
        self.model.muted = QColor(t["muted"])
        self.model.warn = QColor(t["warn"])
        self.theme_btn.setText(tr("☀  Light") if self.theme_name == "dark" else tr("☾  Dark"))
        self._update_stats()
        self.table.viewport().update()

    def toggle_theme(self):
        self.theme_name = "light" if self.theme_name == "dark" else "dark"
        self.apply_theme()

    def toggle_lang(self):
        new = "en" if LANG == "el" else "el"
        if self.worker:
            return
        self.settings.setValue("lang", new)
        self.close()                       # saves everything first
        start_file(Path(__file__).resolve(), ["--no-update"])
        QTimer.singleShot(200, QApplication.quit)

    def _restore(self):
        s = self.settings
        geo = s.value("geometry")
        if geo is not None:
            self.restoreGeometry(geo)
        for key, w in (("splitter", self.splitter), ("vsplit", self.vsplit)):
            st = s.value(key)
            if st is not None:
                w.restoreState(st)
        self.terms_edit.setPlainText(s.value("terms", "") or "")
        self.folders = [f for f in slist(s.value("folders")) if os.path.isdir(f)]
        self.add_paths([f for f in slist(s.value("files")) if os.path.isfile(f)], quiet=True, remember=False)
        if self.folders:
            self.rescan_folders(auto=True, quiet=True)
        self._watch_all()
        mode = s.value("mode", "contains")
        mode = _OLD_MODE_LABELS.get(mode, mode)
        i = self.mode_cb.findData(mode)
        self.mode_cb.setCurrentIndex(max(0, i))
        self.fuzzy_spin.setValue(int(s.value("fuzzy", 85)))
        self.cb_case.setChecked(sbool(s.value("case"), False))
        self.cb_accents.setChecked(sbool(s.value("accents"), True))
        self.cb_greek.setChecked(sbool(s.value("greek"), False))
        self.cb_text_only.setChecked(sbool(s.value("text_only"), False))
        self.cb_one_per_row.setChecked(sbool(s.value("one_per_row"), False))
        self.cb_parallel.setChecked(sbool(s.value("parallel"), True))
        self.cb_cache.setChecked(sbool(s.value("cache"), True))
        self.cols_edit.setText(s.value("columns", "") or "")
        self.sheets_edit.setText(s.value("sheets", "") or "")
        self.preview_btn.setChecked(sbool(s.value("preview"), True))
        self.last_dir = s.value("last_dir", str(Path.home())) or str(Path.home())
        self._mode_changed()

    def closeEvent(self, e):
        if self.worker:
            self.worker.cancel()
        if self.thread:
            self.thread.quit()
            self.thread.wait(3000)
        s = self.settings
        s.setValue("geometry", self.saveGeometry())
        s.setValue("splitter", self.splitter.saveState())
        s.setValue("vsplit", self.vsplit.saveState())
        s.setValue("theme", self.theme_name)
        s.setValue("terms", self.terms_edit.toPlainText())
        folder_files = set(self._files_in_folders())
        s.setValue("files", [f for f in self.file_paths() if f not in folder_files])
        s.setValue("folders", self.folders)
        s.setValue("mode", self.mode_cb.currentData())
        s.setValue("fuzzy", self.fuzzy_spin.value())
        s.setValue("case", self.cb_case.isChecked())
        s.setValue("accents", self.cb_accents.isChecked())
        s.setValue("greek", self.cb_greek.isChecked())
        s.setValue("text_only", self.cb_text_only.isChecked())
        s.setValue("one_per_row", self.cb_one_per_row.isChecked())
        s.setValue("parallel", self.cb_parallel.isChecked())
        s.setValue("cache", self.cb_cache.isChecked())
        s.setValue("columns", self.cols_edit.text())
        s.setValue("sheets", self.sheets_edit.text())
        s.setValue("preview", self.preview_btn.isChecked())
        s.setValue("last_dir", self.last_dir)
        s.sync()
        super().closeEvent(e)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        QTimer.singleShot(0, self._place_overlays)

    def showEvent(self, e):
        super().showEvent(e)
        QTimer.singleShot(0, self._place_overlays)

    def _place_overlays(self):
        self.drop_hint.setGeometry(self.file_list.viewport().rect())
        self.empty_lbl.setGeometry(self.table.viewport().rect())

    def _mode_changed(self):
        fuzzy = self.mode_cb.currentData() == "fuzzy"
        self.fuzzy_spin.setVisible(fuzzy)
        regex = self.mode_cb.currentData() == "regex"
        self.cb_greek.setEnabled(not regex)

    def set_status(self, text: str):
        self.status_lbl.setText(text)
        self.status_lbl.setToolTip(text)

    # ================= files & folders ================= #
    def file_paths(self) -> list[str]:
        return [self.file_list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.file_list.count())]

    def _add_file_item(self, f: str):
        ext = Path(f).suffix.lower()
        icon = "📕" if ext in PDF_EXTS else ("📘" if ext in DOC_EXTS else "📗")
        it = QListWidgetItem(f"{icon}  {os.path.basename(f)}")
        it.setData(Qt.ItemDataRole.UserRole, f)
        it.setToolTip(f)
        self.file_list.addItem(it)

    @staticmethod
    def _wanted(f: str) -> bool:
        name = os.path.basename(f)
        return Path(f).suffix.lower() in ALL_EXTS and not name.startswith("~$")

    @staticmethod
    def _walk(folder: str) -> list[str]:
        out = []
        for dp, dns, fns in os.walk(folder):
            dns[:] = [d for d in dns if not d.startswith(".")]
            out += [os.path.join(dp, f) for f in sorted(fns)]
        return out

    def _files_in_folders(self) -> list[str]:
        out = []
        for d in self.folders:
            out += [os.path.abspath(f) for f in self._walk(d) if self._wanted(f)]
        return out

    def add_paths(self, paths: list[str], quiet=False, remember=True):
        existing = {os.path.normcase(p) for p in self.file_paths()}
        added = 0
        for p in paths:
            if os.path.isdir(p):
                p = os.path.abspath(p)
                if remember:
                    self._remember_folder(p)
                if p not in self.folders:
                    self.folders.append(p)
                cands = self._walk(p)
            else:
                cands = [p]
            for f in cands:
                if not self._wanted(f):
                    continue
                f = os.path.abspath(f)
                key = os.path.normcase(f)
                if key in existing:
                    continue
                existing.add(key)
                self._add_file_item(f)
                added += 1
        if not quiet:
            self.set_status(tr("Added {n} file(s)").format(n=added) if added else tr("No new files found"))
        self._watch_all()
        self.update_counts()

    def _remember_folder(self, folder: str):
        rec = [f for f in slist(self.settings.value("recent_folders")) if os.path.normcase(f) != os.path.normcase(folder)]
        rec.insert(0, folder)
        self.settings.setValue("recent_folders", rec[:10])

    def _fill_recent_menu(self):
        m = self.recent_menu
        m.clear()
        rec = [f for f in slist(self.settings.value("recent_folders")) if os.path.isdir(f)]
        if not rec:
            a = m.addAction(tr("No recent folders yet"))
            a.setEnabled(False)
            return
        for f in rec:
            m.addAction(f"📁  {f}", lambda f=f: self.add_paths([f]))
        m.addSeparator()
        m.addAction(tr("Forget the list"), lambda: self.settings.setValue("recent_folders", []))

    def _watch_all(self):
        old = self.watcher.directories()
        if old:
            self.watcher.removePaths(old)
        dirs = []
        for d in self.folders:
            for dp, dns, _fns in os.walk(d):
                dns[:] = [x for x in dns if not x.startswith(".")]
                dirs.append(dp)
                if len(dirs) >= 400:
                    break
        if dirs:
            self.watcher.addPaths(dirs)
        n = len(self.folders)
        self.watch_btn.setText(tr("👁 Watching {n} folder(s) ▾").format(n=n) if n else tr("👁 No watched folders"))
        self.watch_btn.setEnabled(bool(n))

    def _fill_watch_menu(self):
        m = self.watch_menu
        m.clear()
        m.addAction(tr("Refresh now (F5)"), lambda: self.rescan_folders(auto=False))
        m.addSeparator()
        for f in self.folders:
            m.addAction(tr("Stop watching: {f}").format(f=f), lambda f=f: self._unwatch(f))

    def _unwatch(self, folder: str):
        self.folders = [f for f in self.folders if f != folder]
        self._watch_all()

    def rescan_folders(self, auto=False, quiet=False):
        """New files in the watched folders are added, deleted ones removed from the list."""
        if self.worker:
            self._rescan_timer.start()
            return
        before = {os.path.normcase(p) for p in self.file_paths()}
        removed = 0
        for i in reversed(range(self.file_list.count())):
            if not os.path.exists(self.file_list.item(i).data(Qt.ItemDataRole.UserRole)):
                self.file_list.takeItem(i)
                removed += 1
        added = 0
        for f in self._files_in_folders():
            if os.path.normcase(f) not in before:
                self._add_file_item(f)
                added += 1
        self._watch_all()
        self.update_counts()
        if not quiet and (added or removed or not auto):
            self.set_status(tr("Folders checked: {a} new, {r} removed").format(a=added, r=removed))

    def pick_files(self):
        files, _f = QFileDialog.getOpenFileNames(
            self, tr("Add files"), self.last_dir,
            tr("Supported files") + " (*.xlsx *.xlsm *.xltx *.xltm *.xls *.csv *.docx *.pdf);;"
            + tr("All files") + " (*)")
        if files:
            self.last_dir = os.path.dirname(files[0])
            self.add_paths(files)

    def pick_folder(self):
        d = QFileDialog.getExistingDirectory(self, tr("Add a folder (with its subfolders)"), self.last_dir)
        if d:
            self.last_dir = d
            self.add_paths([d])

    def remove_selected_files(self):
        for it in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(it))
        self.update_counts()

    def clear_files(self):
        self.file_list.clear()
        self.folders = []
        self._watch_all()
        self.update_counts()

    # ================= terms & lists ================= #
    def terms(self) -> list[str]:
        seen, out = set(), []
        for line in self.terms_edit.toPlainText().splitlines():
            s = line.strip()
            if s and s not in seen:
                seen.add(s)
                out.append(s)
        return out

    def update_counts(self):
        n = self.file_list.count()
        self.files_count.setText(tr("{n} files").format(n=n))
        self.drop_hint.setVisible(n == 0)
        k = len(self.terms())
        self.terms_count.setText(tr("{n} terms").format(n=k))

    def _fill_lists_menu(self):
        m = self.lists_menu
        m.clear()
        lists = load_term_lists()
        m.addAction(tr("💾  Save these terms as a list…"), self._save_list)
        load = m.addMenu(tr("📂  Load a list"))
        dele = m.addMenu(tr("🗑  Delete a list"))
        for name in sorted(lists, key=str.casefold):
            load.addAction(f"{name}  ({len(lists[name])})", lambda n=name: self._load_list(n))
            dele.addAction(name, lambda n=name: self._delete_list(n))
        if not lists:
            for sub in (load, dele):
                a = sub.addAction(tr("No saved lists yet"))
                a.setEnabled(False)
        m.addSeparator()
        m.addAction(tr("📊  Import from an Excel column…"), self._import_column)
        m.addAction(tr("📄  Import from a text file…"), self._import_text)

    def _save_list(self):
        terms = self.terms()
        if not terms:
            QMessageBox.information(self, APP_NAME, tr("Type or paste some terms first."))
            return
        name, ok = QInputDialog.getText(self, tr("Save list"), tr("Name of the list:"))
        name = (name or "").strip()
        if not ok or not name:
            return
        lists = load_term_lists()
        if name in lists and QMessageBox.question(
                self, APP_NAME, tr("A list named “{n}” exists. Replace it?").format(n=name)) \
                != QMessageBox.StandardButton.Yes:
            return
        lists[name] = terms
        save_term_lists(lists)
        self.set_status(tr("Saved list “{n}” ({k} terms)").format(n=name, k=len(terms)))

    def _load_list(self, name: str):
        terms = load_term_lists().get(name, [])
        self.terms_edit.setPlainText("\n".join(terms))
        self.set_status(tr("Loaded list “{n}” ({k} terms)").format(n=name, k=len(terms)))

    def _delete_list(self, name: str):
        if QMessageBox.question(self, APP_NAME, tr("Delete the list “{n}”?").format(n=name)) \
                != QMessageBox.StandardButton.Yes:
            return
        lists = load_term_lists()
        lists.pop(name, None)
        save_term_lists(lists)

    def _set_terms(self, terms: list[str], mode: str):
        if mode == "add":
            terms = self.terms() + [t for t in terms if t not in set(self.terms())]
        self.terms_edit.setPlainText("\n".join(terms))

    def _import_column(self):
        f, _f = QFileDialog.getOpenFileName(self, tr("Spreadsheet with the terms"), self.last_dir,
                                            tr("Spreadsheets") + " (*.xlsx *.xlsm *.xls *.csv)")
        if not f:
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            data, _c = load_file(f, str(_cache_dir()))
        except Exception as e:  # noqa: BLE001
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, APP_NAME, tr("Could not read the file:") + f"\n{e}")
            return
        QApplication.restoreOverrideCursor()
        dlg = ColumnImportDialog(self, f, data)
        if dlg.exec() and dlg.result_terms:
            self._set_terms(dlg.result_terms, dlg.mode)
            self.set_status(tr("Imported {k} terms").format(k=len(dlg.result_terms)))

    def _import_text(self):
        f, _f = QFileDialog.getOpenFileName(self, tr("Text file with one term per line"), self.last_dir,
                                            tr("Text files") + " (*.txt *.csv);;" + tr("All files") + " (*)")
        if not f:
            return
        raw = Path(f).read_bytes()
        for enc in ("utf-8-sig", "cp1253", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        terms = [ln.strip() for ln in text.splitlines() if ln.strip()]
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setText(tr("{k} terms in the file. Replace your terms or add to them?").format(k=len(terms)))
        rep = box.addButton(tr("Replace terms"), QMessageBox.ButtonRole.AcceptRole)
        add = box.addButton(tr("Add to terms"), QMessageBox.ButtonRole.AcceptRole)
        box.addButton(tr("Cancel"), QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() in (rep, add):
            self._set_terms(terms, "replace" if box.clickedButton() is rep else "add")

    # ================= search ================= #
    def _params(self, terms: list[str]) -> dict:
        return dict(
            terms=terms, mode=self.mode_cb.currentData(), case=self.cb_case.isChecked(),
            accents=self.cb_accents.isChecked(), greek=self.cb_greek.isChecked() and
            self.mode_cb.currentData() != "regex", fuzzy=self.fuzzy_spin.value(),
            text_only=self.cb_text_only.isChecked(), one_per_row=self.cb_one_per_row.isChecked(),
            cols=parse_columns(self.cols_edit.text()), sheets=parse_sheets(self.sheets_edit.text()),
            cache_dir=str(_cache_dir()) if self.cb_cache.isChecked() else None,
        )

    def toggle_search(self):
        if self.worker:
            self.worker.cancel()
            self.set_status(tr("Cancelling…"))
            return
        files = [f for f in self.file_paths() if os.path.exists(f)]
        terms = self.terms()
        if not files:
            QMessageBox.information(self, APP_NAME, tr("Step 1: add at least one file or folder."))
            return
        if not terms:
            QMessageBox.information(self, APP_NAME, tr("Step 2: type at least one search term (one per line)."))
            self.terms_edit.setFocus()
            return
        if self.mode_cb.currentData() == "regex":
            for t in terms:
                try:
                    re.compile(t)
                except re.error as e:
                    QMessageBox.warning(self, APP_NAME, tr("Invalid regular expression:") + f"\n\n{t}\n\n{e}")
                    return
        self.clear_results()
        self.last_terms = terms
        self.errors = []
        self.worker = SearchWorker(files, self._params(terms), self.cb_parallel.isChecked())
        self.thread = QThread(self)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.hits.connect(self._on_hits)
        self.worker.file_error.connect(lambda p, e: self.errors.append((p, e)))
        self.worker.finished.connect(self._on_finished)
        self.table.setSortingEnabled(False)
        self.search_btn.setText(tr("Cancel  (Esc)"))
        self.progress.setRange(0, len(files))
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.thread.start()

    def _escape(self):
        if self.worker:
            self.worker.cancel()
        elif self.filter_edit.hasFocus():
            self.filter_edit.clear()

    def _on_progress(self, done, total, msg):
        self.progress.setValue(done)
        self.set_status(f"[{done}/{total}]  {msg}  ·  " + tr("{n} matches").format(n=len(self.model.rows)))

    def _on_hits(self, hits):
        self.model.add(hits)
        self._update_stats()
        if self.table.currentIndex().row() < 0 and self.proxy.rowCount():
            self.table.selectRow(0)

    def _on_finished(self, secs, cancelled, cached, procs):
        self.thread.quit()
        self.thread.wait()
        self.thread = None
        self.worker = None
        self.table.setSortingEnabled(True)
        self.search_btn.setText(tr("4 · Search"))
        self.progress.setVisible(False)
        n = len(self.model.rows)
        extra = []
        if procs:
            extra.append(tr("{p} processes").format(p=procs))
        if cached:
            extra.append(tr("{c} from memory").format(c=cached))
        self.set_status((tr("Cancelled") if cancelled else tr("Finished")) + f" — {secs:.2f}s · "
                        + tr("{n} matches").format(n=n) + (f"  ({', '.join(extra)})" if extra else ""))
        self._update_stats(final=True)
        if self.cb_cache.isChecked():
            threading.Thread(target=prune_cache, args=(_cache_dir(),), daemon=True).start()

    def _update_stats(self, final=False):
        t = THEMES[self.theme_name]
        rows = self.model.rows
        n = len(rows)
        shown = self.proxy.rowCount()
        self.stat_hits.setText(tr("{n} matches").format(n=n) if shown == n
                               else tr("{s} of {n} matches").format(s=shown, n=n))
        nf = len({h.path for h in rows})
        self.stat_files.setText(tr("{n} files").format(n=nf))
        if final and self.last_terms:
            found = {h.term for h in rows}
            missing = [x for x in self.last_terms if x not in found]
            self.stat_missing.setVisible(bool(missing))
            self.stat_missing.setText(tr("⚠ {n} not found").format(n=len(missing)))
            self.stat_missing.setToolTip(tr("Click to see which terms had no match"))
        self.stat_missing.setStyleSheet(f"color: {t['danger']};")
        self.stat_errors.setVisible(bool(self.errors))
        if self.errors:
            self.stat_errors.setText(tr("✖ {n} unreadable").format(n=len(self.errors)))
        self.stat_errors.setStyleSheet(f"color: {t['danger']};")
        self._update_empty()

    def _update_empty(self):
        self.empty_lbl.setVisible(self.proxy.rowCount() == 0)
        if self.model.rows and self.proxy.rowCount() == 0:
            self.empty_lbl.setText(tr("No results match the filter."))
        elif not self.model.rows and self.last_terms and not self.worker:
            self.empty_lbl.setText(tr("Nothing found."))
        elif not self.model.rows:
            self.empty_lbl.setText(tr("No results yet.\n\n1 · add files   2 · type terms   3 · options   4 · Search"))

    def _on_filter(self, s):
        self.proxy.set_needle(s)
        self._update_stats()

    def show_missing(self):
        found = {h.term for h in self.model.rows}
        missing = [t for t in self.last_terms if t not in found]
        box = QMessageBox(self)
        box.setWindowTitle(tr("Terms not found"))
        box.setText(tr("{m} of {n} terms had no match:").format(m=len(missing), n=len(self.last_terms)))
        box.setDetailedText("\n".join(missing))
        copy = box.addButton(tr("Copy list"), QMessageBox.ButtonRole.ActionRole)
        fuzzy = box.addButton(tr("Search them as Similar"), QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Close)
        box.exec()
        if box.clickedButton() is copy:
            QGuiApplication.clipboard().setText("\n".join(missing))
        elif box.clickedButton() is fuzzy:
            self.terms_edit.setPlainText("\n".join(missing))
            self.mode_cb.setCurrentIndex(self.mode_cb.findData("fuzzy"))
            self.cb_greek.setChecked(True)
            self.toggle_search()

    def show_errors(self):
        txt = "\n\n".join(f"{p}\n   {e}" for p, e in self.errors)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle(tr("Unreadable files"))
        box.setText(tr("{n} file(s) could not be read (damaged, password-protected, scanned PDF, or open "
                       "in another program).").format(n=len(self.errors)))
        box.setDetailedText(txt)
        box.exec()

    def clear_results(self):
        self.model.clear()
        self.last_terms = []
        self.errors = []
        self.stat_missing.setVisible(False)
        self.stat_errors.setVisible(False)
        self._show_preview(QModelIndex())
        self._update_stats()

    # ================= results ================= #
    def _hit_at(self, proxy_index) -> Hit | None:
        if not proxy_index.isValid():
            return None
        return self.model.rows[self.proxy.mapToSource(proxy_index).row()]

    def _open_hit(self, idx):
        h = self._hit_at(idx)
        if h:
            open_at_cell(h, self.set_status)

    def _show_preview(self, idx):
        h = self._hit_at(idx) if idx is not None else None
        pt = self.preview_table
        pt.setRowCount(0)
        self.preview_open.setEnabled(h is not None)
        if h is None:
            self.preview_title.setText(tr("Row preview — select a result"))
            return
        if h.kind == "sheet":
            self.preview_title.setText(tr("Row {r}  ·  {s}  ·  {f}").format(r=h.row, s=h.sheet, f=h.file_name))
            self.preview_open.setText(tr("Open at {c}").format(c=h.cell))
            vals = h.row_values or (h.value,)
            pt.setRowCount(len(vals))
            for i, v in enumerate(vals):
                c = i + 1
                head = h.header[i] if i < len(h.header) else ""
                items = [QTableWidgetItem(col_letter(c)), QTableWidgetItem(head), QTableWidgetItem(v)]
                for j, it in enumerate(items):
                    if c == h.col:
                        f = it.font()
                        f.setBold(True)
                        it.setFont(f)
                    pt.setItem(i, j, it)
            pt.resizeRowsToContents()
            if 0 < h.col <= len(vals):
                pt.selectRow(h.col - 1)
                pt.scrollToItem(pt.item(h.col - 1, 0), QAbstractItemView.ScrollHint.PositionAtCenter)
        else:
            where = (tr("Paragraph {r}") if h.kind == "doc" else tr("Page {p}, line {r}")).format(
                r=h.row, p=h.sheet)
            self.preview_title.setText(f"{where}  ·  {h.file_name}")
            self.preview_open.setText(tr("Open file"))
            pt.setRowCount(1)
            pt.setItem(0, 0, QTableWidgetItem(h.cell))
            pt.setItem(0, 1, QTableWidgetItem(""))
            pt.setItem(0, 2, QTableWidgetItem(h.value))
            pt.resizeRowsToContents()

    def _selected_hits(self) -> list[Hit]:
        rows = sorted({i.row() for i in self.table.selectionModel().selectedRows()})
        return [self.model.rows[self.proxy.mapToSource(self.proxy.index(r, 0)).row()] for r in rows]

    def _visible_hits(self) -> list[Hit]:
        return [self.model.rows[self.proxy.mapToSource(self.proxy.index(r, 0)).row()]
                for r in range(self.proxy.rowCount())]

    def _table_menu(self, pos):
        h = self._hit_at(self.table.indexAt(pos))
        if not h:
            return
        m = QMenu(self)
        if h.kind == "sheet":
            m.addAction(tr("Open in Excel at {c}").format(c=f"{h.sheet}!{h.cell}"),
                        lambda: open_at_cell(h, self.set_status))
        m.addAction(tr("Open file"), lambda: open_path(h.path))
        m.addAction(tr("Show in folder"), lambda: reveal_in_folder(h.path))
        m.addSeparator()
        m.addAction(tr("Copy found value"), lambda: QGuiApplication.clipboard().setText(h.value))
        m.addAction(tr("Copy whole row"), lambda: QGuiApplication.clipboard().setText("\t".join(h.row_values)))
        m.addAction(tr("Copy file path"), lambda: QGuiApplication.clipboard().setText(h.path))
        m.addAction(tr("Copy selected rows"), self.copy_rows)
        m.addSeparator()
        m.addAction(tr("Filter by term “{t}”").format(t=h.term), lambda: self.filter_edit.setText(h.term))
        m.addAction(tr("Filter by file “{f}”").format(f=h.file_name), lambda: self.filter_edit.setText(h.file_name))
        m.exec(self.table.viewport().mapToGlobal(pos))

    def copy_rows(self):
        hits = self._selected_hits() or self._visible_hits()
        if not hits:
            return
        lines = ["\t".join([tr("Term"), tr("Found value"), tr("Match"), tr("Cell"), tr("Row"), tr("Sheet / page"),
                            tr("File"), tr("Path")])]
        for h in hits:
            lines.append("\t".join([h.term, h.value.replace("\t", " ").replace("\n", " "), f"{h.score}%",
                                    h.cell, str(h.row), h.sheet_display, h.file_name, h.path]))
        QGuiApplication.clipboard().setText("\n".join(lines))
        self.set_status(tr("Copied {n} row(s)").format(n=len(hits)))

    def export(self, kind: str):
        hits = self._visible_hits()
        if not hits and not self.last_terms:
            QMessageBox.information(self, APP_NAME, tr("There are no results to export."))
            return
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M")
        default = os.path.join(self.last_dir, f"{tr('search_report')}_{stamp}.{kind}")
        flt = "CSV (*.csv)" if kind == "csv" else tr("Excel workbook") + " (*.xlsx)"
        path, _f = QFileDialog.getSaveFileName(self, tr("Export results"), default, flt)
        if not path:
            return
        header = ["#", tr("Term"), tr("Found value"), tr("Match"), tr("Cell"), tr("Row"), tr("Column"),
                  tr("Sheet / page"), tr("File"), tr("Path")]
        data = [[i, h.term, h.value, h.score / 100, h.cell, h.row, h.col, h.sheet_display, h.file_name, h.path]
                for i, h in enumerate(hits, 1)]
        try:
            if kind == "csv":
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f, delimiter=";")
                    w.writerow(header)
                    for r in data:
                        w.writerow(r[:3] + [f"{round(r[3] * 100)}%"] + r[4:])
            else:
                self._write_report(path, header, data, hits)
        except PermissionError:
            QMessageBox.warning(self, APP_NAME, tr("Could not write the file — is it open in Excel?"))
            return
        except Exception as e:  # noqa: BLE001
            QMessageBox.warning(self, APP_NAME, tr("Export failed:") + f"\n{e}")
            return
        self.last_dir = os.path.dirname(path)
        self.set_status(tr("Exported {n} rows → {f}").format(n=len(hits), f=os.path.basename(path)))
        if QMessageBox.question(self, APP_NAME, tr("Export complete. Open the file now?")) \
                == QMessageBox.StandardButton.Yes:
            open_path(path)

    def _write_report(self, path, header, data, hits):
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        head_fill = PatternFill("solid", fgColor="1F4E79")
        yes_fill = PatternFill("solid", fgColor="C6EFCE")
        no_fill = PatternFill("solid", fgColor="FFC7CE")
        white_b = Font(bold=True, color="FFFFFF")

        def style_head(ws, widths):
            for c in ws[1]:
                c.font = white_b
                c.fill = head_fill
                c.alignment = Alignment(vertical="center")
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = w
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        wb = Workbook()
        used = set()
        # 1. checklist: every term, found or not
        ck = wb.active
        ck.title = safe_sheet_title(tr("Checklist"), used)
        ck.append([tr("Term"), tr("Found"), tr("Matches"), tr("Files"), tr("First location")])
        by_term: dict[str, list[Hit]] = {}
        for h in hits:
            by_term.setdefault(h.term, []).append(h)
        terms = self.last_terms or list(by_term)
        for t in terms:
            hs = by_term.get(t, [])
            first = f"{hs[0].file_name} › {hs[0].sheet_display} › {hs[0].cell}" if hs else ""
            ck.append([t, tr("YES") if hs else tr("NO"), len(hs), len({h.path for h in hs}), first])
            cell = ck.cell(row=ck.max_row, column=2)
            cell.fill = yes_fill if hs else no_fill
            cell.font = Font(bold=True, color="006100" if hs else "9C0006")
            cell.alignment = Alignment(horizontal="center")
        style_head(ck, [40, 10, 10, 8, 70])
        ck.append([])
        found_n = sum(1 for t in terms if by_term.get(t))
        ck.append([tr("Found {f} of {n} terms").format(f=found_n, n=len(terms))])
        ck.cell(row=ck.max_row, column=1).font = Font(bold=True)

        # 2. all results
        ws = wb.create_sheet(safe_sheet_title(tr("All results"), used))
        ws.append(header)
        for r in data:
            ws.append(r)
            ws.cell(row=ws.max_row, column=4).number_format = "0%"
        style_head(ws, [6, 22, 50, 8, 8, 8, 8, 18, 30, 60])

        # 3. one sheet per term that was found
        for t in terms[:200]:
            hs = by_term.get(t)
            if not hs:
                continue
            s = wb.create_sheet(safe_sheet_title(t, used))
            s.append([tr("Found value"), tr("Match"), tr("Cell"), tr("Sheet / page"), tr("File"), tr("Whole row")])
            for h in hs:
                s.append([h.value, h.score / 100, h.cell, h.sheet_display, h.file_name, " | ".join(h.row_values)])
                s.cell(row=s.max_row, column=2).number_format = "0%"
            style_head(s, [50, 8, 8, 18, 30, 90])
        wb.save(path)


# --------------------------------------------------------------------------- #
#  Startup splash (every launch): components check / install, then the app
# --------------------------------------------------------------------------- #

SPLASH_MIN_STEP = 0.45    # seconds each step stays on screen, so it can be read


class SplashScreen(QWidget):
    def __init__(self):
        super().__init__(None, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint |
                         Qt.WindowType.SplashScreen)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(520, 300)
        self.shown_p, self.target_p, self.phase = 0.0, 0.0, 0.0
        self.status, self.detail = tr("Starting"), ""
        self.logo = make_app_icon().pixmap(64, 64)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(16)
        scr = QApplication.primaryScreen().availableGeometry()
        self.move(scr.center().x() - self.width() // 2, scr.center().y() - self.height() // 2)

    def set_step(self, p, status, detail=""):
        self.target_p = max(self.target_p, p)
        self.status, self.detail = status, detail
        self.update()

    def _tick(self):
        self.shown_p += (self.target_p - self.shown_p) * 0.12
        self.phase = (self.phase + 0.012) % 1.0
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(8, 8, -8, -8)
        path = QPainterPath()
        path.addRoundedRect(r, 20, 20)
        p.fillPath(path, QColor("#0f1115"))
        g = QRadialGradient(QPointF(r.center().x(), r.top()), r.width() * 0.55)
        g.setColorAt(0, QColor(22, 163, 74, 70))
        g.setColorAt(1, QColor(22, 163, 74, 0))
        p.save()
        p.setClipPath(path)
        p.fillRect(r, g)
        p.restore()
        p.setPen(QPen(QColor("#2a2f3a"), 1))
        p.drawPath(path)
        p.drawPixmap(int(r.center().x() - 28), int(r.top() + 30), 56, 56, self.logo)
        f = QFont(self.font())
        f.setPointSizeF(19)
        f.setWeight(QFont.Weight.DemiBold)
        p.setFont(f)
        p.setPen(QColor("#e6e8ee"))
        p.drawText(QRectF(r.left(), r.top() + 96, r.width(), 36), Qt.AlignmentFlag.AlignCenter, APP_NAME)
        f.setPointSizeF(9.5)
        f.setWeight(QFont.Weight.Normal)
        p.setFont(f)
        p.setPen(QColor("#8b93a7"))
        p.drawText(QRectF(r.left(), r.top() + 132, r.width(), 20), Qt.AlignmentFlag.AlignCenter,
                   tr("Version {v}  ·  Excel, CSV, Word and PDF search").format(v=APP_VERSION))
        bar = QRectF(r.left() + 56, r.top() + 180, r.width() - 112, 6)
        bp = QPainterPath()
        bp.addRoundedRect(bar, 3, 3)
        p.fillPath(bp, QColor("#1e222b"))
        fill = QRectF(bar.left(), bar.top(), max(6.0, bar.width() * self.shown_p), bar.height())
        fp = QPainterPath()
        fp.addRoundedRect(fill, 3, 3)
        p.fillPath(fp, QColor("#3b82f6"))
        sx = fill.left() + (fill.width() + 60) * self.phase - 60
        lg = QLinearGradient(sx, 0, sx + 60, 0)
        lg.setColorAt(0, QColor(255, 255, 255, 0))
        lg.setColorAt(0.5, QColor(255, 255, 255, 90))
        lg.setColorAt(1, QColor(255, 255, 255, 0))
        p.save()
        p.setClipPath(fp)
        p.fillRect(fill, lg)
        p.restore()
        f.setPointSizeF(10)
        f.setWeight(QFont.Weight.DemiBold)
        p.setFont(f)
        p.setPen(QColor("#e6e8ee"))
        p.drawText(QRectF(bar.left(), bar.bottom() + 12, bar.width() - 50, 20),
                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                   self.status + "…" * (self.shown_p < 0.995))
        p.setPen(QColor("#8b93a7"))
        p.drawText(QRectF(bar.left(), bar.bottom() + 12, bar.width(), 20),
                   Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, f"{int(self.shown_p * 100)}%")
        f.setPointSizeF(8.5)
        f.setWeight(QFont.Weight.Normal)
        p.setFont(f)
        p.setPen(QColor("#5d636f"))
        det = p.fontMetrics().elidedText(self.detail, Qt.TextElideMode.ElideMiddle, int(bar.width()))
        p.drawText(QRectF(bar.left(), bar.bottom() + 32, bar.width(), 18),
                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, det)
        p.drawText(QRectF(r.left(), r.bottom() - 30, r.width(), 18), Qt.AlignmentFlag.AlignCenter,
                   "by Nikolaos K. Paridis  ·  ALU DEV TEAM")
        p.end()


def _wait(app, seconds, until=None):
    """Keep the splash animating for at least `seconds` (and until `until()` is true)."""
    t_end = time.time() + seconds
    while time.time() < t_end or (until is not None and not until()):
        app.processEvents()
        time.sleep(0.01)


def startup(app) -> bool:
    """Splash: check / install components, shortcuts. False = cannot continue."""
    sp = SplashScreen()
    sp.show()
    sp.set_step(0.15, tr("Starting"), f"Python {sys.version.split()[0]}")
    _wait(app, SPLASH_MIN_STEP)

    sp.set_step(0.35, tr("Checking components"), ", ".join(PACKAGES[1:]))
    _wait(app, SPLASH_MIN_STEP)
    missing = _missing_packages()
    if missing:
        sp.set_step(0.45, tr("Installing components (one time)"), " ".join(missing))
        state = {"rc": None, "out": []}

        def work():
            try:
                pr = subprocess.Popen(_pip_cmd(missing), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                      text=True, creationflags=NO_WINDOW)
                for ln in pr.stdout:
                    state["out"].append(ln)
                    sp.detail = ln.strip()[:90]
                state["rc"] = pr.wait()
            except Exception as e:  # noqa: BLE001
                state["out"].append(str(e))
                state["rc"] = 1

        threading.Thread(target=work, daemon=True).start()
        _wait(app, 0.5, until=lambda: state["rc"] is not None)
        _add_user_site()
        still = _missing_packages()
        if still:
            sp.close()
            QMessageBox.critical(None, APP_NAME,
                                 tr("Could not install: {p}\n\nCheck the internet connection and start the app "
                                    "again.").format(p=", ".join(still)) + "\n\n" + "".join(state["out"])[-1000:])
            return False
        sp.set_step(0.7, tr("Components installed"), " ".join(missing))
        _wait(app, SPLASH_MIN_STEP)

    sp.set_step(0.8, tr("Preparing shortcuts"), tr("Desktop and Start menu"))
    try:
        make_shortcuts()
    except Exception:
        pass
    _wait(app, SPLASH_MIN_STEP)
    sp.set_step(1.0, tr("Ready"), "")
    _wait(app, SPLASH_MIN_STEP + 0.3)
    sp.close()
    return True


def _app_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") if os.name == "nt" else None
    d = (Path(base) / "ExcelUltimateSearch") if base else (Path.home() / ".excel_ultimate_search")
    d.mkdir(parents=True, exist_ok=True)
    return d


def _cache_dir() -> Path:
    d = _app_data_dir() / "cache"
    d.mkdir(parents=True, exist_ok=True)
    return d


def make_shortcuts(force=False) -> bool:
    """Windows: Desktop + Start menu shortcuts that start this file with pythonw.exe (no console)."""
    if os.name != "nt":
        return False
    script = str(Path(__file__).resolve())
    st = QSettings(SETTINGS_ORG, SETTINGS_APP)
    if not force and st.value("shortcut_for", "") == script:
        return True
    ico = _app_data_dir() / "icon.ico"
    try:
        make_app_icon().pixmap(256, 256).save(str(ico), "ICO")
    except Exception:
        pass
    ps = ("$sh = New-Object -ComObject WScript.Shell; "
          "foreach ($dir in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) { "
          "if (-not $dir) { continue }; $l = $sh.CreateShortcut((Join-Path $dir 'Excel Ultimate Search.lnk')); "
          "$l.TargetPath = $env:EUS_PYW; $l.Arguments = '\"' + $env:EUS_SCRIPT + '\"'; "
          "$l.WorkingDirectory = $env:EUS_DIR; $l.Description = 'Excel Ultimate Search'; "
          "if (Test-Path $env:EUS_ICO) { $l.IconLocation = $env:EUS_ICO + ',0' }; $l.Save() }")
    env = dict(os.environ, EUS_PYW=_gui_python(), EUS_SCRIPT=script, EUS_DIR=str(Path(script).parent),
               EUS_ICO=str(ico))
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden",
                            "-Command", ps], env=env, capture_output=True, text=True, timeout=30,
                           creationflags=NO_WINDOW)
        ok = r.returncode == 0
    except Exception:
        ok = False
    if ok:
        st.setValue("shortcut_for", script)
    return ok


# --------------------------------------------------------------------------- #
#  App updates (OTA) from GitHub
# --------------------------------------------------------------------------- #

APP_REPO = "nikkpap/Excel_Ultimate_Search"
APP_PAGE = f"https://github.com/{APP_REPO}"
APP_UPDATE_URL = f"https://raw.githubusercontent.com/{APP_REPO}/main/{APP_FILE_STEM}.pyw"
APP_CHANGES_URL = f"https://raw.githubusercontent.com/{APP_REPO}/main/CHANGELOG.md"


def _http(url, timeout=15):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": f"ExcelUltimateSearch/{APP_VERSION}",
                                               "Cache-Control": "no-cache"})
    return urllib.request.urlopen(req, timeout=timeout)


def _ver(v) -> tuple:
    try:
        return tuple(int(x) for x in str(v).strip().split("."))
    except ValueError:
        return (0,)


def app_update_check():
    """(version, text, notes) of a newer release on GitHub, or None when this is the newest."""
    with _http(f"{APP_UPDATE_URL}?t={int(time.time())}", timeout=20) as r:
        text = r.read().decode("utf-8")
    m = re.search(r'^APP_VERSION = "([0-9.]+)"', text, re.M)
    if not m:
        raise RuntimeError("the file on GitHub has no version")
    ver = m.group(1)
    if _ver(ver) <= _ver(APP_VERSION):
        return None
    compile(text, f"{APP_FILE_STEM}_v{ver}.pyw", "exec")   # never start a half-downloaded file
    notes = ""
    try:
        with _http(APP_CHANGES_URL, timeout=10) as r:
            ch = r.read().decode("utf-8")
        parts = re.split(r"^## ", ch, flags=re.M)
        news = [x for x in parts[1:] if x.split() and _ver(x.split()[0]) > _ver(APP_VERSION)]
        notes = "\n\n".join("• " + x.strip().replace("\n\n", "\n") for x in news[:6])
    except Exception:
        pass
    return ver, text, notes


def app_update_install(ver: str, text: str) -> Path:
    """Saves <stem>_v<ver>.pyw next to this file (the old one stays)."""
    dest = HERE / f"{APP_FILE_STEM}_v{ver}.pyw"
    tmp = dest.with_suffix(".pyw.part")
    tmp.write_text(text, encoding="utf-8", newline="\n")
    os.replace(tmp, dest)
    return dest


def start_file(path: Path, args: list | None = None):
    flags = (DETACHED | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)) if os.name == "nt" else 0
    subprocess.Popen([_gui_python(), str(path)] + (args or []), cwd=str(path.parent), close_fds=True,
                     creationflags=flags)


class _Bridge(QObject):
    """Carries a background thread's result back to the GUI thread."""
    done = pyqtSignal(object, object)   # result, error
    _alive: set = set()


def run_in_background(fn, on_done):
    br = _Bridge()
    br.done.connect(lambda res, err: (on_done(res, err), br.deleteLater()))
    _Bridge._alive.add(br)

    def work():
        try:
            res, err = fn(), None
        except Exception as e:  # noqa: BLE001
            res, err = None, e
        br.done.emit(res, err)
        _Bridge._alive.discard(br)

    threading.Thread(target=work, daemon=True).start()


def offer_app_update(window, res, quiet=False):
    """Asks to update when res holds a newer version. quiet: say nothing when there is none."""
    if not res:
        if not quiet:
            QMessageBox.information(window, tr("No update"),
                                    tr("{a} {v} is the newest version.").format(a=APP_NAME, v=APP_VERSION))
        return
    ver, text, notes = res
    esc = lambda x: x.replace("&", "&amp;").replace("<", "&lt;").replace("\n", "<br>")
    box = QMessageBox(window)
    box.setWindowTitle(tr("Update available"))
    box.setIcon(QMessageBox.Icon.Information)
    box.setTextFormat(Qt.TextFormat.RichText)
    box.setText(
        tr("<b>{a} {v}</b> is out (you have {c}).").format(a=APP_NAME, v=ver, c=APP_VERSION) + "<br><br>"
        + (esc(notes) + "<br><br>" if notes else "")
        + tr("Update now? It is saved next to this file as <b>{f}</b> and opens in place of this window. "
             "Your files, terms and settings stay. The old file stays too.").format(f=f"{APP_FILE_STEM}_v{ver}.pyw"))
    yes = box.addButton(tr("Update now"), QMessageBox.ButtonRole.AcceptRole)
    box.addButton(tr("Later"), QMessageBox.ButtonRole.RejectRole)
    box.exec()
    if box.clickedButton() is not yes:
        return
    try:
        dest = app_update_install(ver, text)
    except Exception as e:  # noqa: BLE001
        QMessageBox.warning(window, tr("Update failed"), tr("Could not save the new version:") + f"\n{e}")
        return
    if window is not None:
        window.close()           # saves settings first, so the new copy starts with them
    start_file(dest)
    QTimer.singleShot(300, QApplication.quit)


class AboutDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(tr("About {a}").format(a=APP_NAME))
        self.setFixedWidth(520)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 22, 24, 18)
        lay.setSpacing(12)
        head = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(make_app_icon().pixmap(52, 52))
        head.addWidget(logo)
        head.addWidget(QLabel(f"<span style='font-size:16pt;font-weight:700'>{APP_NAME}</span><br>"
                              + tr("Version {v}").format(v=APP_VERSION)), 1)
        lay.addLayout(head)
        info = QLabel(
            tr("Search many Excel, CSV, Word and PDF files for many terms at once.") + "<br><br>"
            "<b>Nikolaos K. Paridis</b> (@nikkpap) · ALU DEV TEAM<br>"
            f"<a href='{APP_PAGE}'>{APP_PAGE.replace('https://', '')}</a><br><br>"
            + tr("Built with Python, PyQt6, openpyxl, xlrd, RapidFuzz and PyMuPDF."))
        info.setTextFormat(Qt.TextFormat.RichText)
        info.setOpenExternalLinks(True)
        info.setWordWrap(True)
        lay.addWidget(info)

        st = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.auto = QCheckBox(tr("Check for a new version when the app starts"))
        self.auto.setChecked(str(st.value("app_update_check", "true")).lower() == "true")
        self.auto.toggled.connect(lambda v: st.setValue("app_update_check", "true" if v else "false"))
        lay.addWidget(self.auto)

        row = QHBoxLayout()
        self.check_btn = QPushButton(tr("Check for updates"), objectName="primary")
        self.check_btn.clicked.connect(self._check)
        sc = QPushButton(tr("Make shortcuts"))
        sc.setToolTip(tr("Desktop and Start menu shortcuts to this file"))
        sc.setVisible(os.name == "nt")
        sc.clicked.connect(lambda: QMessageBox.information(
            self, APP_NAME, tr("Shortcuts created.") if make_shortcuts(force=True)
            else tr("Could not create shortcuts.")))
        self.cache_btn = QPushButton()
        self._cache_label()
        self.cache_btn.setToolTip(tr("Files remembered for faster repeat searches. Safe to clear."))
        self.cache_btn.clicked.connect(self._clear_cache)
        close = QPushButton(tr("Close"))
        close.clicked.connect(self.accept)
        for b in (self.check_btn, sc, self.cache_btn):
            row.addWidget(b)
        row.addStretch()
        row.addWidget(close)
        lay.addLayout(row)

    def _cache_label(self):
        mb = cache_size(_cache_dir()) / (1024 * 1024)
        self.cache_btn.setText(tr("Clear memory ({mb:.0f} MB)").format(mb=mb))

    def _clear_cache(self):
        for f in _cache_dir().glob("*"):
            try:
                f.unlink()
            except OSError:
                pass
        self._cache_label()

    def _check(self):
        self.check_btn.setEnabled(False)
        self.check_btn.setText(tr("Checking…"))

        def done(res, err):
            self.check_btn.setEnabled(True)
            self.check_btn.setText(tr("Check for updates"))
            if err is not None:
                QMessageBox.warning(self, tr("Could not check"), tr("GitHub did not answer:") + f"\n{err}")
                return
            self.accept()
            offer_app_update(self.parent(), res)

        run_in_background(app_update_check, done)


# --------------------------------------------------------------------------- #

def _excepthook(etype, value, tb):
    """Under pythonw an uncaught error would just make the window vanish: log it and keep running."""
    import traceback
    text = "".join(traceback.format_exception(etype, value, tb))
    try:
        with open(_app_data_dir() / "error.log", "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + f"{APP_NAME} {APP_VERSION}\n{text}\n")
    except OSError:
        pass
    if sys.__stderr__:
        sys.__stderr__.write(text)


def main():
    global LANG
    sys.excepthook = _excepthook
    if sys.platform.startswith("win"):
        try:  # own taskbar icon instead of Python's
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("nikkpap.ExcelUltimateSearch")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyle("Fusion")
    f = QFont("Segoe UI" if sys.platform.startswith("win") else app.font().family(), 10)
    app.setFont(f)
    app.setWindowIcon(make_app_icon())
    st = QSettings(SETTINGS_ORG, SETTINGS_APP)
    lang = st.value("lang")
    if lang is None:   # first start: follow Windows' language
        from PyQt6.QtCore import QLocale
        lang = "el" if QLocale.system().name().startswith("el") else "en"
        st.setValue("lang", lang)
    LANG = "el" if lang == "el" else "en"

    if not startup(app):
        sys.exit(1)
    w = MainWindow()
    w.show()

    # silent update check, every start, in the background; asks only when a newer version exists
    if str(st.value("app_update_check", "true")).lower() == "true" and "--no-update" not in sys.argv:
        QTimer.singleShot(1500, lambda: run_in_background(
            app_update_check, lambda res, err: err is None and res and offer_app_update(w, res, quiet=True)))
    sys.exit(app.exec())


# --------------------------------------------------------------------------- #
#  Greek interface
# --------------------------------------------------------------------------- #

EL: dict = {
    "Starting": "Εκκίνηση",
    "Checking components": "Έλεγχος στοιχείων",
    "Preparing shortcuts": "Προετοιμασία συντομεύσεων",
    "Desktop and Start menu": "Επιφάνεια εργασίας και μενού Έναρξη",
    "Ready": "Έτοιμο",
    "Update available": "Υπάρχει ενημέρωση",
    "Update now": "Ενημέρωση τώρα",
    "Later": "Αργότερα",
    "Document": "Έγγραφο",
    "Done": "Τέλος",
    "Import terms from a column": "Εισαγωγή όρων από στήλη",
    "First row is a header (skip it)": "Η πρώτη γραμμή είναι επικεφαλίδα (παράλειψη)",
    "Replace terms": "Αντικατάσταση όρων",
    "Add to terms": "Προσθήκη στους όρους",
    "Cancel": "Άκυρο",
    "Change the language (the app restarts)": "Αλλαγή γλώσσας (η εφαρμογή ξεκινά ξανά)",
    "Light / dark theme": "Ανοιχτό / σκούρο θέμα",
    "About, and check for updates": "Σχετικά, και έλεγχος για ενημερώσεις",
    "＋ Files": "＋ Αρχεία",
    "Excel, CSV, Word (.docx) or PDF files — Ctrl+O": "Αρχεία Excel, CSV, Word (.docx) ή PDF — Ctrl+O",
    "＋ Folder": "＋ Φάκελος",
    "A folder with everything inside it; it is watched for new files":
        "Ένας φάκελος με ό,τι έχει μέσα· παρακολουθείται για νέα αρχεία",
    "Recent ▾": "Πρόσφατοι ▾",
    "Drop Excel / CSV / Word / PDF\nfiles or folders here": "Σύρετε εδώ αρχεία ή φακέλους\nExcel / CSV / Word / PDF",
    "Folders watched for new or deleted files": "Φάκελοι που παρακολουθούνται για νέα ή διαγραμμένα αρχεία",
    "Remove": "Αφαίρεση",
    "Remove the selected files from the list (Delete)": "Αφαίρεση των επιλεγμένων αρχείων από τη λίστα (Delete)",
    "Clear": "Καθαρισμός",
    "Lists ▾": "Λίστες ▾",
    "Save, load and import lists of terms": "Αποθήκευση, φόρτωση και εισαγωγή λιστών όρων",
    "One term per line…\n\nOr: Lists ▾ → Import from an Excel column":
        "Ένας όρος ανά γραμμή…\n\nΉ: Λίστες ▾ → Εισαγωγή από στήλη Excel",
    "How alike a text must be to count (Similar mode). Lower = more results.":
        "Πόσο όμοιο πρέπει να είναι ένα κείμενο για να μετρήσει (λειτουργία Παρόμοιο). Χαμηλότερο = περισσότερα.",
    "Match case": "Διάκριση πεζών-κεφαλαίων",
    "Ignore accents (ά = α, é = e)": "Αγνόηση τόνων (ά = α, é = e)",
    "Greek = Latin (Greeklish, wrong keyboard)": "Ελληνικά = Λατινικά (Greeklish, πληκτρολόγιο)",
    "skyrodema finds σκυρόδεμα, beton finds μπετόν, and a code typed with the Greek keyboard (Β500C with a "
    "Greek Β) finds B500C. Ignores case.":
        "Το skyrodema βρίσκει σκυρόδεμα, το beton βρίσκει μπετόν, και ένας κωδικός γραμμένος με ελληνικό "
        "πληκτρολόγιο (Β500C με ελληνικό Β) βρίσκει το B500C. Αγνοεί πεζά-κεφαλαία.",
    "Text cells only (skip numbers && dates)": "Μόνο κείμενο (όχι αριθμοί && ημερομηνίες)",
    "One hit per row per term": "Ένα εύρημα ανά γραμμή ανά όρο",
    "all  ·  e.g.  B, D:F, Περιγραφή": "όλες  ·  π.χ.  B, D:F, Περιγραφή",
    "Column letters (B), ranges (D:F), numbers, or header names from the first row. Put a short header name "
    "in quotes: \"Qty\". Empty = all columns.":
        "Γράμματα στηλών (B), περιοχές (D:F), αριθμοί, ή ονόματα επικεφαλίδων από την πρώτη γραμμή. Ένα σύντομο "
        "όνομα επικεφαλίδας σε εισαγωγικά: \"Qty\". Κενό = όλες οι στήλες.",
    "all  ·  e.g.  Επιμετρήσεις*, Υλικά": "όλα  ·  π.χ.  Επιμετρήσεις*, Υλικά",
    "Sheet names or parts of them, comma separated; * and ? as wildcards. Empty = all sheets. Word and PDF "
    "files are always searched whole.":
        "Ονόματα φύλλων ή μέρη τους, χωρισμένα με κόμμα· * και ? ως μπαλαντέρ. Κενό = όλα τα φύλλα. Τα Word "
        "και PDF ψάχνονται πάντα ολόκληρα.",
    "Use all processor cores (faster)": "Χρήση όλων των πυρήνων (γρηγορότερα)",
    "Remember unchanged files (faster repeat searches)": "Μνήμη αρχείων που δεν άλλαξαν (γρηγορότερα)",
    "4 · Search": "4 · Αναζήτηση",
    "Filter results…  (Ctrl+F)": "Φιλτράρισμα αποτελεσμάτων…  (Ctrl+F)",
    "Row preview — select a result": "Προεπισκόπηση γραμμής — επιλέξτε ένα αποτέλεσμα",
    "Open at this cell": "Άνοιγμα σε αυτό το κελί",
    "Row preview": "Προεπισκόπηση",
    "Copy": "Αντιγραφή",
    "Copy the selected rows (or all) — Ctrl+C": "Αντιγραφή των επιλεγμένων γραμμών (ή όλων) — Ctrl+C",
    "Export CSV": "Εξαγωγή CSV",
    "Excel report": "Αναφορά Excel",
    "Checklist found / not found, all results, and one sheet per term":
        "Λίστα ελέγχου βρέθηκε / δεν βρέθηκε, όλα τα αποτελέσματα, και ένα φύλλο ανά όρο",
    "Clear results": "Καθαρισμός αποτελεσμάτων",
    "Forget the list": "Διαγραφή της λίστας",
    "Refresh now (F5)": "Ανανέωση τώρα (F5)",
    "Add files": "Προσθήκη αρχείων",
    "Add a folder (with its subfolders)": "Προσθήκη φακέλου (με τους υποφακέλους του)",
    "💾  Save these terms as a list…": "💾  Αποθήκευση των όρων ως λίστα…",
    "📂  Load a list": "📂  Φόρτωση λίστας",
    "🗑  Delete a list": "🗑  Διαγραφή λίστας",
    "📊  Import from an Excel column…": "📊  Εισαγωγή από στήλη Excel…",
    "📄  Import from a text file…": "📄  Εισαγωγή από αρχείο κειμένου…",
    "Save list": "Αποθήκευση λίστας",
    "Name of the list:": "Όνομα της λίστας:",
    "Spreadsheet with the terms": "Υπολογιστικό φύλλο με τους όρους",
    "Text file with one term per line": "Αρχείο κειμένου με έναν όρο ανά γραμμή",
    "Cancel  (Esc)": "Ακύρωση  (Esc)",
    "Terms not found": "Όροι που δεν βρέθηκαν",
    "Copy list": "Αντιγραφή λίστας",
    "Search them as Similar": "Αναζήτησή τους ως Παρόμοιο",
    "Unreadable files": "Αρχεία που δεν διαβάστηκαν",
    "Open file": "Άνοιγμα αρχείου",
    "Show in folder": "Εμφάνιση στον φάκελο",
    "Copy found value": "Αντιγραφή της τιμής",
    "Copy whole row": "Αντιγραφή όλης της γραμμής",
    "Copy file path": "Αντιγραφή διαδρομής αρχείου",
    "Copy selected rows": "Αντιγραφή επιλεγμένων γραμμών",
    "Export results": "Εξαγωγή αποτελεσμάτων",
    "Term": "Όρος",
    "Found value": "Τιμή που βρέθηκε",
    "Match": "Ταίριασμα",
    "Cell": "Κελί",
    "Row": "Γραμμή",
    "Column": "Στήλη",
    "Sheet / page": "Φύλλο / σελίδα",
    "File": "Αρχείο",
    "Path": "Διαδρομή",
    "Checklist": "Λίστα ελέγχου",
    "Installing components (one time)": "Εγκατάσταση στοιχείων (μία φορά)",
    "Components installed": "Τα στοιχεία εγκαταστάθηκαν",
    "Check for a new version when the app starts": "Έλεγχος για νέα έκδοση κατά την εκκίνηση",
    "Check for updates": "Έλεγχος για ενημερώσεις",
    "Make shortcuts": "Δημιουργία συντομεύσεων",
    "Desktop and Start menu shortcuts to this file": "Συντομεύσεις στην Επιφάνεια εργασίας και στο μενού Έναρξη",
    "Files remembered for faster repeat searches. Safe to clear.":
        "Αρχεία που θυμάται η εφαρμογή για γρηγορότερη επανάληψη. Ασφαλές να καθαριστεί.",
    "Close": "Κλείσιμο",
    "Checking…": "Έλεγχος…",
    "Opening in Excel at {where}…": "Άνοιγμα στο Excel στο {where}…",
    "Excel could not be driven; opening the file normally": "Το Excel δεν ανταποκρίθηκε· άνοιγμα του αρχείου κανονικά",
    "Sheet": "Φύλλο",
    "Search many spreadsheets, Word and PDF files for many terms — at once.":
        "Αναζήτηση πολλών όρων σε πολλά Excel, Word και PDF — με μία κίνηση.",
    "1 · Files": "1 · Αρχεία",
    "2 · Search terms": "2 · Όροι αναζήτησης",
    "3 · Options": "3 · Επιλογές",
    "Only these columns": "Μόνο αυτές οι στήλες",
    "Only these sheets": "Μόνο αυτά τα φύλλα",
    "Results": "Αποτελέσματα",
    "Header": "Επικεφαλίδα",
    "Value": "Τιμή",
    "☀  Light": "☀  Ανοιχτό",
    "☾  Dark": "☾  Σκούρο",
    "No recent folders yet": "Δεν υπάρχουν ακόμα πρόσφατοι φάκελοι",
    "👁 No watched folders": "👁 Κανένας φάκελος σε παρακολούθηση",
    "Type or paste some terms first.": "Γράψτε ή επικολλήστε πρώτα κάποιους όρους.",
    "Spreadsheets": "Υπολογιστικά φύλλα",
    "Cancelling…": "Ακύρωση…",
    "Step 1: add at least one file or folder.": "Βήμα 1: προσθέστε τουλάχιστον ένα αρχείο ή φάκελο.",
    "Step 2: type at least one search term (one per line).":
        "Βήμα 2: γράψτε τουλάχιστον έναν όρο αναζήτησης (έναν ανά γραμμή).",
    "Click to see which terms had no match": "Πατήστε για να δείτε ποιοι όροι δεν βρέθηκαν",
    "No results match the filter.": "Κανένα αποτέλεσμα δεν ταιριάζει με το φίλτρο.",
    "There are no results to export.": "Δεν υπάρχουν αποτελέσματα για εξαγωγή.",
    "Excel workbook": "Βιβλίο Excel",
    "Export complete. Open the file now?": "Η εξαγωγή ολοκληρώθηκε. Άνοιγμα του αρχείου τώρα;",
    "Found": "Βρέθηκε",
    "Matches": "Ευρήματα",
    "Files": "Αρχεία",
    "First location": "Πρώτη θέση",
    "All results": "Όλα τα αποτελέσματα",
    "No update": "Καμία ενημέρωση",
    "Update failed": "Η ενημέρωση απέτυχε",
    "Built with Python, PyQt6, openpyxl, xlrd, RapidFuzz and PyMuPDF.":
        "Φτιαγμένο με Python, PyQt6, openpyxl, xlrd, RapidFuzz και PyMuPDF.",
    "Page": "Σελ.",
    "Exact match": "Ακριβές ταίριασμα",
    "Could not open:": "Δεν άνοιξε:",
    "No new files found": "Δεν βρέθηκαν νέα αρχεία",
    "All files": "Όλα τα αρχεία",
    "{n} files": "{n} αρχεία",
    "{n} terms": "{n} όροι",
    "No saved lists yet": "Δεν υπάρχουν ακόμα αποθηκευμένες λίστες",
    "Saved list “{n}” ({k} terms)": "Αποθηκεύτηκε η λίστα «{n}» ({k} όροι)",
    "Loaded list “{n}” ({k} terms)": "Φορτώθηκε η λίστα «{n}» ({k} όροι)",
    "{k} terms in the file. Replace your terms or add to them?":
        "{k} όροι στο αρχείο. Αντικατάσταση των όρων σας ή προσθήκη σε αυτούς;",
    "Nothing found.": "Δεν βρέθηκε τίποτα.",
    "{m} of {n} terms had no match:": "{m} από {n} όρους δεν βρέθηκαν:",
    "{n} file(s) could not be read (damaged, password-protected, scanned PDF, or open in another program).":
        "{n} αρχείο(α) δεν διαβάστηκαν (κατεστραμμένα, με κωδικό, σκαναρισμένο PDF, ή ανοιχτά σε άλλο πρόγραμμα).",
    "Filter by term “{t}”": "Φίλτρο με όρο «{t}»",
    "Filter by file “{f}”": "Φίλτρο με αρχείο «{f}»",
    "Copied {n} row(s)": "Αντιγράφηκαν {n} γραμμές",
    "search_report": "αναφορά_αναζήτησης",
    "Could not write the file — is it open in Excel?": "Δεν γράφτηκε το αρχείο — μήπως είναι ανοιχτό στο Excel;",
    "Exported {n} rows → {f}": "Εξήχθησαν {n} γραμμές → {f}",
    "Whole row": "Όλη η γραμμή",
    "Version {v}  ·  Excel, CSV, Word and PDF search": "Έκδοση {v}  ·  αναζήτηση σε Excel, CSV, Word και PDF",
    "Update now? It is saved next to this file as <b>{f}</b> and opens in place of this window. Your files, "
    "terms and settings stay. The old file stays too.":
        "Ενημέρωση τώρα; Αποθηκεύεται δίπλα σε αυτό το αρχείο ως <b>{f}</b> και ανοίγει στη θέση αυτού του "
        "παραθύρου. Τα αρχεία, οι όροι και οι ρυθμίσεις σας μένουν. Και το παλιό αρχείο μένει.",
    "Could not save the new version:": "Δεν αποθηκεύτηκε η νέα έκδοση:",
    "About {a}": "Σχετικά με το {a}",
    "Search many Excel, CSV, Word and PDF files for many terms at once.":
        "Αναζήτηση πολλών όρων σε πολλά αρχεία Excel, CSV, Word και PDF με μία κίνηση.",
    "Clear memory ({mb:.0f} MB)": "Καθαρισμός μνήμης ({mb:.0f} MB)",
    "Could not check": "Ο έλεγχος απέτυχε",
    "Reading {name}": "Ανάγνωση {name}",
    "Opened at {where}": "Άνοιξε στο {where}",
    "{n} distinct values": "{n} διαφορετικές τιμές",
    "👁 Watching {n} folder(s) ▾": "👁 Φάκελοι σε παρακολούθηση: {n} ▾",
    "Stop watching: {f}": "Τέλος παρακολούθησης: {f}",
    "Folders checked: {a} new, {r} removed": "Έλεγχος φακέλων: {a} νέα, {r} αφαιρέθηκαν",
    "Supported files": "Υποστηριζόμενα αρχεία",
    "Delete the list “{n}”?": "Διαγραφή της λίστας «{n}»;",
    "Could not read the file:": "Δεν διαβάστηκε το αρχείο:",
    "Imported {k} terms": "Εισήχθησαν {k} όροι",
    "Text files": "Αρχεία κειμένου",
    "{n} matches": "{n} ευρήματα",
    "{p} processes": "{p} διεργασίες",
    "{c} from memory": "{c} από τη μνήμη",
    "{s} of {n} matches": "{s} από {n} ευρήματα",
    "⚠ {n} not found": "⚠ {n} δεν βρέθηκαν",
    "✖ {n} unreadable": "✖ {n} δεν διαβάστηκαν",
    "No results yet.\n\n1 · add files   2 · type terms   3 · options   4 · Search":
        "Δεν υπάρχουν αποτελέσματα ακόμα.\n\n1 · αρχεία   2 · όροι   3 · επιλογές   4 · Αναζήτηση",
    "Row {r}  ·  {s}  ·  {f}": "Γραμμή {r}  ·  {s}  ·  {f}",
    "Open at {c}": "Άνοιγμα στο {c}",
    "Paragraph {r}": "Παράγραφος {r}",
    "Page {p}, line {r}": "Σελίδα {p}, γραμμή {r}",
    "Open in Excel at {c}": "Άνοιγμα στο Excel στο {c}",
    "Export failed:": "Η εξαγωγή απέτυχε:",
    "YES": "ΝΑΙ",
    "NO": "ΟΧΙ",
    "Found {f} of {n} terms": "Βρέθηκαν {f} από {n} όρους",
    "{a} {v} is the newest version.": "Το {a} {v} είναι η νεότερη έκδοση.",
    "Shortcuts created.": "Οι συντομεύσεις δημιουργήθηκαν.",
    "Could not create shortcuts.": "Δεν δημιουργήθηκαν οι συντομεύσεις.",
    "GitHub did not answer:": "Το GitHub δεν απάντησε:",
    "Searching with {n} processes": "Αναζήτηση με {n} διεργασίες",
    "Similar: {p}% alike": "Παρόμοιο: {p}% ίδιο",
    "Added {n} file(s)": "Προστέθηκαν {n} αρχεία",
    "A list named “{n}” exists. Replace it?": "Υπάρχει λίστα «{n}». Αντικατάσταση;",
    "Cancelled": "Ακυρώθηκε",
    "Finished": "Ολοκληρώθηκε",
    "<b>{a} {v}</b> is out (you have {c}).": "Βγήκε το <b>{a} {v}</b> (έχετε το {c}).",
    "Version {v}": "Έκδοση {v}",
    "Invalid regular expression:": "Μη έγκυρη κανονική έκφραση:",
    "Could not install: {p}\n\nCheck the internet connection and start the app again.":
        "Δεν εγκαταστάθηκαν: {p}\n\nΕλέγξτε τη σύνδεση στο διαδίκτυο και ξεκινήστε ξανά την εφαρμογή.",
    "Contains": "Περιέχει",
    "Whole cell": "Ολόκληρο κελί",
    "Starts with": "Ξεκινά με",
    "Whole word": "Ολόκληρη λέξη",
    "Similar (typos)": "Παρόμοιο (ορθογραφικά λάθη)",
    "Regular expression": "Κανονική έκφραση",
}


if __name__ == "__main__" and not os.environ.get("EUS_NO_MAIN"):
    main()
