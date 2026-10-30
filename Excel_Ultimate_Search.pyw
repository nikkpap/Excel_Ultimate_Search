# -*- coding: utf-8 -*-
"""
Excel Ultimate Search - search many Excel workbooks (.xlsx / .xlsm / .xls) and CSV files
for a list of terms at once. Python / PyQt6 rewrite of the original JavaFX tool.

Double-click this file on Windows: it opens without a console window, installs what it needs
on the first start (behind a splash screen) and looks for a newer version on GitHub at every start.
Each release is named Excel_Ultimate_Search_v<version>.pyw.

    python Excel_Ultimate_Search.pyw --console    keep the console, to read errors

github.com/nikkpap/Excel_Ultimate_Search  ·  Nikolaos K. Paridis (nikkpap) · ALU DEV TEAM
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "Excel Ultimate Search"
APP_VERSION = "2.1.0"
APP_FILE_STEM = "Excel_Ultimate_Search"
HERE = Path(__file__).resolve().parent
PACKAGES = ["PyQt6", "openpyxl", "xlrd"]                    # pip names
MODULES = {"PyQt6": "PyQt6.QtWidgets", "openpyxl": "openpyxl", "xlrd": "xlrd"}

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
if os.name == "nt" and "--console" not in sys.argv and Path(sys.executable).name.lower() == "python.exe" \
        and _gui_python() != sys.executable:
    import subprocess as _sp
    _sp.Popen([_gui_python(), str(Path(__file__).resolve())] + sys.argv[1:],
              creationflags=DETACHED | NO_WINDOW, close_fds=True)
    sys.exit(0)


def _missing_packages() -> list[str]:
    import importlib
    out = []
    for pkg in PACKAGES:
        try:
            importlib.import_module(MODULES[pkg])
        except ImportError:
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
    """First start only: PyQt6 is needed to draw anything, so it is installed behind a small Tk splash."""
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


_ensure_qt()
os.environ["QT_LOGGING_RULES"] = ";".join(filter(None, [os.environ.get("QT_LOGGING_RULES", ""),
                                                        "qt.qpa.windowclass.warning=false"]))

# ------------------------------------------------------------------------------------------------
import csv
import datetime as _dt
import json
import re
import subprocess
import threading
import time
from dataclasses import dataclass

from PyQt6.QtCore import (
    QAbstractTableModel, QModelIndex, QObject, QPointF, QRectF, QSettings,
    QSortFilterProxyModel, Qt, QThread, QTimer, QUrl, pyqtSignal,
)
from PyQt6.QtGui import (
    QAction, QBrush, QColor, QDesktopServices, QFont, QFontMetrics, QGuiApplication, QIcon,
    QKeySequence, QLinearGradient, QPainter, QPainterPath, QPen, QPixmap, QRadialGradient, QShortcut,
)
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog, QFileDialog,
    QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget,
    QListWidgetItem, QMainWindow, QMenu, QMessageBox, QPlainTextEdit,
    QProgressBar, QPushButton, QSizePolicy, QSplitter, QStyle,
    QStyledItemDelegate, QStyleOptionViewItem, QTableView, QVBoxLayout,
    QWidget,
)

EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".csv"}

# --------------------------------------------------------------------------- #
#  Theme
# --------------------------------------------------------------------------- #

THEMES = {
    "dark": dict(
        bg="#0f1115", surface="#171a21", surface2="#1e222b", border="#2a2f3a",
        text="#e6e8ee", muted="#8b93a7", accent="#3b82f6", accent_hover="#5b9bff",
        accent_text="#ffffff", hl_bg="#facc15", hl_text="#111111",
        danger="#f87171", ok="#34d399", sel="#26406b", row_alt="#1a1e26",
    ),
    "light": dict(
        bg="#f4f6fa", surface="#ffffff", surface2="#f0f2f7", border="#dde1ea",
        text="#1b1f2a", muted="#667085", accent="#2563eb", accent_hover="#1d4ed8",
        accent_text="#ffffff", hl_bg="#fde047", hl_text="#111111",
        danger="#dc2626", ok="#059669", sel="#cfe0ff", row_alt="#f8f9fc",
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
        color: {t['muted']}; font-size: 8.5pt; font-weight: 700; letter-spacing: 1px;
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

    QSplitter::handle {{ background: transparent; width: 8px; }}
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
    row: int          # 1-based, as Excel shows it
    col: int          # 1-based
    sheet: str
    path: str
    span: tuple[int, int] | None = None   # highlighted part of value

    @property
    def cell(self) -> str:
        return f"{col_letter(self.col)}{self.row}"

    @property
    def file_name(self) -> str:
        return os.path.basename(self.path)


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def to_text(v) -> str | None:
    """Convert any cell value to the text the user would see (roughly)."""
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


# ---- workbook readers: each yields (sheet_name, row_no, col_no, value) ----- #

def iter_xlsx(path: str, cancelled):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if cancelled():
                    return
                for c_idx, v in enumerate(row, start=1):
                    if v is not None:
                        yield ws.title, r_idx, c_idx, v
    finally:
        wb.close()


def iter_xls(path: str, cancelled):
    import xlrd
    book = xlrd.open_workbook(path, on_demand=True)
    try:
        for si in range(book.nsheets):
            sh = book.sheet_by_index(si)
            for r in range(sh.nrows):
                if cancelled():
                    return
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    t = cell.ctype
                    if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                        continue
                    v = cell.value
                    if t == xlrd.XL_CELL_DATE:
                        try:
                            v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                        except Exception:
                            pass
                    elif t == xlrd.XL_CELL_BOOLEAN:
                        v = bool(v)
                    yield sh.name, r + 1, c + 1, v
            book.unload_sheet(si)
    finally:
        book.release_resources()


def iter_csv(path: str, cancelled):
    raw = Path(path).read_bytes()
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
    sheet = Path(path).stem
    for r_idx, row in enumerate(csv.reader(text.splitlines(), dialect), start=1):
        if cancelled():
            return
        for c_idx, v in enumerate(row, start=1):
            if v != "":
                yield sheet, r_idx, c_idx, v


def iter_cells(path: str, cancelled):
    ext = Path(path).suffix.lower()
    if ext == ".xls":
        return iter_xls(path, cancelled)
    if ext == ".csv":
        return iter_csv(path, cancelled)
    return iter_xlsx(path, cancelled)


# --------------------------------------------------------------------------- #
#  Search worker
# --------------------------------------------------------------------------- #

MODES = ["Contains", "Whole cell", "Starts with", "Whole word", "Regular expression"]


class Matcher:
    def __init__(self, term: str, mode: str, match_case: bool, ignore_accents: bool):
        self.term = term
        self.mode = mode
        self.match_case = match_case
        self.ignore_accents = ignore_accents
        flags = 0 if match_case else re.IGNORECASE
        if mode == "Regular expression":
            pattern = term
        elif mode == "Whole cell":
            pattern = r"^\s*" + re.escape(self._norm(term)) + r"\s*$"
        elif mode == "Starts with":
            pattern = r"^\s*" + re.escape(self._norm(term))
        elif mode == "Whole word":
            pattern = r"(?<!\w)" + re.escape(self._norm(term)) + r"(?!\w)"
        else:
            pattern = re.escape(self._norm(term))
        self.rx = re.compile(pattern, flags | re.UNICODE)

    def _norm(self, s: str) -> str:
        return strip_accents(s) if self.ignore_accents else s

    def search(self, text: str):
        m = self.rx.search(self._norm(text))
        if not m:
            return None
        # strip_accents keeps string length for precomposed chars, so span maps 1:1
        return m.span()


_ACCENT_MAP = str.maketrans(
    "άέήίόύώϊϋΐΰΆΈΉΊΌΎΏΪΫ" "àáâãäåèéêëìíîïòóôõöùúûüýÿñçÀÁÂÃÄÅÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÝÑÇ",
    "αεηιουωιυιυΑΕΗΙΟΥΩΙΥ" "aaaaaaeeeeiiiiooooouuuuyyncAAAAAAEEEEIIIIOOOOOUUUUYNC",
)


def strip_accents(s: str) -> str:
    return s.translate(_ACCENT_MAP)


class SearchWorker(QObject):
    progress = pyqtSignal(int, int, str)       # done, total, message
    hits = pyqtSignal(list)                    # batch of Hit
    file_error = pyqtSignal(str, str)          # path, error
    finished = pyqtSignal(float, bool)         # seconds, cancelled

    def __init__(self, files, terms, mode, match_case, ignore_accents, text_only, one_per_row):
        super().__init__()
        self.files = files
        self.terms = terms
        self.mode = mode
        self.match_case = match_case
        self.ignore_accents = ignore_accents
        self.text_only = text_only
        self.one_per_row = one_per_row
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        t0 = time.perf_counter()
        matchers = [Matcher(t, self.mode, self.match_case, self.ignore_accents) for t in self.terms]
        total = len(self.files)
        cancelled = lambda: self._cancel
        for i, path in enumerate(self.files):
            if self._cancel:
                break
            self.progress.emit(i, total, f"Reading {os.path.basename(path)}")
            batch: list[Hit] = []
            seen_rows: set = set()
            try:
                for sheet, r, c, v in iter_cells(path, cancelled):
                    if self.text_only and not isinstance(v, str):
                        continue
                    text = to_text(v)
                    if not text:
                        continue
                    for m in matchers:
                        span = m.search(text)
                        if span is None:
                            continue
                        if self.one_per_row:
                            key = (m.term, sheet, r)
                            if key in seen_rows:
                                continue
                            seen_rows.add(key)
                        batch.append(Hit(m.term, text, r, c, sheet, path, span))
                    if len(batch) >= 500:
                        self.hits.emit(batch)
                        batch = []
            except Exception as e:  # noqa: BLE001
                self.file_error.emit(path, f"{type(e).__name__}: {e}")
            if batch:
                self.hits.emit(batch)
        self.progress.emit(total, total, "Done")
        self.finished.emit(time.perf_counter() - t0, self._cancel)


# --------------------------------------------------------------------------- #
#  Results model / view
# --------------------------------------------------------------------------- #

COLUMNS = ["#", "Term", "Found value", "Cell", "Row", "Sheet", "File"]
SORT_ROLE = Qt.ItemDataRole.UserRole + 1
HIT_ROLE = Qt.ItemDataRole.UserRole + 2


class ResultsModel(QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self.rows: list[Hit] = []
        self.muted = QColor("#8b93a7")

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(COLUMNS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return COLUMNS[section].upper()
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        h = self.rows[index.row()]
        col = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            return [
                index.row() + 1, h.term, h.value.replace("\n", " ⏎ "), h.cell, h.row,
                h.sheet, h.file_name,
            ][col]
        if role == SORT_ROLE:
            return [
                index.row(), h.term.casefold(), h.value.casefold(), (h.col, h.row),
                h.row, h.sheet.casefold(), h.file_name.casefold(),
            ][col]
        if role == HIT_ROLE:
            return h
        if role == Qt.ItemDataRole.ToolTipRole:
            if col == 6:
                return h.path
            if col == 2:
                return h.value
        if role == Qt.ItemDataRole.ForegroundRole and col in (0, 4):
            return QBrush(self.muted)
        if role == Qt.ItemDataRole.TextAlignmentRole and col in (0, 3, 4):
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
        hay = f"{h.term}\t{h.value}\t{h.sheet}\t{h.file_name}\t{h.cell}".casefold()
        return self.needle in hay


class HighlightDelegate(QStyledItemDelegate):
    """Paints the 'Found value' column with the matched part highlighted."""

    def __init__(self, parent, theme_getter):
        super().__init__(parent)
        self.theme = theme_getter

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index):
        h: Hit = index.data(HIT_ROLE)
        if h is None or h.span is None:
            return super().paint(painter, option, index)
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = opt.text
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)

        t = self.theme()
        s, e = h.span
        # value shown has \n replaced by " ⏎ " (3 chars) — recompute span offset
        pre_raw = h.value[:s]
        mid_raw = h.value[s:e]
        pre = pre_raw.replace("\n", " ⏎ ")
        mid = mid_raw.replace("\n", " ⏎ ")
        post = text[len(pre) + len(mid):]

        fm = QFontMetrics(opt.font)
        rect = opt.rect.adjusted(8, 0, -6, 0)
        avail = rect.width()
        # if the match is far right, trim the prefix so the hit stays visible
        if fm.horizontalAdvance(pre + mid) > avail and len(pre) > 12:
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
            painter.fillPath(path, QColor(t["hl_bg"]))
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
            for it in self.selectedItems():
                self.takeItem(self.row(it))
            self.window().update_counts()
            return
        super().keyPressEvent(e)


# --------------------------------------------------------------------------- #
#  Main window
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
        QMessageBox.warning(None, APP_NAME, f"Could not open:\n{path}\n\n{e}")


def reveal_in_folder(path: str):
    if sys.platform.startswith("win"):
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        open_path(os.path.dirname(path))


def card() -> QFrame:
    f = QFrame()
    f.setObjectName("card")
    return f


def section(text: str) -> QLabel:
    lb = QLabel(text.upper())
    lb.setObjectName("section")
    return lb


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("nikkpap", "ExcelUltimateSearch")
        self.theme_name = self.settings.value("theme", "dark")
        self.thread: QThread | None = None
        self.worker: SearchWorker | None = None
        self.errors: list[tuple[str, str]] = []
        self.last_terms: list[str] = []

        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setWindowIcon(make_app_icon())
        self.resize(1280, 780)
        self.setMinimumSize(980, 600)

        self._build_ui()
        self._restore()
        self.apply_theme()

    # ---------------- UI ---------------- #
    def _build_ui(self):
        root = QWidget(objectName="root")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 14, 18, 14)
        outer.setSpacing(12)

        # header
        hdr = QHBoxLayout()
        icon = QLabel()
        icon.setPixmap(make_app_icon().pixmap(36, 36))
        titles = QVBoxLayout()
        titles.setSpacing(0)
        t = QLabel(APP_NAME, objectName="title")
        st = QLabel("Search many spreadsheets for many terms — at once.", objectName="subtitle")
        titles.addWidget(t)
        titles.addWidget(st)
        hdr.addWidget(icon)
        hdr.addSpacing(6)
        hdr.addLayout(titles)
        hdr.addStretch()
        self.theme_btn = QPushButton(objectName="ghost")
        self.theme_btn.setToolTip("Toggle light / dark theme")
        self.theme_btn.clicked.connect(self.toggle_theme)
        hdr.addWidget(self.theme_btn)
        self.about_btn = QPushButton(f"ⓘ  v{APP_VERSION}", objectName="ghost")
        self.about_btn.setToolTip("About, and check for updates")
        self.about_btn.clicked.connect(lambda: AboutDialog(self).exec())
        hdr.addWidget(self.about_btn)
        outer.addLayout(hdr)

        split = QSplitter(Qt.Orientation.Horizontal)
        split.setChildrenCollapsible(False)
        outer.addWidget(split, 1)
        self.splitter = split

        # ---------- left panel ---------- #
        left = card()
        left.setMinimumWidth(320)
        lv = QVBoxLayout(left)
        lv.setContentsMargins(16, 14, 16, 16)
        lv.setSpacing(8)

        fh = QHBoxLayout()
        fh.addWidget(section("Files"))
        fh.addStretch()
        self.files_count = QLabel("0", objectName="subtitle")
        fh.addWidget(self.files_count)
        lv.addLayout(fh)

        self.file_list = FileList()
        self.file_list.files_dropped.connect(self.add_paths)
        self.file_list.setMinimumHeight(120)
        self.drop_hint = QLabel("Drop Excel / CSV files or folders here", objectName="dropHint")
        self.drop_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_hint.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        stack = QWidget()
        sl = QVBoxLayout(stack)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.addWidget(self.file_list)
        self.drop_hint.setParent(self.file_list.viewport())
        lv.addWidget(stack, 2)

        fb = QHBoxLayout()
        fb.setSpacing(6)
        b_add = QPushButton("＋ Files")
        b_add.clicked.connect(self.pick_files)
        b_dir = QPushButton("＋ Folder")
        b_dir.clicked.connect(self.pick_folder)
        b_rm = QPushButton("Remove")
        b_rm.clicked.connect(self.remove_selected_files)
        b_clr = QPushButton("Clear", objectName="danger")
        b_clr.clicked.connect(self.clear_files)
        for b in (b_add, b_dir, b_rm, b_clr):
            fb.addWidget(b)
        lv.addLayout(fb)

        lv.addSpacing(6)
        th = QHBoxLayout()
        th.addWidget(section("Search terms"))
        th.addStretch()
        self.terms_count = QLabel("0", objectName="subtitle")
        th.addWidget(self.terms_count)
        lv.addLayout(th)
        self.terms_edit = QPlainTextEdit()
        self.terms_edit.setPlaceholderText("Paste the strings to search for, one per line…")
        self.terms_edit.textChanged.connect(self.update_counts)
        lv.addWidget(self.terms_edit, 2)

        lv.addSpacing(6)
        lv.addWidget(section("Options"))
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Match"))
        self.mode_cb = QComboBox()
        self.mode_cb.addItems(MODES)
        mode_row.addWidget(self.mode_cb, 1)
        lv.addLayout(mode_row)

        self.cb_case = QCheckBox("Match case")
        self.cb_accents = QCheckBox("Ignore accents (ά = α, é = e)")
        self.cb_text_only = QCheckBox("Text cells only (skip numbers && dates)")
        self.cb_one_per_row = QCheckBox("One hit per row per term")
        self.cb_accents.setChecked(True)
        for cb in (self.cb_case, self.cb_accents, self.cb_text_only, self.cb_one_per_row):
            lv.addWidget(cb)

        lv.addSpacing(8)
        self.search_btn = QPushButton("Search", objectName="primary")
        self.search_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.search_btn.setToolTip("Ctrl+Enter")
        self.search_btn.clicked.connect(self.toggle_search)
        lv.addWidget(self.search_btn)

        split.addWidget(left)

        # ---------- right panel ---------- #
        right = card()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(16, 14, 16, 12)
        rv.setSpacing(10)

        rh = QHBoxLayout()
        rh.addWidget(section("Results"))
        rh.addSpacing(8)
        self.stat_hits = QLabel("0 matches", objectName="stat")
        self.stat_files = QLabel("0 files", objectName="stat")
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
        self.filter_edit.setPlaceholderText("Filter results…  (Ctrl+F)")
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.setFixedWidth(260)
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
        widths = [56, 140, 0, 70, 64, 130, 220]
        for i, w in enumerate(widths):
            if w:
                tv.setColumnWidth(i, w)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        tv.setItemDelegateForColumn(2, HighlightDelegate(tv, lambda: THEMES[self.theme_name]))
        tv.doubleClicked.connect(self._open_hit)
        tv.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        tv.customContextMenuRequested.connect(self._table_menu)
        self.table = tv

        self.empty_lbl = QLabel(
            "No results yet.\n\nAdd files on the left, paste your terms and press Search.",
            objectName="dropHint",
        )
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setParent(tv.viewport())
        self.empty_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        rv.addWidget(tv, 1)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setVisible(False)
        rv.addWidget(self.progress)

        bottom = QHBoxLayout()
        self.status_lbl = QLabel("Ready", objectName="subtitle")
        bottom.addWidget(self.status_lbl, 1)
        b_copy = QPushButton("Copy")
        b_copy.setToolTip("Copy selected rows (or all) to clipboard — Ctrl+C")
        b_copy.clicked.connect(self.copy_rows)
        b_csv = QPushButton("Export CSV")
        b_csv.clicked.connect(lambda: self.export("csv"))
        b_xlsx = QPushButton("Export Excel")
        b_xlsx.clicked.connect(lambda: self.export("xlsx"))
        b_clear = QPushButton("Clear results", objectName="danger")
        b_clear.clicked.connect(self.clear_results)
        for b in (b_copy, b_csv, b_xlsx, b_clear):
            bottom.addWidget(b)
        rv.addLayout(bottom)

        split.addWidget(right)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([360, 900])

        # shortcuts
        QShortcut(QKeySequence("Ctrl+Return"), self, activated=self.toggle_search)
        QShortcut(QKeySequence("Ctrl+Enter"), self, activated=self.toggle_search)
        QShortcut(QKeySequence("Ctrl+F"), self, activated=lambda: (self.filter_edit.setFocus(), self.filter_edit.selectAll()))
        QShortcut(QKeySequence("Ctrl+O"), self, activated=self.pick_files)
        QShortcut(QKeySequence("Escape"), self, activated=self._escape)
        copy_act = QAction(self)
        copy_act.setShortcut(QKeySequence.StandardKey.Copy)
        copy_act.setShortcutContext(Qt.ShortcutContext.WidgetShortcut)
        copy_act.triggered.connect(self.copy_rows)
        tv.addAction(copy_act)

        self.update_counts()
        self._update_empty()

    # ---------------- theme / persistence ---------------- #
    def apply_theme(self):
        t = THEMES[self.theme_name]
        QApplication.instance().setStyleSheet(build_qss(t))
        self.model.muted = QColor(t["muted"])
        self.theme_btn.setText("☀  Light" if self.theme_name == "dark" else "☾  Dark")
        self.table.viewport().update()

    def toggle_theme(self):
        self.theme_name = "light" if self.theme_name == "dark" else "dark"
        self.apply_theme()

    def _restore(self):
        s = self.settings
        geo = s.value("geometry")
        if geo is not None:
            self.restoreGeometry(geo)
        spl = s.value("splitter")
        if spl is not None:
            self.splitter.restoreState(spl)
        self.terms_edit.setPlainText(s.value("terms", "") or "")
        files = s.value("files", []) or []
        if isinstance(files, str):
            files = [files]
        self.add_paths([f for f in files if os.path.isfile(f)], quiet=True)
        self.mode_cb.setCurrentText(s.value("mode", MODES[0]))
        b = lambda k, d: str(s.value(k, d)).lower() in ("true", "1")
        self.cb_case.setChecked(b("case", False))
        self.cb_accents.setChecked(b("accents", True))
        self.cb_text_only.setChecked(b("text_only", False))
        self.cb_one_per_row.setChecked(b("one_per_row", False))
        self.last_dir = s.value("last_dir", str(Path.home()))

    def closeEvent(self, e):
        if self.worker:
            self.worker.cancel()
        if self.thread:
            self.thread.quit()
            self.thread.wait(3000)
        s = self.settings
        s.setValue("geometry", self.saveGeometry())
        s.setValue("splitter", self.splitter.saveState())
        s.setValue("theme", self.theme_name)
        s.setValue("terms", self.terms_edit.toPlainText())
        s.setValue("files", self.file_paths())
        s.setValue("mode", self.mode_cb.currentText())
        s.setValue("case", self.cb_case.isChecked())
        s.setValue("accents", self.cb_accents.isChecked())
        s.setValue("text_only", self.cb_text_only.isChecked())
        s.setValue("one_per_row", self.cb_one_per_row.isChecked())
        s.setValue("last_dir", self.last_dir)
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

    # ---------------- files ---------------- #
    def file_paths(self) -> list[str]:
        return [self.file_list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.file_list.count())]

    def add_paths(self, paths: list[str], quiet=False):
        existing = {os.path.normcase(p) for p in self.file_paths()}
        added = 0
        for p in paths:
            cands = []
            if os.path.isdir(p):
                for dp, _, fns in os.walk(p):
                    cands += [os.path.join(dp, f) for f in sorted(fns)]
            else:
                cands.append(p)
            for f in cands:
                name = os.path.basename(f)
                if Path(f).suffix.lower() not in EXCEL_EXTS or name.startswith("~$"):
                    continue
                key = os.path.normcase(os.path.abspath(f))
                if key in existing:
                    continue
                existing.add(key)
                it = QListWidgetItem(f"📄  {name}")
                it.setData(Qt.ItemDataRole.UserRole, os.path.abspath(f))
                it.setToolTip(os.path.abspath(f))
                self.file_list.addItem(it)
                added += 1
        if not quiet:
            self.status_lbl.setText(f"Added {added} file(s)" if added else "No new spreadsheet files found")
        self.update_counts()

    def pick_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add spreadsheets", self.last_dir,
            "Spreadsheets (*.xlsx *.xlsm *.xltx *.xltm *.xls *.csv);;All files (*)",
        )
        if files:
            self.last_dir = os.path.dirname(files[0])
            self.add_paths(files)

    def pick_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Add folder (searched recursively)", self.last_dir)
        if d:
            self.last_dir = d
            self.add_paths([d])

    def remove_selected_files(self):
        for it in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(it))
        self.update_counts()

    def clear_files(self):
        self.file_list.clear()
        self.update_counts()

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
        self.files_count.setText(f"{n} file{'s' if n != 1 else ''}")
        self.drop_hint.setVisible(n == 0)
        k = len(self.terms())
        self.terms_count.setText(f"{k} term{'s' if k != 1 else ''}")

    # ---------------- search ---------------- #
    def toggle_search(self):
        if self.worker:
            self.worker.cancel()
            self.status_lbl.setText("Cancelling…")
            return
        files = self.file_paths()
        terms = self.terms()
        if not files:
            QMessageBox.information(self, APP_NAME, "Add at least one Excel or CSV file first.")
            return
        if not terms:
            QMessageBox.information(self, APP_NAME, "Enter at least one search term (one per line).")
            self.terms_edit.setFocus()
            return
        mode = self.mode_cb.currentText()
        if mode == "Regular expression":
            for t in terms:
                try:
                    re.compile(t)
                except re.error as e:
                    QMessageBox.warning(self, APP_NAME, f"Invalid regular expression:\n\n{t}\n\n{e}")
                    return

        self.clear_results()
        self.last_terms = terms
        self.errors = []
        self.worker = SearchWorker(
            files, terms, mode, self.cb_case.isChecked(), self.cb_accents.isChecked(),
            self.cb_text_only.isChecked(), self.cb_one_per_row.isChecked(),
        )
        self.thread = QThread(self)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.hits.connect(self._on_hits)
        self.worker.file_error.connect(lambda p, e: self.errors.append((p, e)))
        self.worker.finished.connect(self._on_finished)
        self.table.setSortingEnabled(False)
        self.search_btn.setText("Cancel  (Esc)")
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
        self.status_lbl.setText(f"[{done}/{total}]  {msg}  ·  {len(self.model.rows)} matches")

    def _on_hits(self, hits):
        self.model.add(hits)
        self._update_stats()

    def _on_finished(self, secs, cancelled):
        self.thread.quit()
        self.thread.wait()
        self.thread = None
        self.worker = None
        self.table.setSortingEnabled(True)
        self.search_btn.setText("Search")
        self.progress.setVisible(False)
        n = len(self.model.rows)
        self.status_lbl.setText(
            f"{'Cancelled' if cancelled else 'Finished'} in {secs:.2f}s — {n} match{'es' if n != 1 else ''}"
        )
        self._update_stats(final=True)

    def _update_stats(self, final=False):
        rows = self.model.rows
        n = len(rows)
        shown = self.proxy.rowCount()
        self.stat_hits.setText(f"{n} matches" if shown == n else f"{shown} of {n} matches")
        nf = len({h.path for h in rows})
        self.stat_files.setText(f"{nf} file{'s' if nf != 1 else ''}")
        if final and self.last_terms:
            found = {h.term for h in rows}
            missing = [t for t in self.last_terms if t not in found]
            self.stat_missing.setVisible(bool(missing))
            self.stat_missing.setText(f"⚠ {len(missing)} not found")
            self.stat_missing.setStyleSheet(f"color: {THEMES[self.theme_name]['danger']};")
            self.stat_missing.setToolTip("Click to see which terms had no match")
        self.stat_errors.setVisible(bool(self.errors))
        if self.errors:
            self.stat_errors.setText(f"✖ {len(self.errors)} unreadable")
            self.stat_errors.setStyleSheet(f"color: {THEMES[self.theme_name]['danger']};")
        self._update_empty()

    def _update_empty(self):
        self.empty_lbl.setVisible(self.proxy.rowCount() == 0)
        if self.model.rows and self.proxy.rowCount() == 0:
            self.empty_lbl.setText("No results match the filter.")
        elif not self.model.rows and self.last_terms and not self.worker:
            self.empty_lbl.setText("Nothing found.")
        elif not self.model.rows:
            self.empty_lbl.setText("No results yet.\n\nAdd files on the left, paste your terms and press Search.")

    def _on_filter(self, s):
        self.proxy.set_needle(s)
        self._update_stats()

    def show_missing(self):
        found = {h.term for h in self.model.rows}
        missing = [t for t in self.last_terms if t not in found]
        box = QMessageBox(self)
        box.setWindowTitle("Terms not found")
        box.setText(f"{len(missing)} of {len(self.last_terms)} terms had no match:")
        box.setDetailedText("\n".join(missing))
        copy = box.addButton("Copy list", QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Close)
        box.exec()
        if box.clickedButton() is copy:
            QGuiApplication.clipboard().setText("\n".join(missing))

    def show_errors(self):
        txt = "\n\n".join(f"{p}\n   {e}" for p, e in self.errors)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle("Unreadable files")
        box.setText(f"{len(self.errors)} file(s) could not be read (corrupt, password-protected or open/locked).")
        box.setDetailedText(txt)
        box.exec()

    def clear_results(self):
        self.model.clear()
        self.last_terms = []
        self.errors = []
        self.stat_missing.setVisible(False)
        self.stat_errors.setVisible(False)
        self._update_stats()

    # ---------------- results actions ---------------- #
    def _hit_at(self, proxy_index) -> Hit | None:
        if not proxy_index.isValid():
            return None
        return self.model.rows[self.proxy.mapToSource(proxy_index).row()]

    def _open_hit(self, idx):
        h = self._hit_at(idx)
        if h:
            open_path(h.path)

    def _selected_hits(self) -> list[Hit]:
        rows = sorted({i.row() for i in self.table.selectionModel().selectedRows()})
        return [self.model.rows[self.proxy.mapToSource(self.proxy.index(r, 0)).row()] for r in rows]

    def _visible_hits(self) -> list[Hit]:
        return [self.model.rows[self.proxy.mapToSource(self.proxy.index(r, 0)).row()]
                for r in range(self.proxy.rowCount())]

    def _table_menu(self, pos):
        idx = self.table.indexAt(pos)
        h = self._hit_at(idx)
        if not h:
            return
        m = QMenu(self)
        m.addAction("Open file", lambda: open_path(h.path))
        m.addAction("Show in folder", lambda: reveal_in_folder(h.path))
        m.addSeparator()
        m.addAction("Copy found value", lambda: QGuiApplication.clipboard().setText(h.value))
        m.addAction("Copy file path", lambda: QGuiApplication.clipboard().setText(h.path))
        m.addAction("Copy selected rows", self.copy_rows)
        m.addSeparator()
        m.addAction(f"Filter by term “{h.term}”", lambda: self.filter_edit.setText(h.term))
        m.addAction(f"Filter by file “{h.file_name}”", lambda: self.filter_edit.setText(h.file_name))
        m.exec(self.table.viewport().mapToGlobal(pos))

    def copy_rows(self):
        hits = self._selected_hits() or self._visible_hits()
        if not hits:
            return
        lines = ["\t".join(["Term", "Found value", "Cell", "Row", "Sheet", "File", "Path"])]
        for h in hits:
            lines.append("\t".join([h.term, h.value.replace("\t", " ").replace("\n", " "),
                                    h.cell, str(h.row), h.sheet, h.file_name, h.path]))
        QGuiApplication.clipboard().setText("\n".join(lines))
        self.status_lbl.setText(f"Copied {len(hits)} row(s) to clipboard")

    def export(self, kind: str):
        hits = self._visible_hits()
        if not hits:
            QMessageBox.information(self, APP_NAME, "There are no results to export.")
            return
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M")
        default = os.path.join(self.last_dir, f"search_results_{stamp}.{kind}")
        flt = "CSV (*.csv)" if kind == "csv" else "Excel workbook (*.xlsx)"
        path, _ = QFileDialog.getSaveFileName(self, "Export results", default, flt)
        if not path:
            return
        header = ["#", "Term", "Found value", "Cell", "Row", "Column", "Sheet", "File", "Path"]
        data = [[i, h.term, h.value, h.cell, h.row, col_letter(h.col), h.sheet, h.file_name, h.path]
                for i, h in enumerate(hits, 1)]
        try:
            if kind == "csv":
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f, delimiter=";")
                    w.writerow(header)
                    w.writerows(data)
            else:
                self._write_xlsx(path, header, data)
        except PermissionError:
            QMessageBox.warning(self, APP_NAME, "Could not write the file — is it open in Excel?")
            return
        except Exception as e:  # noqa: BLE001
            QMessageBox.warning(self, APP_NAME, f"Export failed:\n{e}")
            return
        self.last_dir = os.path.dirname(path)
        self.status_lbl.setText(f"Exported {len(hits)} rows → {os.path.basename(path)}")
        if QMessageBox.question(self, APP_NAME, "Export complete. Open the file now?") == QMessageBox.StandardButton.Yes:
            open_path(path)

    def _write_xlsx(self, path, header, data):
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(header)
        for row in data:
            ws.append(row)
        fill = PatternFill("solid", fgColor="1F4E79")
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = fill
            c.alignment = Alignment(vertical="center")
        widths = [6, 22, 50, 8, 8, 8, 20, 30, 60]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        # summary sheet
        s = wb.create_sheet("Summary")
        s.append(["Term", "Matches"])
        counts: dict[str, int] = {}
        for r in data:
            counts[r[1]] = counts.get(r[1], 0) + 1
        for t in self.last_terms or counts:
            s.append([t, counts.get(t, 0)])
        for c in s[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = fill
        s.column_dimensions["A"].width = 40
        s.column_dimensions["B"].width = 12
        wb.save(path)


# --------------------------------------------------------------------------- #

# --------------------------------------------------------------------------- #
#  Startup splash (every launch): packages check, short pause, then the app
# --------------------------------------------------------------------------- #

SPLASH_MIN_STEP = 0.45    # seconds each step stays on screen, so it can be read
SETTINGS_ORG, SETTINGS_APP = "nikkpap", "ExcelUltimateSearch"


class SplashScreen(QWidget):
    def __init__(self):
        super().__init__(None, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint |
                         Qt.WindowType.SplashScreen)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(520, 300)
        self.shown_p, self.target_p, self.phase = 0.0, 0.0, 0.0
        self.status, self.detail = "Starting", ""
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

    def paintEvent(self, _):
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
                   f"Version {APP_VERSION}  ·  search many spreadsheets at once")
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
    """Splash: check / install packages, shortcuts. False = cannot continue."""
    sp = SplashScreen()
    sp.show()
    sp.set_step(0.15, "Starting", f"Python {sys.version.split()[0]}")
    _wait(app, SPLASH_MIN_STEP)

    sp.set_step(0.35, "Checking components", "openpyxl, xlrd")
    _wait(app, SPLASH_MIN_STEP)
    missing = _missing_packages()
    if missing:
        sp.set_step(0.45, "Installing components", " ".join(missing))
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
                                 "Could not install: " + ", ".join(still) +
                                 "\n\nCheck the internet connection and start the app again.\n\n" +
                                 "".join(state["out"])[-1000:])
            return False
        sp.set_step(0.7, "Components installed", " ".join(missing))
        _wait(app, SPLASH_MIN_STEP)

    sp.set_step(0.8, "Preparing shortcuts", "Desktop and Start menu")
    try:
        make_shortcuts()
    except Exception:
        pass
    _wait(app, SPLASH_MIN_STEP)
    sp.set_step(1.0, "Ready", "")
    _wait(app, SPLASH_MIN_STEP + 0.3)
    sp.close()
    return True


def _app_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") if os.name == "nt" else None
    d = (Path(base) / "ExcelUltimateSearch") if base else (Path.home() / ".excel_ultimate_search")
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


def start_file(path: Path):
    flags = (DETACHED | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)) if os.name == "nt" else 0
    subprocess.Popen([_gui_python(), str(path)], cwd=str(path.parent), close_fds=True, creationflags=flags)


class _Bridge(QObject):
    """Carries a background thread's result back to the GUI thread."""
    done = pyqtSignal(object, object)   # result, error


def run_in_background(fn, on_done):
    br = _Bridge()
    br.done.connect(lambda res, err: (on_done(res, err), br.deleteLater()))
    _Bridge._alive = getattr(_Bridge, "_alive", set())
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
            QMessageBox.information(window, "No update", f"{APP_NAME} {APP_VERSION} is the newest version.")
        return
    ver, text, notes = res
    esc = lambda x: x.replace("&", "&amp;").replace("<", "&lt;").replace("\n", "<br>")
    box = QMessageBox(window)
    box.setWindowTitle("Update available")
    box.setIcon(QMessageBox.Icon.Information)
    box.setTextFormat(Qt.TextFormat.RichText)
    box.setText(
        f"<b>{APP_NAME} {ver}</b> is out (you have {APP_VERSION}).<br><br>"
        + (esc(notes) + "<br><br>" if notes else "")
        + f"Update now? It is saved next to this file as <b>{APP_FILE_STEM}_v{ver}.pyw</b> and opens in "
          "place of this window. Your files, terms and settings stay. The old file stays too.")
    yes = box.addButton("Update now", QMessageBox.ButtonRole.AcceptRole)
    box.addButton("Later", QMessageBox.ButtonRole.RejectRole)
    box.exec()
    if box.clickedButton() is not yes:
        return
    try:
        dest = app_update_install(ver, text)
    except Exception as e:  # noqa: BLE001
        QMessageBox.warning(window, "Update failed", f"Could not save the new version:\n{e}")
        return
    if window is not None:
        window.close()           # saves settings first, so the new copy starts with them
    start_file(dest)
    QTimer.singleShot(300, QApplication.quit)


class AboutDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(f"About {APP_NAME}")
        self.setFixedWidth(500)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 22, 24, 18)
        lay.setSpacing(12)
        head = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(make_app_icon().pixmap(52, 52))
        head.addWidget(logo)
        t = QLabel(f"<span style='font-size:16pt;font-weight:700'>{APP_NAME}</span><br>"
                   f"Version {APP_VERSION}")
        head.addWidget(t, 1)
        lay.addLayout(head)
        info = QLabel(
            "Search many Excel / CSV files for many terms at once.<br><br>"
            "<b>Nikolaos K. Paridis</b> (@nikkpap) · ALU DEV TEAM<br>"
            f"<a href='{APP_PAGE}'>{APP_PAGE.replace('https://', '')}</a><br><br>"
            "Built with Python, PyQt6, openpyxl and xlrd.")
        info.setTextFormat(Qt.TextFormat.RichText)
        info.setOpenExternalLinks(True)
        info.setWordWrap(True)
        lay.addWidget(info)

        st = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.auto = QCheckBox("Check for a new version when the app starts")
        self.auto.setChecked(str(st.value("app_update_check", "true")).lower() == "true")
        self.auto.toggled.connect(lambda v: st.setValue("app_update_check", "true" if v else "false"))
        lay.addWidget(self.auto)

        row = QHBoxLayout()
        self.check_btn = QPushButton("Check for updates", objectName="primary")
        self.check_btn.clicked.connect(self._check)
        sc = QPushButton("Make shortcuts")
        sc.setToolTip("Desktop and Start menu shortcuts to this file (Windows)")
        sc.setVisible(os.name == "nt")
        sc.clicked.connect(lambda: QMessageBox.information(
            self, APP_NAME, "Shortcuts created." if make_shortcuts(force=True) else "Could not create shortcuts."))
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        row.addWidget(self.check_btn)
        row.addWidget(sc)
        row.addStretch()
        row.addWidget(close)
        lay.addLayout(row)

    def _check(self):
        self.check_btn.setEnabled(False)
        self.check_btn.setText("Checking…")

        def done(res, err):
            self.check_btn.setEnabled(True)
            self.check_btn.setText("Check for updates")
            if err is not None:
                QMessageBox.warning(self, "Could not check", f"GitHub did not answer:\n{err}")
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

    if not startup(app):
        sys.exit(1)
    w = MainWindow()
    w.show()

    # silent update check, every start, in the background; asks only when a newer version exists
    if str(QSettings(SETTINGS_ORG, SETTINGS_APP).value("app_update_check", "true")).lower() == "true" \
            and "--no-update" not in sys.argv:
        QTimer.singleShot(1500, lambda: run_in_background(
            app_update_check, lambda res, err: err is None and res and offer_app_update(w, res, quiet=True)))
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
