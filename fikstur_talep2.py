#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Fikstür Talep Uygulaması
OKTM — Optik Kaplama Tasarım Müdürlüğü
Mac / Windows uyumlu PySide6 sürümü
"""

import math
import os
import re
import sys
import shutil
import time
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

IS_WINDOWS = sys.platform.startswith("win")
try:
    if IS_WINDOWS:
        import pythoncom
        import win32com.client
        try:
            import win32timezone  # PyInstaller paketinde COM zaman dilimi desteği için gerekli
        except Exception:
            pass
    else:
        pythoncom = None
        win32com = None
except Exception:
    pythoncom = None
    win32com = None

from PySide6.QtCore import Qt, QSize, QTimer, Signal, QPointF
from PySide6.QtGui import QFont, QAction, QPainter, QColor, QPen, QPolygonF
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QComboBox,
    QButtonGroup, QTextEdit, QPushButton, QFileDialog, QMessageBox,
    QHBoxLayout, QVBoxLayout, QGridLayout, QFrame, QScrollArea, QSizePolicy,
    QStyleOptionComboBox, QStyle, QStylePainter, QStyledItemDelegate
)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment as XLAlign, Font as XLFont, PatternFill, Border, Side
import openpyxl.utils as xlutils

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ========================= CONFIG =========================

APP_TITLE = "Fikstür Talep Uygulaması"
APP_VERSION = "v2.0"
MAIL_TO = "oguzhanb@aselsan.com.tr"
FOOTER_TEXT = "OKTM — Optik Kaplama Tasarım Müdürlüğü"
LOG_FILENAME = "fikstur_talep_log.xlsx"
MAIL_SUBJECT_PREFIX = "Fikstür Talep"

COATING_SYSTEMS = [f"F{i}" for i in range(1, 13)] + [f"S{i}" for i in range(1, 5)]
MATERIALS = ["Alüminyum", "Paslanmaz Çelik", "Titanyum", "PEEK", "Delrin", "Diğer"]
QTY_OPTIONS = [str(i) for i in range(1, 11)]

URGENCY_CONFIG = [
    ("Düşük",   "#dbeafe", "#93c5fd", "#1e40af"),
    ("Normal",  "#d1fae5", "#6ee7b7", "#065f46"),
    ("Yüksek",  "#fef3c7", "#fbbf24", "#92400e"),
    ("Kritik",  "#fee2e2", "#f87171", "#991b1b"),
]

DEFAULT_BASE_DIR = Path.home() / "Documents" / "FixtureRequests"
BASE_DIR = Path(os.environ.get("FIXTURE_BASE_DIR", str(DEFAULT_BASE_DIR))).expanduser()


# ========================= HELPERS =========================

def sanitize_name(text: str) -> str:
    text = (text or "").strip()
    text = re.sub(r"[^\w\-\.\(\)]+", "_", text, flags=re.UNICODE)
    text = re.sub(r"_+", "_", text)
    return text[:120] if text else "X"


def format_size(path: Path) -> str:
    try:
        sz = path.stat().st_size
        if sz < 1024:
            return f"{sz} B"
        if sz < 1024 ** 2:
            return f"{sz / 1024:.1f} KB"
        return f"{sz / 1024 ** 2:.1f} MB"
    except Exception:
        return ""


def next_request_no(base_dir: Path, project: str, lens: str) -> str:
    """
    Proje-Lens-N formatında talep numarası üretir.
    """
    base_dir.mkdir(parents=True, exist_ok=True)

    safe_project = sanitize_name(project or "Proje")
    safe_lens = sanitize_name(lens or "Lens")
    prefix = f"{safe_project}-{safe_lens}-"

    max_n = 0
    for p in base_dir.iterdir():
        if p.is_dir() and p.name.startswith(prefix):
            suffix = p.name[len(prefix):]
            if suffix.isdigit():
                max_n = max(max_n, int(suffix))

    return f"{prefix}{max_n + 1}"


def copy_into_folder(src: Path, dest_folder: Path, prefix: str) -> Path:
    dest_folder.mkdir(parents=True, exist_ok=True)
    src = src.expanduser().resolve()
    name = sanitize_name(f"{prefix}__{src.name}")
    dest = dest_folder / name
    shutil.copy2(src, dest)
    return dest


_URGENCY_ROW_COLORS = {
    "Düşük":  "DBEAFE",  # açık mavi
    "Normal": "D1FAE5",  # açık yeşil
    "Yüksek": "FEF3C7",  # açık sarı
    "Kritik": "FEE2E2",  # açık kırmızı
}


def _row_fill(urgency: str) -> PatternFill | None:
    hex_color = _URGENCY_ROW_COLORS.get(urgency)
    if hex_color:
        return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    return None


def append_to_master_log(base_dir: Path, data: dict):
    log_path = base_dir / LOG_FILENAME
    new_keys = list(data.keys())

    # Stiller
    header_font  = XLFont(bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    center_align = XLAlign(horizontal="center", vertical="center", wrap_text=True)
    thin         = Side(style="thin", color="AAAAAA")
    cell_border  = Border(left=thin, right=thin, top=thin, bottom=thin)

    urgency_val = str(data.get("Aciliyet", ""))
    row_fill    = _row_fill(urgency_val)

    col_widths = {
        "Talep No": 24, "Talep Tipi": 12, "Zaman Damgası": 22,
        "Proje": 30, "Lens Doküman No": 24, "Fikstür Adedi": 14,
        "Kaplama Sistemi": 18, "Fikstür Malzemesi": 20,
        "Mesh Uygulaması": 18, "Buss Bar Kalınlığı mm": 22,
        "Aciliyet": 12, "Not": 40,
        "STP Dosyası": 50, "PDF Dosyası": 50, "Talep Klasörü": 60,
    }

    if log_path.exists():
        wb = load_workbook(log_path)
        ws = wb.active

        col_map = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v:
                col_map[str(v)] = c

        for key in new_keys:
            if key not in col_map:
                nc = ws.max_column + 1
                cell = ws.cell(1, nc, value=key)
                cell.font      = header_font
                cell.fill      = header_fill
                cell.alignment = center_align
                cell.border    = cell_border
                col_map[key]   = nc

        nr = ws.max_row + 1
        for key, value in data.items():
            c = col_map.get(key)
            if c:
                cell           = ws.cell(nr, c, value="" if value is None else str(value))
                cell.alignment = center_align
                cell.border    = cell_border
                if row_fill:
                    cell.fill  = row_fill
        # Eksik sütunları da kenarlıkla doldur
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(nr, c)
            cell.border = cell_border
            if row_fill and not cell.value:
                cell.fill = row_fill

    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Talep Log"
        ws.row_dimensions[1].height = 28

        for c, key in enumerate(new_keys, start=1):
            cell           = ws.cell(1, c, value=key)
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = center_align
            cell.border    = cell_border

        nr = 2
        for c, value in enumerate(data.values(), start=1):
            cell           = ws.cell(nr, c, value="" if value is None else str(value))
            cell.alignment = center_align
            cell.border    = cell_border
            if row_fill:
                cell.fill  = row_fill

    # Satır yüksekliği
    ws.row_dimensions[nr].height = 22

    # Sütun genişlikleri
    for c in range(1, ws.max_column + 1):
        header = ws.cell(1, c).value
        letter = xlutils.get_column_letter(c)
        width  = col_widths.get(str(header), 18) if header else 18
        if ws.column_dimensions[letter].width < width:
            ws.column_dimensions[letter].width = width

    wb.save(log_path)


_PDF_FONT_REGISTERED = False
_PDF_FONT_NAME = "Helvetica"
_PDF_FONT_BOLD = "Helvetica-Bold"


def _ensure_pdf_font():
    global _PDF_FONT_REGISTERED, _PDF_FONT_NAME, _PDF_FONT_BOLD

    if _PDF_FONT_REGISTERED:
        return

    win_fonts = Path(os.environ.get("SystemRoot", "C:/Windows")) / "Fonts"
    candidates = [
        # Windows — Türkçe karakter desteği olan fontlar
        win_fonts / "arial.ttf",
        win_fonts / "tahoma.ttf",
        win_fonts / "calibri.ttf",
        win_fonts / "segoeui.ttf",
        win_fonts / "verdana.ttf",
        # macOS
        Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/Library/Fonts/Arial.ttf"),
        Path("/System/Library/Fonts/Supplemental/Helvetica.ttc"),
        # Linux
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
    ]

    for p in candidates:
        if p.exists():
            try:
                pdfmetrics.registerFont(TTFont("PDFBody", str(p)))
                _PDF_FONT_NAME = "PDFBody"
                _PDF_FONT_BOLD = "PDFBody"
                _PDF_FONT_REGISTERED = True
                return
            except Exception:
                pass

    _PDF_FONT_REGISTERED = True


_URGENCY_PDF_ACCENT = {
    "Düşük":  (0.37, 0.58, 0.93),
    "Normal": (0.02, 0.55, 0.38),
    "Yüksek": (0.85, 0.45, 0.00),
    "Kritik": (0.75, 0.10, 0.10),
}


def write_pdf(path: Path, title: str, data: dict):
    _ensure_pdf_font()

    c = rl_canvas.Canvas(str(path), pagesize=A4)
    W, H = A4

    urgency      = str(data.get("Aciliyet", "Normal"))
    accent_color = _URGENCY_PDF_ACCENT.get(urgency, (0.15, 0.23, 0.37))

    # Pastel renk paleti — ferah beyaz/açık gri zemin
    page_bg      = (0.97, 0.98, 1.00)   # çok açık mavi-beyaz
    header_bg    = (0.12, 0.23, 0.37)   # lacivert
    label_bg     = (0.94, 0.96, 0.99)   # çok açık lavanta
    row_even_bg  = (1.00, 1.00, 1.00)   # beyaz
    row_odd_bg   = (0.96, 0.97, 1.00)   # çok hafif mavi
    text_dark    = (0.08, 0.14, 0.24)   # koyu lacivert metin
    text_mid     = (0.30, 0.38, 0.50)   # orta ton etiket
    border_color = (0.80, 0.86, 0.94)   # açık mavi-gri çizgi

    # ── Sayfa arka planı ───────────────────────────────────────
    c.setFillColorRGB(*page_bg)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # ── Üst başlık bandı ───────────────────────────────────────
    c.setFillColorRGB(*header_bg)
    c.rect(0, H - 88, W, 88, fill=1, stroke=0)

    # Accent şeridi (ince alt kenar)
    c.setFillColorRGB(*accent_color)
    c.rect(0, H - 91, W, 3, fill=1, stroke=0)

    # Başlık
    c.setFillColorRGB(1, 1, 1)
    c.setFont(_PDF_FONT_BOLD, 20)
    c.drawCentredString(W / 2, H - 42, title)

    # Alt bilgi satırı
    talep_no = str(data.get("Talep No", ""))
    tarih    = str(data.get("Zaman Damgası", ""))
    c.setFont(_PDF_FONT_NAME, 9)
    c.setFillColorRGB(0.75, 0.85, 1.00)
    c.drawCentredString(W / 2, H - 62, f"{talep_no}   ·   {tarih}")

    # Kurum adı sağ üst
    c.setFont(_PDF_FONT_NAME, 8)
    c.setFillColorRGB(0.65, 0.78, 0.95)
    c.drawRightString(W - 36, H - 78, "OKTM — Optik Kaplama Tasarım Müdürlüğü")

    # Aciliyet pill — sol üst
    pill_w, pill_h = 80, 22
    c.setFillColorRGB(*accent_color)
    c.roundRect(36, H - 78, pill_w, pill_h, 5, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont(_PDF_FONT_BOLD, 9)
    c.drawCentredString(36 + pill_w / 2, H - 78 + 7, urgency)

    # ── Tablo ──────────────────────────────────────────────────
    skip_keys = {"Talep No", "Zaman Damgası", "Talep Klasörü"}
    rows = [(k, "" if v is None else str(v))
            for k, v in data.items() if k not in skip_keys]

    mx      = 36
    table_w = W - 2 * mx
    col1_w  = 175
    col2_w  = table_w - col1_w
    row_h   = 28
    y_start = H - 108

    # Hafif gölge
    c.setFillColorRGB(0.82, 0.87, 0.94)
    c.roundRect(mx + 2, y_start - len(rows) * row_h - 2, table_w, len(rows) * row_h, 6, fill=1, stroke=0)

    for i, (key, val) in enumerate(rows):
        y_bot = y_start - (i + 1) * row_h
        text_cy = y_bot + row_h / 2 - 5   # dikey orta

        # Satır arka planı
        c.setFillColorRGB(*(row_even_bg if i % 2 == 0 else row_odd_bg))
        c.rect(mx, y_bot, table_w, row_h, fill=1, stroke=0)

        # Etiket sütunu arka planı
        c.setFillColorRGB(*label_bg)
        c.rect(mx, y_bot, col1_w, row_h, fill=1, stroke=0)

        # Etiket metni — ortalı
        c.setFillColorRGB(*text_mid)
        c.setFont(_PDF_FONT_BOLD, 9)
        c.drawCentredString(mx + col1_w / 2, text_cy, key)

        # Değer metni — ortalı, uzunsa kırp
        c.setFillColorRGB(*text_dark)
        c.setFont(_PDF_FONT_NAME, 10)
        max_chars = int(col2_w / 6.0)
        display_val = val if len(val) <= max_chars else val[:max_chars - 3] + "..."
        c.drawCentredString(mx + col1_w + col2_w / 2, text_cy, display_val)

        # Yatay ayraç
        c.setStrokeColorRGB(*border_color)
        c.setLineWidth(0.4)
        c.line(mx, y_bot, mx + table_w, y_bot)

    # Tablo dış çerçevesi
    total_h = len(rows) * row_h
    c.setStrokeColorRGB(*accent_color)
    c.setLineWidth(1.2)
    c.roundRect(mx, y_start - total_h, table_w, total_h, 6, fill=0, stroke=1)

    # Dikey sütun ayracı
    c.setStrokeColorRGB(*border_color)
    c.setLineWidth(0.8)
    c.line(mx + col1_w, y_start, mx + col1_w, y_start - total_h)

    # ── Alt footer ─────────────────────────────────────────────
    c.setFillColorRGB(*header_bg)
    c.rect(0, 0, W, 32, fill=1, stroke=0)
    c.setFillColorRGB(*accent_color)
    c.rect(0, 32, W, 2, fill=1, stroke=0)
    c.setFillColorRGB(0.75, 0.85, 1.00)
    c.setFont(_PDF_FONT_NAME, 8)
    c.drawCentredString(W / 2, 11, "Bu belge uygulama tarafından otomatik oluşturulmuştur.")

    c.save()


def build_mail_subject(req_no: str) -> str:
    return f"{MAIL_SUBJECT_PREFIX} | {req_no}"


def build_mail_body(req: "RequestData", log_data: dict) -> str:
    lines = [
        "Merhaba,",
        "",
        "Yeni fikstür talebi oluşturulmuştur.",
        "",
        f"Talep No: {req.request_no}",
        f"Talep Tipi: {req.request_type}",
        f"Proje: {req.project}",
        f"Lens Doküman No: {req.lens_doc}",
        f"Fikstür Adedi: {req.qty}",
        f"Kaplama Sistemi: {req.coating_system}",
        f"Fikstür Malzemesi: {req.material}",
        f"Mesh Uygulaması: {req.mesh}",
        f"Buss Bar Kalınlığı (mm): {req.buss_bar_mm or '-'}",
        f"Aciliyet: {req.urgency}",
        "",
        f"Talep Klasörü: {log_data.get('Talep Klasörü', '')}",
        "",
        "Ekler:",
        "- Yüklenen STP/STEP dosyası (varsa)",
        "- Yüklenen PDF dosyası (varsa)",
        "- Sistem tarafından üretilen talep özeti PDF",
        "",
        "Not:",
        req.note or "-",
        "",
        "Bu e-posta uygulama tarafından otomatik gönderilmiştir.",
    ]
    return "\n".join(lines)


def _auto_classify_titus():
    """
    TITUS sınıflandırma dialogunu arka planda otomatik olarak 'Tasnif Dışı'
    seçerek kapatır.  mail.Send() çağrılmadan önce bir daemon thread olarak
    başlatılmalıdır; dialog görünür görünmez tetiklenir.
    """
    try:
        import win32gui
        import win32con
        import win32api as _wa

        deadline = time.time() + 25  # 25 saniye içinde dialog açılmazsa çık

        while time.time() < deadline:
            time.sleep(0.25)

            matches: list = []

            def _find_top(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if any(kw in title for kw in ("TITUS", "Tasnif", "Classification", "Sınıflandırma")):
                    matches.append(hwnd)
                return True

            win32gui.EnumWindows(_find_top, None)
            if not matches:
                continue

            dlg = matches[0]

            # Alt kontrolleri topla
            items: list = []

            def _find_child(hwnd, _):
                text = win32gui.GetWindowText(hwnd)
                if text:
                    items.append((hwnd, text))
                return True

            try:
                win32gui.EnumChildWindows(dlg, _find_child, None)
            except Exception:
                pass

            # "Tasnif Dışı" seçeneğini tıkla (radio button veya liste öğesi)
            for hwnd, text in items:
                if "Tasnif D" in text:
                    _wa.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
                    time.sleep(0.4)
                    break

            # OK / Tamam / Onayla butonunu bul ve tıkla
            items2: list = []
            try:
                win32gui.EnumChildWindows(
                    dlg,
                    lambda h, _: items2.append((h, win32gui.GetWindowText(h))) or True,
                    None,
                )
            except Exception:
                pass

            for hwnd, text in items2:
                clean = text.strip().replace("&", "")
                if clean in ("OK", "Tamam", "Onayla", "Gönder", "Send", "Devam"):
                    _wa.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
                    return

            # Buton bulunamadıysa Enter tuşu ile kapat
            try:
                win32gui.SetForegroundWindow(dlg)
                _wa.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                time.sleep(0.1)
                _wa.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
            except Exception:
                pass
            return

    except Exception:
        pass  # pywin32 yoksa veya dialog farklı yapıdaysa sessizce geç


def send_outlook_email(to_addr: str, subject: str, body: str, attachments: list[Path]):
    if not IS_WINDOWS:
        raise RuntimeError("Otomatik Outlook gönderimi yalnızca Windows üzerinde kullanılabilir.")

    if win32com is None or pythoncom is None:
        raise RuntimeError(
            "pywin32 modülü bulunamadı. Windows ortamında 'pip install pywin32' kurulumunu yapmalısın."
        )

    pythoncom.CoInitializeEx(0)   # 0 = COINIT_APARTMENTTHREADED (STA) — EXE ortamında gerekli
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_addr
        mail.Subject = subject
        mail.Body = body

        for attachment in attachments:
            if attachment and Path(attachment).exists():
                mail.Attachments.Add(str(attachment))

        # TITUS sınıflandırma dialogu varsa arka planda otomatik 'Tasnif Dışı' seç
        threading.Thread(target=_auto_classify_titus, daemon=True).start()

        mail.Send()
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass



# ========================= STYLES =========================

GLOBAL_QSS = """
QMainWindow {
    background-color: #f0f4f8;
}
QWidget#Central {
    background-color: #f0f4f8;
}
QWidget {
    font-size: 13px;
}

QScrollArea {
    border: none;
    background: transparent;
}
QScrollBar:vertical {
    background-color: #e2e8f0;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background-color: #94a3b8;
    border-radius: 4px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background-color: #2563eb;
}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    height: 0px;
}

QLabel {
    color: #1e293b;
}

QLineEdit {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 9px 12px;
    color: #0f172a;
    font-size: 13px;
    selection-background-color: #2563eb;
    selection-color: #ffffff;
}
QLineEdit:focus {
    border: 2px solid #2563eb;
    padding: 8px 11px;
}
QLineEdit:read-only {
    background-color: #f8fafc;
    color: #64748b;
    border: 1px solid #e2e8f0;
}

QTextEdit {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 8px;
    color: #0f172a;
    font-size: 13px;
}
QTextEdit:focus {
    border: 2px solid #2563eb;
}

QComboBox {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 9px 12px;
    color: #0f172a;
    font-size: 13px;
    min-width: 80px;
}
QComboBox:focus,
QComboBox:on {
    border: 2px solid #2563eb;
    padding: 8px 11px;
}
QComboBox::drop-down {
    subcontrol-origin: border;
    subcontrol-position: center right;
    width: 28px;
    border-left: 1px solid #e2e8f0;
    border-radius: 0 8px 8px 0;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    color: #0f172a;
    selection-background-color: #dbeafe;
    selection-color: #1e40af;
    outline: none;
    padding: 2px;
}

QPushButton {
    background-color: #ffffff;
    color: #1e293b;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 9px 16px;
    font-size: 13px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #f1f5f9;
    border: 1px solid #94a3b8;
}
QPushButton:pressed {
    background-color: #e2e8f0;
}
"""

_PILL_ON = (
    "background-color:{bg}; color:{txt}; border:2px solid {border};"
    "border-radius:10px; font-weight:700; font-size:13px;"
)
_PILL_OFF = (
    "background-color:#f8fafc; color:#94a3b8; border:1px solid #e2e8f0;"
    "border-radius:10px; font-weight:500; font-size:13px;"
)
_PILL_OFF_HV = (
    "QPushButton:hover { background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; }"
)


# ========================= SMALL HELPERS =========================

def _section_row(text: str, icon_widget=None) -> QHBoxLayout:
    """Bölüm başlığı — isteğe bağlı sol ikon + büyük harfli etiket."""
    label = QLabel(text)
    label.setStyleSheet(
        "color:#64748b; font-size:11px; font-weight:700;"
        "letter-spacing:0.5px; background:transparent; border:none;"
    )
    row = QHBoxLayout()
    row.setContentsMargins(0, 0, 0, 0)
    row.setSpacing(6)
    if icon_widget is not None:
        row.addWidget(icon_widget, alignment=Qt.AlignmentFlag.AlignVCenter)
    row.addWidget(label, alignment=Qt.AlignmentFlag.AlignVCenter)
    row.addStretch()
    return row


def _field_label(text: str, center: bool = False) -> QLabel:
    label = QLabel(text)
    label.setStyleSheet("color:#64748b; font-size:12px; background:transparent; border:none;")
    if center:
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    return label


class CenteredItemDelegate(QStyledItemDelegate):
    """Dropdown listesi öğelerini ortalar."""
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter


class CenteredComboBox(QComboBox):
    def paintEvent(self, event):
        painter = QStylePainter(self)
        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)
        painter.drawComplexControl(QStyle.ComplexControl.CC_ComboBox, opt)
        rect = self.style().subControlRect(
            QStyle.ComplexControl.CC_ComboBox,
            opt,
            QStyle.SubControl.SC_ComboBoxEditField,
            self,
        )
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, self.currentText())


def _make_combo(items: list[str]) -> CenteredComboBox:
    cb = CenteredComboBox()
    cb.setItemDelegate(CenteredItemDelegate(cb))
    for item in items:
        cb.addItem(item)
    return cb


# ========================= CUSTOM WIDGETS =========================

class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.setStyleSheet(
            "QFrame#Card { background:#ffffff; border:1px solid #e2e8f0; border-radius:14px; }"
        )

    def apply_layout(self, layout, margins=(18, 16, 18, 18), spacing=12):
        layout.setContentsMargins(*margins)
        layout.setSpacing(spacing)
        self.setLayout(layout)


class SegmentedControl(QWidget):
    selectionChanged = Signal(int)

    def __init__(self, left_text: str, right_text: str, parent=None):
        super().__init__(parent)
        self._sel = 0

        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background-color: #f1f5f9;
                border: 1.5px solid #cbd5e1;
                border-radius: 12px;
            }
        """)

        self.btn_left = QPushButton(left_text)
        self.btn_right = QPushButton(right_text)

        for btn in (self.btn_left, self.btn_right):
            btn.setFixedHeight(38)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            btn.setAutoDefault(False)
            btn.setDefault(False)

        self.btn_left.clicked.connect(lambda: self._select(0))
        self.btn_right.clicked.connect(lambda: self._select(1))

        inner = QHBoxLayout(container)
        inner.setContentsMargins(3, 3, 3, 3)
        inner.setSpacing(3)
        inner.addWidget(self.btn_left)
        inner.addWidget(self.btn_right)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(container)

        self._select(0)

    def _select(self, idx: int):
        self._sel = idx
        self._refresh()
        self.selectionChanged.emit(idx)

    def _refresh(self):
        on_ss = """
            QPushButton {
                background-color:#2563eb; color:#ffffff;
                border:none; border-radius:9px;
                font-weight:700; font-size:13px;
            }
            QPushButton:hover { background-color:#1d4ed8; }
            QPushButton:pressed { background-color:#1e40af; }
        """
        off_ss = """
            QPushButton {
                background-color:transparent; color:#64748b;
                border:none; border-radius:9px;
                font-weight:500; font-size:13px;
            }
            QPushButton:hover { background-color:#e2e8f0; color:#334155; }
        """
        self.btn_left.setStyleSheet(on_ss if self._sel == 0 else off_ss)
        self.btn_right.setStyleSheet(on_ss if self._sel == 1 else off_ss)

    def is_left(self) -> bool:
        return self._sel == 0


class UrgencyBar(QWidget):
    def __init__(self, config: list, parent=None):
        super().__init__(parent)
        self._buttons = []
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)

        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)

        for i, (label, bg, border, txt) in enumerate(config):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setFixedHeight(38)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            btn._bg = bg
            btn._border = border
            btn._txt = txt
            self._group.addButton(btn, i)
            self._buttons.append(btn)
            h.addWidget(btn)
            btn.toggled.connect(self._refresh)

        self._buttons[1].setChecked(True)
        self._refresh()

    def _refresh(self, *_):
        for btn in self._buttons:
            if btn.isChecked():
                btn.setStyleSheet(
                    f"QPushButton {{ {_PILL_ON.format(bg=btn._bg, border=btn._border, txt=btn._txt)} }}"
                )
            else:
                btn.setStyleSheet(f"QPushButton {{ {_PILL_OFF} }} {_PILL_OFF_HV}")

    def selected(self) -> str:
        for btn in self._buttons:
            if btn.isChecked():
                return btn.text()
        return "Normal"

    def reset(self):
        self._buttons[1].setChecked(True)


class PillToggle(QWidget):
    def __init__(self, options: list[str], default_idx: int = 0, on_change=None, parent=None):
        super().__init__(parent)
        self._on_change = on_change
        self._default_idx = default_idx
        self._buttons = []
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)

        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)

        for i, text in enumerate(options):
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.setFixedHeight(36)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self._group.addButton(btn, i)
            self._buttons.append(btn)
            h.addWidget(btn)
            btn.toggled.connect(self._on_toggle)

        self._buttons[default_idx].setChecked(True)
        self._refresh_styles()

    def _on_toggle(self, *_):
        self._refresh_styles()
        if self._on_change:
            self._on_change(self.selected())

    def _refresh_styles(self):
        for btn in self._buttons:
            if btn.isChecked():
                btn.setStyleSheet(
                    f"QPushButton {{ {_PILL_ON.format(bg='#dbeafe', border='#93c5fd', txt='#1e40af')} }}"
                )
            else:
                btn.setStyleSheet(f"QPushButton {{ {_PILL_OFF} }} {_PILL_OFF_HV}")

    def selected(self) -> str:
        for btn in self._buttons:
            if btn.isChecked():
                return btn.text()
        return self._buttons[0].text()

    def reset(self):
        self._buttons[self._default_idx].setChecked(True)


class LensIcon(QWidget):
    """Elektrooptik mercek simgesi — Temel Bilgiler başlığı için."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor("#2563eb")
        p.setPen(QPen(c, 1.4))
        p.setBrush(Qt.NoBrush)
        # Dış daire
        p.drawEllipse(1, 1, 18, 18)
        # Mercek iç eğrileri
        p.drawArc(4, 4, 12, 12, 30 * 16, 120 * 16)
        p.drawArc(4, 4, 12, 12, 210 * 16, 120 * 16)
        p.end()


class ClockIcon(QWidget):
    """Saat simgesi — Aciliyet başlığı için."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor("#b45309")
        p.setPen(QPen(c, 1.4))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(1, 1, 18, 18)
        cx, cy = 10, 10
        # Yelkovan (12'ye doğru)
        p.drawLine(cx, cy, cx, cy - 6)
        # Akrep (3'e doğru)
        p.drawLine(cx, cy, cx + 4, cy)
        p.end()


class NoteIcon(QWidget):
    """Not kağıdı + kalem simgesi — Not başlığı için."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor("#475569")
        p.setPen(QPen(c, 1.3))
        p.setBrush(Qt.NoBrush)
        # Kağıt gövdesi
        p.drawRoundedRect(2, 1, 12, 16, 2, 2)
        # Yatay çizgiler
        p.drawLine(5, 6, 11, 6)
        p.drawLine(5, 9, 11, 9)
        p.drawLine(5, 12, 9, 12)
        # Kalem (sağ alt köşe)
        p.setPen(QPen(c, 1.2))
        p.drawLine(13, 14, 17, 10)
        p.drawLine(13, 14, 15, 16)
        p.drawLine(17, 10, 15, 16)
        p.end()


class FileIcon(QWidget):
    """Dosya simgesi — Dosyalar başlığı için."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor("#0891b2")
        p.setPen(QPen(c, 1.3))
        p.setBrush(Qt.NoBrush)
        # Katlanmış köşeli dosya
        p.drawLine(3, 1, 12, 1)
        p.drawLine(12, 1, 17, 6)
        p.drawLine(17, 6, 17, 19)
        p.drawLine(17, 19, 3, 19)
        p.drawLine(3, 19, 3, 1)
        # Katlanmış köşe üçgeni
        p.drawLine(12, 1, 12, 6)
        p.drawLine(12, 6, 17, 6)
        # İç çizgiler
        p.drawLine(6, 10, 14, 10)
        p.drawLine(6, 13, 14, 13)
        p.end()


class MonitorIcon(QWidget):
    """Monitör simgesi — İşlemler başlığı için."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor("#2563eb")
        p.setPen(QPen(c, 1.3))
        p.setBrush(Qt.NoBrush)
        # Ekran çerçevesi
        p.drawRoundedRect(1, 2, 18, 12, 2, 2)
        # Ekran içi çizgi
        p.drawLine(5, 8, 15, 8)
        # Stand
        p.drawLine(10, 14, 10, 17)
        p.drawLine(7, 17, 13, 17)
        p.end()


class MeshPreview(QWidget):
    """Petek (hex mesh) deseni önizleme simgesi."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(60, 32)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        fill = QColor("#fbbf24")
        fill.setAlpha(35)
        painter.setPen(QPen(QColor("#b45309"), 1.2))
        painter.setBrush(fill)

        r = 5.2
        dx = r * math.sqrt(3)
        dy = r * 1.5

        def hex_poly(cx, cy):
            return QPolygonF([
                QPointF(cx + r * math.cos(math.pi / 6 + i * math.pi / 3),
                        cy + r * math.sin(math.pi / 6 + i * math.pi / 3))
                for i in range(6)
            ])

        centers = [
            (16, 10), (16 + dx, 10), (16 + 2 * dx, 10),
            (16 + dx / 2, 10 + dy), (16 + 1.5 * dx, 10 + dy),
        ]
        for cx, cy in centers:
            if 0 < cx < 62 and 0 < cy < 34:
                painter.drawPolygon(hex_poly(cx, cy))
        painter.end()


class StpIcon(QWidget):
    """CAD/teknik çizim simgesi — STP dosyası için."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(56, 56)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        # Arka plan dairesi
        bg = QColor("#eff6ff")
        p.setPen(Qt.NoPen)
        p.setBrush(bg)
        p.drawEllipse(2, 2, 52, 52)

        # Eksen renkleri (CAD koordinat eksenleri)
        cx, cy = 28, 30

        # Z ekseni (yukarı, gri-mavi)
        p.setPen(QPen(QColor("#64748b"), 1.6))
        p.drawLine(cx, cy, cx, cy - 14)
        p.drawLine(cx, cy - 14, cx - 3, cy - 10)
        p.drawLine(cx, cy - 14, cx + 3, cy - 10)

        # X ekseni (sağ, mavi)
        p.setPen(QPen(QColor("#2563eb"), 1.6))
        p.drawLine(cx, cy, cx + 14, cy)
        p.drawLine(cx + 14, cy, cx + 10, cy - 3)
        p.drawLine(cx + 14, cy, cx + 10, cy + 3)

        # Y ekseni (sol-aşağı perspektif, camgöbeği)
        p.setPen(QPen(QColor("#0891b2"), 1.6))
        ex = int(cx - 10 * math.cos(math.radians(30)))
        ey = int(cy + 10 * math.sin(math.radians(30)))
        p.drawLine(cx, cy, ex, ey)

        # Küçük izometrik kutu (wireframe)
        p.setPen(QPen(QColor("#2563eb"), 1.1))
        p.setBrush(Qt.NoBrush)
        # üst yüzey
        top = QPolygonF([
            QPointF(cx, cy - 8),
            QPointF(cx + 7, cy - 4),
            QPointF(cx, cy),
            QPointF(cx - 7, cy - 4),
        ])
        p.drawPolygon(top)
        # alt kenarlıklar
        p.drawLine(int(cx - 7), int(cy - 4), int(cx - 7), int(cy + 4))
        p.drawLine(int(cx + 7), int(cy - 4), int(cx + 7), int(cy + 4))
        p.drawLine(int(cx - 7), int(cy + 4), int(cx), int(cy + 8))
        p.drawLine(int(cx + 7), int(cy + 4), int(cx), int(cy + 8))
        p.drawLine(int(cx), int(cy), int(cx), int(cy + 8))
        p.end()


class DropZone(QFrame):
    _SS_EMPTY = "QFrame#DropZone { background:#f8fafc; border:2px dashed #cbd5e1; border-radius:14px; }"
    _SS_HOVER = "QFrame#DropZone { background:#eff6ff; border:2px dashed #2563eb; border-radius:14px; }"
    _SS_FILLED = "QFrame#DropZone { background:#f0fdf4; border:2px solid #86efac; border-radius:14px; }"

    def __init__(self, icon, title: str, hint: str, accept_exts: tuple[str, ...], parent=None):
        """
        icon: str (emoji) veya QWidget (custom icon widget)
        """
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.accept_exts = tuple(e.lower() for e in accept_exts)
        self.file_path: Path | None = None

        self.setAcceptDrops(True)
        self.setMinimumHeight(145)
        self.setStyleSheet(self._SS_EMPTY)

        label_style = "background:transparent; border:none;"

        # icon: string emoji veya custom widget
        if isinstance(icon, str):
            self._icon_text = icon
            self._icon_widget = QLabel(icon)
            self._icon_widget.setAlignment(Qt.AlignCenter)
            self._icon_widget.setStyleSheet(f"font-size:28px; {label_style}")
            self._icon_is_label = True
        else:
            self._icon_text = None
            self._icon_widget = icon
            self._icon_is_label = False

        self._title_lbl = QLabel(title)
        self._title_lbl.setAlignment(Qt.AlignCenter)
        self._title_lbl.setStyleSheet(
            f"font-size:13px; font-weight:600; color:#475569; {label_style}"
        )

        self._hint_lbl = QLabel(hint)
        self._hint_lbl.setAlignment(Qt.AlignCenter)
        self._hint_lbl.setWordWrap(True)
        self._hint_lbl.setStyleSheet(f"font-size:11px; color:#94a3b8; {label_style}")

        self._file_lbl = QLabel("")
        self._file_lbl.setAlignment(Qt.AlignCenter)
        self._file_lbl.setWordWrap(True)
        self._file_lbl.setStyleSheet(
            f"font-size:12px; color:#16a34a; font-weight:600; {label_style}"
        )
        self._file_lbl.hide()

        self._ok_lbl = QLabel("✅")
        self._ok_lbl.setAlignment(Qt.AlignCenter)
        self._ok_lbl.setStyleSheet(f"font-size:28px; {label_style}")
        self._ok_lbl.hide()

        self._pick_btn = QPushButton("Seç…")
        self._pick_btn.setFixedSize(88, 32)
        self._pick_btn.clicked.connect(self.pick_file)

        self._clear_btn = QPushButton("✕")
        self._clear_btn.setFixedSize(32, 32)
        self._clear_btn.setStyleSheet("""
            QPushButton {
                background:#fee2e2;
                color:#dc2626;
                border:1px solid #fca5a5;
                border-radius:8px;
                font-weight:700;
                font-size:12px;
                padding:0;
            }
            QPushButton:hover { background:#fecaca; }
        """)
        self._clear_btn.clicked.connect(self.clear_file)
        self._clear_btn.hide()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addStretch()
        btn_row.addWidget(self._pick_btn)
        btn_row.addWidget(self._clear_btn)
        btn_row.addStretch()

        icon_row = QHBoxLayout()
        icon_row.setContentsMargins(0, 0, 0, 0)
        icon_row.addStretch()
        icon_row.addWidget(self._icon_widget)
        icon_row.addWidget(self._ok_lbl)
        icon_row.addStretch()

        v = QVBoxLayout(self)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(5)
        v.addStretch()
        v.addLayout(icon_row)
        v.addWidget(self._title_lbl)
        v.addWidget(self._hint_lbl)
        v.addWidget(self._file_lbl)
        v.addSpacing(8)
        v.addLayout(btn_row)
        v.addStretch()

    def _valid(self, p: Path) -> bool:
        if not p.is_file():
            return False
        # Standart uzantı kontrolü (örn: parca.stp)
        if p.suffix.lower() in self.accept_exts:
            return True
        # CREO versiyonlu dosyalar: parca.prt.3 → suffixes = ['.prt', '.3']
        return any(s.lower() in self.accept_exts for s in p.suffixes)

    def set_file(self, p: Path | None):
        self.file_path = p
        if p:
            self._file_lbl.setText(f"✓  {p.name}   ({format_size(p)})")
            self._file_lbl.show()
            self._hint_lbl.hide()
            self._icon_widget.hide()
            self._ok_lbl.show()
            self._clear_btn.show()
            self.setStyleSheet(self._SS_FILLED)
        else:
            self._file_lbl.hide()
            self._hint_lbl.show()
            self._ok_lbl.hide()
            self._icon_widget.show()
            self._clear_btn.hide()
            self.setStyleSheet(self._SS_EMPTY)

    def clear_file(self):
        self.set_file(None)

    def pick_file(self):
        exts_str = " ".join(f"*{e}" for e in self.accept_exts)
        filt = f"Desteklenen Dosyalar ({exts_str});;Tüm Dosyalar (*.*)"
        fname, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", str(Path.home()), filt)
        if fname:
            p = Path(fname)
            if not self._valid(p):
                QMessageBox.warning(
                    self,
                    "Hata",
                    f"Geçersiz dosya türü.\nKabul edilenler: {', '.join(self.accept_exts)}",
                )
                return
            self.set_file(p)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.mimeData().hasText():
            self.setStyleSheet(self._SS_HOVER)
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() or event.mimeData().hasText():
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self._SS_FILLED if self.file_path else self._SS_EMPTY)

    def _path_from_event(self, event) -> "Path | None":
        """MIME verisinden dosya yolunu çıkarır — URL veya metin formatını destekler."""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                return Path(urls[0].toLocalFile())
        if event.mimeData().hasText():
            text = event.mimeData().text().strip().strip('"\'')
            candidate = Path(text)
            if candidate.exists():
                return candidate
        return None

    def dropEvent(self, event):
        p = self._path_from_event(event)
        if p is None:
            return

        if not self._valid(p):
            QMessageBox.warning(
                self,
                "Hata",
                f"Geçersiz dosya türü.\nKabul edilenler: {', '.join(self.accept_exts)}",
            )
            self.setStyleSheet(self._SS_FILLED if self.file_path else self._SS_EMPTY)
            return

        self.set_file(p)


# ========================= DATA MODEL =========================

@dataclass
class RequestData:
    request_type: str
    request_no: str
    original_no: str | None
    project: str
    lens_doc: str
    qty: int
    coating_system: str
    material: str
    mesh: str
    buss_bar_mm: str
    urgency: str
    note: str


# ========================= MAIN WINDOW =========================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} — {APP_VERSION}")
        self.setMinimumSize(QSize(1080, 780))
        self.setStyleSheet(GLOBAL_QSS)

        central = QWidget()
        central.setObjectName("Central")
        self.setCentralWidget(central)

        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        outer.addWidget(self._build_header())
        outer.addWidget(self._build_scroll(), 1)
        outer.addWidget(self._build_bottom_bar())

        # Debounce timer — disk taramasını harf başına değil, yazmayı bırakınca yap
        self._reqno_timer = QTimer(self)
        self._reqno_timer.setSingleShot(True)
        self._reqno_timer.setInterval(400)
        self._reqno_timer.timeout.connect(self._do_refresh_reqno)

        self._refresh_reqno()

        dir_action = QAction("Kayıt Dizini…", self)
        dir_action.triggered.connect(self.show_basedir_info)
        self.menuBar().addAction(dir_action)

    # ---------------- HEADER ----------------

    def _build_header(self) -> QWidget:
        bar = QWidget()
        bar.setFixedHeight(72)
        bar.setStyleSheet(
            "QWidget { background-color:#1e3a5f; border-bottom:1px solid #1a3356; }"
        )

        h = QHBoxLayout(bar)
        h.setContentsMargins(22, 0, 22, 0)
        h.setSpacing(14)

        title = QLabel(APP_TITLE)
        title.setFont(QFont("Arial", 17, QFont.Bold))
        title.setStyleSheet("color:#ffffff; background:transparent; border:none;")

        sub = QLabel("OKTM · Optik Kaplama Tasarım Müdürlüğü")
        sub.setStyleSheet(
            "color:rgba(255,255,255,0.60); font-size:12px; background:transparent; border:none;"
        )

        left = QVBoxLayout()
        left.setSpacing(2)
        left.setContentsMargins(0, 0, 0, 0)
        left.addStretch()
        left.addWidget(title)
        left.addWidget(sub)
        left.addStretch()

        no_lbl = QLabel("Talep No")
        no_lbl.setStyleSheet(
            "color:rgba(255,255,255,0.55); font-size:11px; background:transparent; border:none;"
        )

        self.reqno_badge = QLabel("—")
        self.reqno_badge.setAlignment(Qt.AlignCenter)
        self.reqno_badge.setFixedHeight(36)
        self.reqno_badge.setMinimumWidth(220)

        self._badge_ss = (
            "background:rgba(255,255,255,0.15); color:#ffffff;"
            "font-size:14px; font-weight:700;"
            "border:1px solid rgba(255,255,255,0.35);"
            "border-radius:10px; padding:0 14px;"
        )
        self._badge_flash_ss = (
            "background:rgba(255,255,255,0.38); color:#1e3a5f;"
            "font-size:14px; font-weight:700;"
            "border:1px solid rgba(255,255,255,0.75);"
            "border-radius:10px; padding:0 14px;"
        )
        self.reqno_badge.setStyleSheet(self._badge_ss)

        info_btn = QPushButton("📁")
        info_btn.setToolTip("Kayıt dizini bilgisi")
        info_btn.setFixedSize(36, 36)
        info_btn.setStyleSheet("""
            QPushButton {
                background:rgba(255,255,255,0.12);
                border:1px solid rgba(255,255,255,0.28);
                border-radius:8px;
                font-size:16px;
                color:#ffffff;
                padding:0;
            }
            QPushButton:hover { background:rgba(255,255,255,0.25); }
            QPushButton:pressed { background:rgba(255,255,255,0.38); }
        """)
        info_btn.clicked.connect(self.show_basedir_info)

        right = QVBoxLayout()
        right.setSpacing(2)
        right.setContentsMargins(0, 0, 0, 0)
        right.addStretch()
        right.addWidget(no_lbl, alignment=Qt.AlignHCenter)
        right.addWidget(self.reqno_badge)
        right.addStretch()

        h.addLayout(left)
        h.addStretch()
        h.addLayout(right)
        h.addWidget(info_btn, alignment=Qt.AlignVCenter)

        return bar

    # ---------------- BODY ----------------

    def _build_scroll(self) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        container = QWidget()
        container.setStyleSheet("background-color:#f0f4f8;")
        scroll.setWidget(container)

        h = QHBoxLayout(container)
        h.setContentsMargins(20, 20, 20, 20)
        h.setSpacing(16)

        h.addWidget(self._build_left_panel(), 55)
        h.addWidget(self._build_right_panel(), 45)
        return scroll

    def _build_left_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet("background:transparent;")

        v = QVBoxLayout(panel)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)

        v.addWidget(self._card_type())
        v.addWidget(self._card_fields())
        v.addWidget(self._card_urgency())
        v.addWidget(self._card_note())
        v.addStretch()
        return panel

    def _build_right_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet("background:transparent;")

        v = QVBoxLayout(panel)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)

        v.addWidget(self._card_files())
        v.addWidget(self._card_actions())
        v.addStretch()
        return panel

    # ---------------- CARDS ----------------

    def _card_type(self) -> Card:
        card = Card()
        v = QVBoxLayout()

        v.addLayout(_section_row("TALEP TİPİ"))
        v.addSpacing(4)

        self.seg = SegmentedControl("  Yeni Fikstür  ", "  Rework  ")
        self.seg.selectionChanged.connect(lambda _: self._on_mode_changed())
        v.addWidget(self.seg)

        self.rw_container = QWidget()
        self.rw_container.setStyleSheet("background:transparent;")
        rw_v = QVBoxLayout(self.rw_container)
        rw_v.setContentsMargins(0, 8, 0, 0)
        rw_v.setSpacing(6)

        rw_v.addWidget(_field_label("Orijinal Talep No"))
        self.ed_orig = QLineEdit()
        self.ed_orig.setPlaceholderText("Örn: Proje-Lens-1")
        self.ed_orig.setAlignment(Qt.AlignCenter)
        self.ed_orig.textChanged.connect(self._refresh_reqno)
        rw_v.addWidget(self.ed_orig)

        self.rw_container.hide()
        v.addWidget(self.rw_container)

        card.apply_layout(v)
        return card

    def _card_fields(self) -> Card:
        card = Card()
        grid = QGridLayout()

        grid.addLayout(_section_row("TEMEL BİLGİLER", LensIcon()), 0, 0, 1, 2)

        grid.addWidget(_field_label("Proje", center=True), 1, 0)
        self.ed_project = QLineEdit()
        self.ed_project.setPlaceholderText("Proje adı")
        self.ed_project.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.ed_project.textChanged.connect(self._on_project_changed)
        grid.addWidget(self.ed_project, 2, 0)

        grid.addWidget(_field_label("Lens Doküman No", center=True), 1, 1)
        self.ed_lens = QLineEdit()
        self.ed_lens.setPlaceholderText("Lens doküman no")
        self.ed_lens.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.ed_lens.textChanged.connect(self._refresh_reqno)
        grid.addWidget(self.ed_lens, 2, 1)

        grid.addWidget(_field_label("Fikstür Adedi", center=True), 3, 0)
        self.cb_qty = _make_combo(QTY_OPTIONS)
        grid.addWidget(self.cb_qty, 4, 0)

        grid.addWidget(_field_label("Kaplama Sistemi", center=True), 3, 1)
        self.cb_coat = _make_combo(COATING_SYSTEMS)
        grid.addWidget(self.cb_coat, 4, 1)

        grid.addWidget(_field_label("Fikstür Malzemesi", center=True), 5, 0, 1, 2)
        self.cb_mat = _make_combo(MATERIALS)
        grid.addWidget(self.cb_mat, 6, 0, 1, 2)

        mesh_lbl = _field_label("Mesh Uygulaması", center=True)
        mesh_title_row = QHBoxLayout()
        mesh_title_row.setContentsMargins(0, 0, 0, 0)
        mesh_title_row.setSpacing(4)
        mesh_title_row.addStretch()
        mesh_title_row.addWidget(mesh_lbl)
        mesh_title_row.addWidget(MeshPreview(), alignment=Qt.AlignmentFlag.AlignVCenter)
        mesh_title_row.addStretch()
        grid.addLayout(mesh_title_row, 7, 0, 1, 2)

        self.mesh_toggle = PillToggle(["Yok", "Var", "Bilinmiyor"], default_idx=0, on_change=self._on_mesh_changed)
        grid.addWidget(self.mesh_toggle, 8, 0, 1, 2)

        self.buss_lbl = _field_label("Buss Bar Kalınlığı (mm)", center=True)
        grid.addWidget(self.buss_lbl, 9, 0, 1, 2)

        self.ed_buss_bar = QLineEdit()
        self.ed_buss_bar.setPlaceholderText("Örn: 5")
        self.ed_buss_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.ed_buss_bar.setEnabled(False)
        grid.addWidget(self.ed_buss_bar, 10, 0, 1, 2)

        card.apply_layout(grid)
        return card

    def _card_urgency(self) -> Card:
        card = Card()
        v = QVBoxLayout()

        v.addLayout(_section_row("ACİLİYET", ClockIcon()))
        self.urgency_bar = UrgencyBar(URGENCY_CONFIG)
        v.addWidget(self.urgency_bar)

        card.apply_layout(v)
        return card

    def _card_note(self) -> Card:
        card = Card()
        v = QVBoxLayout()

        v.addLayout(_section_row("NOT", NoteIcon()))
        self.txt_note = QTextEdit()
        self.txt_note.setPlaceholderText("İsteğe bağlı açıklama, üretim notu, özel durumlar...")
        self.txt_note.setMinimumHeight(140)
        v.addWidget(self.txt_note)

        card.apply_layout(v)
        return card

    def _card_files(self) -> Card:
        card = Card()
        v = QVBoxLayout()

        v.addLayout(_section_row("DOSYALAR", FileIcon()))

        self.drop_stp = DropZone(
            StpIcon(),
            "STP / STEP / PRT Dosyası",
            "Sürükle bırak yapabilir veya dosya seçebilirsin.",
            (".stp", ".step", ".prt"),
        )
        self.drop_pdf = DropZone(
            "📄",
            "PDF Dosyası",
            "Sürükle bırak yapabilir veya dosya seçebilirsin.",
            (".pdf",),
        )

        v.addWidget(self.drop_stp)
        v.addWidget(self.drop_pdf)

        card.apply_layout(v)
        return card

    def _card_actions(self) -> Card:
        card = Card()
        v = QVBoxLayout()

        v.addLayout(_section_row("İŞLEMLER", MonitorIcon()))

        self.status_lbl = QLabel("Hazır.")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setStyleSheet(
            "background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px; color:#334155;"
        )
        v.addWidget(self.status_lbl)

        _warn_ss = (
            "background:#fffbeb; border:1px solid #fde68a; border-radius:8px;"
            "padding:10px; color:#92400e; font-size:11px;"
        )
        warn1 = QLabel(
            "⚠  Mesh uygulaması var ise EMI mesh kalınlık bilgilerini "
            "Mikro Desenleme Birimi'ne danışın."
        )
        warn1.setWordWrap(True)
        warn1.setStyleSheet(_warn_ss)
        v.addWidget(warn1)

        warn2 = QLabel(
            "ℹ  Optik elemanlardaki pah kısmını lütfen Optik Tasarım "
            "birimleriyle teyit ediniz."
        )
        warn2.setWordWrap(True)
        warn2.setStyleSheet(_warn_ss)
        v.addWidget(warn2)

        btn_row = QHBoxLayout()

        self.btn_reset = QPushButton("Formu Sıfırla")
        self.btn_reset.clicked.connect(self.reset_form)

        self.btn_submit = QPushButton("Kaydet ve Gönder")
        self.btn_submit.setStyleSheet("""
            QPushButton {
                background:#2563eb;
                color:white;
                border:none;
                border-radius:10px;
                padding:10px 18px;
                font-weight:700;
            }
            QPushButton:hover { background:#1d4ed8; }
            QPushButton:pressed { background:#1e40af; }
        """)
        self.btn_submit.clicked.connect(self.on_submit)

        btn_row.addWidget(self.btn_reset)
        btn_row.addWidget(self.btn_submit)

        v.addLayout(btn_row)

        card.apply_layout(v)
        return card

    # ---------------- FOOTER ----------------

    def _build_bottom_bar(self) -> QWidget:
        bar = QWidget()
        bar.setFixedHeight(52)
        bar.setStyleSheet("""
            QWidget {
                background:#ffffff;
                border-top:1px solid #dbe3ec;
            }
        """)

        h = QHBoxLayout(bar)
        h.setContentsMargins(18, 0, 18, 0)

        lbl = QLabel(FOOTER_TEXT)
        lbl.setStyleSheet("color:#64748b; font-size:12px; background:transparent; border:none;")

        right = QLabel(APP_VERSION)
        right.setStyleSheet("color:#94a3b8; font-size:12px; background:transparent; border:none;")

        h.addWidget(lbl)
        h.addStretch()
        h.addWidget(right)
        return bar

    # ---------------- LOGIC ----------------

    def _on_project_changed(self):
        field = self.ed_project
        text = field.text()
        upper = text.upper()
        if text != upper:
            pos = field.cursorPosition()
            field.blockSignals(True)
            field.setText(upper)
            field.setCursorPosition(pos)
            field.blockSignals(False)
        self._refresh_reqno()

    def _on_mode_changed(self):
        is_new = self.seg.is_left()
        self.rw_container.setVisible(not is_new)
        self._refresh_reqno()

    def _on_mesh_changed(self, value: str):
        if not hasattr(self, "ed_buss_bar"):
            return
        enabled = (value == "Var")
        self.ed_buss_bar.setEnabled(enabled)
        if not enabled:
            self.ed_buss_bar.clear()

    def _refresh_reqno(self):
        """Anlık UI güncellemesi — disk taraması yok, sadece "..." göster."""
        if self.seg.is_left():
            project = self.ed_project.text().strip() if hasattr(self, "ed_project") else ""
            lens    = self.ed_lens.text().strip() if hasattr(self, "ed_lens") else ""

            if project and lens:
                self.reqno_badge.setText(f"{sanitize_name(project)}-{sanitize_name(lens)}-…")
                if hasattr(self, "_reqno_timer"):
                    self._reqno_timer.start()   # 400ms sonra disk taraması
            elif project:
                self.reqno_badge.setText(f"{sanitize_name(project)}-…-?")
                if hasattr(self, "_reqno_timer"):
                    self._reqno_timer.stop()
            else:
                self.reqno_badge.setText("—")
                if hasattr(self, "_reqno_timer"):
                    self._reqno_timer.stop()
                return
        else:
            orig = self.ed_orig.text().strip() if hasattr(self, "ed_orig") else ""
            self.reqno_badge.setText(f"RW-{orig}" if orig else "RW-…")

        self.reqno_badge.setStyleSheet(self._badge_flash_ss)
        QTimer.singleShot(280, lambda: self.reqno_badge.setStyleSheet(self._badge_ss))

    def _do_refresh_reqno(self):
        """Disk taraması yaparak gerçek numara hesaplar (debounced)."""
        if not self.seg.is_left():
            return
        project = self.ed_project.text().strip() if hasattr(self, "ed_project") else ""
        lens    = self.ed_lens.text().strip() if hasattr(self, "ed_lens") else ""
        if project and lens:
            self.reqno_badge.setText(next_request_no(BASE_DIR, project, lens))
            self.reqno_badge.setStyleSheet(self._badge_flash_ss)
            QTimer.singleShot(280, lambda: self.reqno_badge.setStyleSheet(self._badge_ss))

    def set_status(self, text: str, kind: str = "info"):
        styles = {
            "info": "background:#eff6ff; border:1px solid #bfdbfe; color:#1e3a8a;",
            "success": "background:#ecfdf5; border:1px solid #a7f3d0; color:#065f46;",
            "error": "background:#fef2f2; border:1px solid #fecaca; color:#991b1b;",
        }
        self.status_lbl.setStyleSheet(
            f"border-radius:10px; padding:12px; {styles.get(kind, styles['info'])}"
        )
        self.status_lbl.setText(text)

    def show_basedir_info(self):
        QMessageBox.information(
            self,
            "Kayıt Dizini",
            f"Tüm talepler şu dizine kaydedilir:\n\n{BASE_DIR}",
        )

    def build_request(self) -> RequestData | None:
        req_type = "Yeni" if self.seg.is_left() else "Rework"
        project = self.ed_project.text().strip()
        lens = self.ed_lens.text().strip()
        qty = int(self.cb_qty.currentText())
        coat = self.cb_coat.currentText()
        mat = self.cb_mat.currentText()
        urg = self.urgency_bar.selected()
        mesh = self.mesh_toggle.selected()
        buss_bar = self.ed_buss_bar.text().strip() if mesh == "Var" else ""
        note = self.txt_note.toPlainText().strip()

        if not project:
            self.set_status("❌  Proje ismi zorunludur.", "error")
            return None

        if not lens:
            self.set_status("❌  Lens doküman numarası zorunludur.", "error")
            return None

        if mesh == "Var" and not buss_bar:
            self.set_status("❌  Mesh 'Var' seçildiğinde Buss Bar Kalınlığı zorunludur.", "error")
            return None

        if req_type == "Yeni":
            req_no, orig = next_request_no(BASE_DIR, project, lens), None
        else:
            orig = self.ed_orig.text().strip()
            if not orig:
                self.set_status("❌  Rework için Orijinal Talep No zorunludur.", "error")
                return None
            req_no = f"RW-{orig}"

        return RequestData(
            request_type=req_type,
            request_no=req_no,
            original_no=orig,
            project=project,
            lens_doc=lens,
            qty=qty,
            coating_system=coat,
            material=mat,
            mesh=mesh,
            buss_bar_mm=buss_bar,
            urgency=urg,
            note=note,
        )

    def on_submit(self):
        try:
            BASE_DIR.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.set_status(f"❌  Kayıt dizini oluşturulamadı:\n{e}", "error")
            return

        req = self.build_request()
        if req is None:
            return

        self.set_status("⏳  Kaydediliyor…", "info")
        QApplication.processEvents()

        folder_name = sanitize_name(req.request_no)
        req_dir = BASE_DIR / folder_name

        try:
            req_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.set_status(f"❌  Talep klasörü oluşturulamadı:\n{e}", "error")
            return

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        stp_dest = None
        uploaded_pdf_dest = None

        try:
            if self.drop_stp.file_path:
                stp_dest = copy_into_folder(self.drop_stp.file_path, req_dir, folder_name)
            if self.drop_pdf.file_path:
                uploaded_pdf_dest = copy_into_folder(self.drop_pdf.file_path, req_dir, folder_name)
        except Exception as e:
            self.set_status(f"❌  Dosya kopyalama hatası:\n{e}", "error")
            return

        log_data = {
            "Talep No": req.request_no,
            "Talep Tipi": req.request_type,
            "Zaman Damgası": ts,
            "Proje": req.project,
            "Lens Doküman No": req.lens_doc,
            "Fikstür Adedi": req.qty,
            "Kaplama Sistemi": req.coating_system,
            "Fikstür Malzemesi": req.material,
            "Mesh Uygulaması": req.mesh,
            "Buss Bar Kalınlığı mm": req.buss_bar_mm,
            "Aciliyet": req.urgency,
            "Not": req.note,
            "STP Dosyası": "" if stp_dest is None else str(stp_dest),
            "PDF Dosyası": "" if uploaded_pdf_dest is None else str(uploaded_pdf_dest),
            "Talep Klasörü": str(req_dir),
        }

        summary_pdf_path = req_dir / f"{folder_name}.pdf"

        try:
            append_to_master_log(BASE_DIR, log_data)
            write_pdf(summary_pdf_path, "Fikstür Talep Özeti", log_data)
        except Exception as e:
            self.set_status(f"❌  Kayıt hatası:\n{e}", "error")
            return

        self.set_status("⏳  Outlook maili gönderiliyor…", "info")
        QApplication.processEvents()

        attachments = [p for p in [stp_dest, uploaded_pdf_dest, summary_pdf_path] if p is not None]

        try:
            subject = build_mail_subject(req.request_no)
            body = build_mail_body(req, log_data)
            send_outlook_email(MAIL_TO, subject, body, attachments)
        except Exception as e:
            self.set_status(
                f"⚠️  Talep kaydedildi ancak mail gönderilemedi.\n"
                f"Klasör: {req_dir.name}\n"
                f"Hata: {e}",
                "error",
            )
            return

        self.set_status(
            f"✅  Talep başarıyla kaydedildi ve mail gönderildi.\n"
            f"Klasör: {req_dir.name}\n"
            f"Log: {LOG_FILENAME}",
            "success"
        )
        # TODO: Jira entegrasyonu — talep oluşturulduğunda otomatik Jira task aç

    def reset_form(self):
        self.seg._select(0)
        self.ed_orig.clear()
        self.ed_project.clear()
        self.ed_lens.clear()
        self.cb_qty.setCurrentIndex(0)
        self.cb_coat.setCurrentIndex(0)
        self.cb_mat.setCurrentIndex(0)
        self.mesh_toggle.reset()
        self.ed_buss_bar.clear()
        self.ed_buss_bar.setEnabled(False)
        self.urgency_bar.reset()
        self.txt_note.clear()
        self.drop_stp.clear_file()
        self.drop_pdf.clear_file()
        self.set_status("Hazır.", "info")
        self._refresh_reqno()


# ========================= ENTRY =========================

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()