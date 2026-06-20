#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Rapor metnini sade bir PDF dosyasına döker (Türkçe karakter desteğiyle)."""

from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 20 * mm
FOOTER_TEXT = "Optik Kaplama & Malzeme Mühendisi Araştırma Ajanı"

# Helvetica'nın WinAnsi kodlaması Türkçe'ye özgü ı, ş, ğ, İ harflerini
# içermez; bu yüzden Unicode destekli bir TTF font (DejaVu Sans) kaydedilir.
_FONT_REGULAR = "Helvetica"
_FONT_BOLD = "Helvetica-Bold"
_FONT_REGISTERED = False


def _register_turkish_font() -> None:
    global _FONT_REGULAR, _FONT_BOLD, _FONT_REGISTERED
    if _FONT_REGISTERED:
        return

    candidates = [
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
         Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
        (Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
         Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf")),
        (Path("/Library/Fonts/Arial.ttf"),
         Path("/Library/Fonts/Arial Bold.ttf")),
    ]

    for regular_path, bold_path in candidates:
        if regular_path.exists():
            try:
                pdfmetrics.registerFont(TTFont("ReportBody", str(regular_path)))
                _FONT_REGULAR = "ReportBody"
                if bold_path.exists():
                    pdfmetrics.registerFont(TTFont("ReportBody-Bold", str(bold_path)))
                    _FONT_BOLD = "ReportBody-Bold"
                else:
                    _FONT_BOLD = "ReportBody"
                _FONT_REGISTERED = True
                return
            except Exception:
                continue

    _FONT_REGISTERED = True  # font bulunamadıysa Helvetica ile devam edilir


def _wrap_line(c: rl_canvas.Canvas, text: str, font: str, size: int, max_width: float) -> list[str]:
    words = text.split(" ")
    lines, current = [], ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if c.stringWidth(candidate, font, size) <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def write_report_pdf(topic: str, report_text: str, output_path: Path) -> None:
    _register_turkish_font()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = rl_canvas.Canvas(str(output_path), pagesize=A4)
    max_width = PAGE_WIDTH - 2 * MARGIN

    def new_page(first: bool = False) -> float:
        if not first:
            c.showPage()
        c.setFont(_FONT_REGULAR, 8)
        c.drawString(MARGIN, 10 * mm, FOOTER_TEXT)
        return PAGE_HEIGHT - MARGIN

    y = new_page(first=True)

    c.setFont(_FONT_BOLD, 16)
    for line in _wrap_line(c, topic, _FONT_BOLD, 16, max_width):
        c.drawString(MARGIN, y, line)
        y -= 8 * mm
    y -= 4 * mm

    body_size, leading = 11, 6 * mm
    c.setFont(_FONT_REGULAR, body_size)

    for paragraph in report_text.split("\n"):
        if not paragraph.strip():
            y -= leading / 2
            continue

        stripped = paragraph.strip()
        is_heading = (
            len(stripped) < 60
            and stripped[:1].isdigit()
            and ". " in stripped[:5]
        )
        font_for_line = _FONT_BOLD if is_heading else _FONT_REGULAR

        for line in _wrap_line(c, paragraph, font_for_line, body_size, max_width):
            if y < MARGIN + 10 * mm:
                y = new_page()
                c.setFont(_FONT_REGULAR, body_size)
            c.setFont(font_for_line, body_size)
            c.drawString(MARGIN, y, line)
            y -= leading

    c.save()
