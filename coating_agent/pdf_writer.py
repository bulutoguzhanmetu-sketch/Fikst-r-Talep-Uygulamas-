#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Rapor metnini sade bir PDF dosyasına döker."""

from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas

PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 20 * mm
FOOTER_TEXT = "Optik Kaplama & Malzeme Mühendisi Araştırma Ajanı"


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
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = rl_canvas.Canvas(str(output_path), pagesize=A4)
    max_width = PAGE_WIDTH - 2 * MARGIN

    def new_page(first: bool = False) -> float:
        if not first:
            c.showPage()
        c.setFont("Helvetica", 8)
        c.drawString(MARGIN, 10 * mm, FOOTER_TEXT)
        return PAGE_HEIGHT - MARGIN

    y = new_page(first=True)

    c.setFont("Helvetica-Bold", 16)
    for line in _wrap_line(c, topic, "Helvetica-Bold", 16, max_width):
        c.drawString(MARGIN, y, line)
        y -= 8 * mm
    y -= 4 * mm

    c.setFont("Helvetica", 11)
    body_font, body_size, leading = "Helvetica", 11, 6 * mm

    for paragraph in report_text.split("\n"):
        if not paragraph.strip():
            y -= leading / 2
            continue
        for line in _wrap_line(c, paragraph, body_font, body_size, max_width):
            if y < MARGIN + 10 * mm:
                y = new_page()
                c.setFont(body_font, body_size)
            c.drawString(MARGIN, y, line)
            y -= leading

    c.save()
