# -*- coding: utf-8 -*-
"""
consultation_script_pdf_builder.py — оборачивает build_script_sections() из
consultation_script_builder.py в PDF: рабочий документ ДЛЯ ИГОРЯ (не для
клиента), чтобы удобно листать/печатать на консультации.

Использует тот же шрифт (Roboto) и ту же палитру, что и pdf_report_builder.py
(Отчёт) и pptx_build/base.js (Презентация) — чтобы все 3 артефакта одной
консультации визуально были частью одного семейства.

Отличия от Отчёта (pdf_report_builder.py):
- Не клиентский документ — можно компактнее, без обложки-титула для клиента.
- Маркеры [→ ПЕРЕКЛЮЧИТЬ НА СЛАЙД N: «...»] выделены цветной плашкой — это
  рабочая инструкция Игорю, должна бросаться в глаза при беглом чтении.
- Реплики клиента (в кавычках «...») выделены курсивом/цветом — чтобы легко
  отличать "что сказать" от "что процитировать".
"""
import re
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether, PageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BASE = Path(__file__).parent
ASSETS = BASE / "assets"

# --- Шрифты (те же файлы, что уже используются в pdf_report_builder.py) ---
pdfmetrics.registerFont(TTFont("Roboto", str(ASSETS / "fonts" / "Roboto-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Roboto-Bold", str(ASSETS / "fonts" / "Roboto-Bold.ttf")))
pdfmetrics.registerFont(TTFont("Roboto-Black", str(ASSETS / "fonts" / "Roboto-Black.ttf")))

# --- Палитра — как в Отчёте и Презентации ---
NAVY = colors.HexColor("#0b3041")
ORANGE = colors.HexColor("#D5530B")
GOLD = colors.HexColor("#FCA700")
TAUPE = colors.HexColor("#B3927E")
TEAL = colors.HexColor("#1A7573")
INK = colors.HexColor("#22303A")
LIGHT_BG = colors.HexColor("#EEF3F5")
QUOTE_COLOR = colors.HexColor("#5A6B74")

MARGIN = 18 * mm

styles = getSampleStyleSheet()
styles.add(ParagraphStyle("ScriptTitle", parent=styles["Normal"], fontName="Roboto-Black",
                           fontSize=20, textColor=NAVY, spaceAfter=10))
styles.add(ParagraphStyle("ScriptSubtitle", parent=styles["Normal"], fontName="Roboto",
                           fontSize=11, textColor=QUOTE_COLOR, spaceAfter=14))
styles.add(ParagraphStyle("SectionHeader", parent=styles["Normal"], fontName="Roboto-Bold",
                           fontSize=13.5, textColor=colors.white, leading=16))
styles.add(ParagraphStyle("PartHeader", parent=styles["Normal"], fontName="Roboto-Bold",
                           fontSize=11, textColor=TEAL, spaceBefore=10, spaceAfter=4))
styles.add(ParagraphStyle("Body", parent=styles["Normal"], fontName="Roboto", fontSize=10,
                           leading=14, textColor=INK, spaceAfter=3, alignment=TA_LEFT))
styles.add(ParagraphStyle("ScriptBullet", parent=styles["Body"], leftIndent=10, spaceAfter=2))
styles.add(ParagraphStyle("SlideMarker", parent=styles["Normal"], fontName="Roboto-Bold",
                           fontSize=9, textColor=colors.white, leading=12))
styles.add(ParagraphStyle("Numbered", parent=styles["Body"], leftIndent=10))
styles.add(ParagraphStyle("ObjectionPlaque", parent=styles["Normal"], fontName="Roboto-Bold",
                           fontSize=10.5, textColor=colors.white, leading=13))


def _escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _inline_format(text):
    """«...» -> курсив/приглушённый цвет (реплики клиента / реплики для
    произнесения вслух), **...** уже не используется в тексте Скрипта, но на
    всякий случай поддержано."""
    text = _escape(text)
    text = re.sub(r'«([^»]*)»', rf'«<i><font color="#5A6B74">\1</font></i>»', text)
    return text


def _is_slide_marker(line):
    return line.strip().startswith("[→") or line.strip().startswith("[⚠")


def _is_bullet(line):
    return line.strip().startswith("•")


def _is_numbered(line):
    return bool(re.match(r'^\s*\d+[.)]\s', line))


def _is_part_header(line):
    return line.strip().startswith("---") and line.strip().endswith("---")


def render_section_body(text, flowables, numbered_as_plaque=False):
    """Превращает уже сформированный (reflow'нутый) текст Раздела в список
    Flowable — по тем же правилам разметки, что и reflow() в
    consultation_script_builder.py (заголовки/маркеры/буллеты/списки).

    numbered_as_plaque: True только для Раздела 6 «Возражения» — там строки
    вида "N. «Название возражения»" — это заголовки блоков, а не элементы
    списка (в отличие, например, от нумерованных цитат клиента в Разделе 4),
    поэтому оформляются как цветная (оранжевая) плашка, а не обычным текстом."""
    for raw_line in text.split("\n"):
        line = raw_line.rstrip()
        stripped = line.strip()

        if not stripped:
            flowables.append(Spacer(1, 4))
            continue

        if _is_part_header(stripped):
            title = stripped.strip("- ").strip()
            flowables.append(Spacer(1, 6))
            flowables.append(Paragraph(_escape(title), styles["PartHeader"]))
            flowables.append(HRFlowable(width="100%", thickness=0.6, color=TEAL,
                                         spaceAfter=6, spaceBefore=0))
            continue

        if _is_slide_marker(stripped):
            label = stripped.strip("[]")
            t = Table([[Paragraph(_escape(label), styles["SlideMarker"])]],
                      colWidths=[None])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), NAVY if stripped.startswith("[→") else ORANGE),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            flowables.append(Spacer(1, 5))
            flowables.append(t)
            flowables.append(Spacer(1, 4))
            continue

        if _is_bullet(stripped):
            content = stripped.lstrip("•").strip()
            flowables.append(Paragraph("•  " + _inline_format(content), styles["ScriptBullet"]))
            continue

        if _is_numbered(stripped):
            if numbered_as_plaque:
                label = re.sub(r'^\d+[.)]\s*', '', stripped)
                num_prefix = re.match(r'^(\d+[.)])', stripped).group(1)
                plaque_text = _escape(f"{num_prefix} {label}").upper()
                t = Table([[Paragraph(plaque_text, styles["ObjectionPlaque"])]], colWidths=[None])
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), ORANGE),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]))
                flowables.append(Spacer(1, 8))
                flowables.append(t)
                flowables.append(Spacer(1, 5))
            else:
                flowables.append(Paragraph(_inline_format(stripped), styles["Numbered"]))
            continue

        if stripped.startswith("## "):
            flowables.append(Spacer(1, 4))
            flowables.append(Paragraph(_escape(stripped[3:]).upper(),
                                        ParagraphStyle("Tier", parent=styles["Body"],
                                                       fontName="Roboto-Bold", fontSize=10.5,
                                                       textColor=ORANGE)))
            continue

        if stripped.startswith("### "):
            flowables.append(Paragraph(_escape(stripped[4:]),
                                        ParagraphStyle("KseName", parent=styles["Body"],
                                                       fontName="Roboto-Bold", fontSize=10.5,
                                                       textColor=NAVY, spaceBefore=6)))
            continue

        if stripped.startswith("**") and stripped.endswith("**"):
            flowables.append(Paragraph(f'<b>{_escape(stripped.strip("*"))}</b>', styles["Body"]))
            continue

        flowables.append(Paragraph(_inline_format(stripped), styles["Body"]))


def draw_section_band(canvas_obj, doc, title, page_state):
    """Не используется как onPage — заголовок Раздела рисуется как обычный
    Flowable (Table с цветным фоном), см. build_pdf()."""
    pass


def section_band(title):
    t = Table([[Paragraph(_escape(title), styles["SectionHeader"])]], colWidths=[None])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), TEAL),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def build_pdf(sections, qualification, stage_name, output_path):
    """sections: результат build_script_sections() — [(title, text), ...].
    qualification/stage_name: те же dict/строка, что уже передавались в
    build_full_script() — для титульной строки. output_path: путь к файлу
    (str/Path) ИЛИ файлоподобный объект с .write() (например io.BytesIO —
    так вызывает app.py, без записи на диск, для одноразового ответа
    веб-запроса)."""
    doc = SimpleDocTemplate(
        output_path if hasattr(output_path, "write") else str(output_path),
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN, topMargin=MARGIN, bottomMargin=MARGIN,
        title=f"Скрипт консультации — {qualification.get('name', '')}",
    )

    story = []
    story.append(Paragraph("СКРИПТ КОНСУЛЬТАЦИИ", styles["ScriptTitle"]))
    subtitle = (f'{qualification.get("name", "")} · {qualification.get("company", "")} · '
                f'{stage_name} · Отчёт №{qualification.get("report_number", "")}')
    story.append(Paragraph(_escape(subtitle), styles["ScriptSubtitle"]))
    story.append(HRFlowable(width="100%", thickness=1.2, color=NAVY, spaceAfter=12))

    for title, text in sections:
        if title.startswith("РАЗДЕЛ 3."):
            story.append(PageBreak())
        section_flowables = [section_band(title), Spacer(1, 8)]
        render_section_body(text, section_flowables, numbered_as_plaque=title.startswith("РАЗДЕЛ 6."))
        section_flowables.append(Spacer(1, 14))
        # Заголовок Раздела не должен остаться в самом низу страницы одиноко —
        # держим его вместе с первыми парой строк содержимого.
        story.append(KeepTogether(section_flowables[:3]))
        story.extend(section_flowables[3:])

    def _footer(canvas_obj, doc_obj):
        canvas_obj.saveState()
        canvas_obj.setFont("Roboto", 7.5)
        canvas_obj.setFillColor(QUOTE_COLOR)
        canvas_obj.drawString(MARGIN, 10 * mm, "ЛАБОРАТОРИЯ БИЗНЕС ЛИДЕРСТВА «ФЕНИКС» — "
                                                 "рабочий документ, не для клиента")
        canvas_obj.drawRightString(A4[0] - MARGIN, 10 * mm, f"Стр. {doc_obj.page}")
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return output_path
