"""
Сборка полного PDF-отчёта «Полная оценка состояния бизнеса».

Использует:
- scoring_algorithm.py — расчёт диагностики
- report_text_generator.py — персонализированные текстовые блоки (Разделы 1, 8, 11, 12)
- data/stage_level_report_texts.json — стандартные тексты по Стадии (Раздел 1 РЕЗЮМЕ, Разделы 2-6)
- assets/ — логотипы (PNG, отрендерены из SVG) и шрифт Roboto

ИЗВЕСТНОЕ УПРОЩЕНИЕ (см. итоговое сообщение в чате): для Разделов 2-6 логика
выбора нужной формулировки внутри блока ВЫВОДЫ (когда там несколько условных
подветок вида "ЕСЛИ ВЫШЕ / ЕСЛИ НИЖЕ") сейчас упрощена — показывается полный
текст найденной ветки (соответствует/не соответствует) целиком, включая все
подветки, а не только ту, что точно соответствует направлению расхождения
клиента. Точная логика выбора конкретной подветки — отдельная задача.
"""

import json
import re
from functools import partial
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image,
    PageBreak, KeepTogether, Flowable, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth

import scoring_algorithm as sa
import report_text_generator as rtg

BASE = Path(__file__).parent
ASSETS = BASE / "assets"

REPORT_COUNTER_START = 355  # следующий номер после последнего уже выданного (354) до переноса на постоянный диск


def get_next_report_number():
    """Возвращает следующий по порядку номер отчёта в формате 'ХХХХХХ' (6 цифр
    с ведущими нулями).

    20.08.2026, миграция Render → Timeweb Cloud: раньше счётчик хранился в
    файле на постоянном диске Render (/var/data) — на Timeweb App Platform
    такого диска нет (см. подробный разбор в app.py про ORDERS_FILE), поэтому
    счётчик перенесён в ту же базу PostgreSQL (fenix_orders / DATABASE_URL),
    что уже используется для хранения заказов Prodamus.

    Атомарность инкремента обеспечивается самой PostgreSQL: используется
    последовательность (SEQUENCE) — при параллельных запросах СУБД
    гарантированно выдаёт каждому вызову свой, уникальный, последовательный
    номер, без явных блокировок в коде приложения (в отличие от файловой
    версии, где для этого требовался fcntl.flock).

    Первый вызов возвращает REPORT_COUNTER_START. Каждый следующий — на 1
    больше предыдущего.
    """
    import os
    import psycopg

    database_url = os.environ.get("DATABASE_URL", "")
    if not database_url:
        raise RuntimeError("DATABASE_URL не задана в переменных окружения")

    with psycopg.connect(database_url) as conn:
        with conn.cursor() as cur:
            # Последовательность создаётся один раз при первом вызове, если
            # её ещё нет — стартует с REPORT_COUNTER_START.
            # ВАЖНО: CREATE SEQUENCE — это DDL-команда, PostgreSQL не
            # поддерживает для неё параметризованные значения (%s) в
            # START WITH — значение нужно подставлять прямо в текст SQL.
            # Это безопасно: REPORT_COUNTER_START — целочисленная константа
            # из кода, а не пользовательский ввод, риска SQL-инъекции нет.
            cur.execute(
                f"CREATE SEQUENCE IF NOT EXISTS report_number_seq "
                f"START WITH {REPORT_COUNTER_START}"
            )
            cur.execute("SELECT nextval('report_number_seq')")
            next_num = cur.fetchone()[0]
        conn.commit()
    return f"{next_num:06d}"

# ---------------------------------------------------------------------------
# Шрифты и цвета бренда
# ---------------------------------------------------------------------------

pdfmetrics.registerFont(TTFont("Roboto", str(ASSETS / "fonts" / "Roboto-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Roboto-Bold", str(ASSETS / "fonts" / "Roboto-Bold.ttf")))
pdfmetrics.registerFont(TTFont("Roboto-Black", str(ASSETS / "fonts" / "Roboto-Black.ttf")))
pdfmetrics.registerFont(TTFont("DejaVuSans", str(ASSETS / "fonts" / "DejaVuSans.ttf")))  # символ ☑, отсутствующий в Roboto

NAVY = colors.HexColor("#0b3041")
ORANGE = colors.HexColor("#D5530B")
GOLD = colors.HexColor("#FCA700")
TAUPE = colors.HexColor("#B3927E")
TEAL = colors.HexColor("#1A7573")
WHITE = colors.white
INK = colors.HexColor("#22303A")

# светофор для % выполнения Непреложных правил
ALT_VARIANT_TEXT_INTRO = (
    "В Вашем плане со списком Консалтинговых программ, рекомендуемых к внедрению, есть "
    "несколько Программ, которые входят в Комплексную консалтинговую программу "
    "«Возрождение малого бизнеса»:"
)
ALT_VARIANT_PATH_A = (
    "Путь А – полные программы по отдельности. Каждая Программа даёт максимальную глубину "
    "проработки своего Системного элемента – от диагностики конкретно Вашей ситуации до "
    "полного внедрения и обкатки в рабочем процессе."
)
ALT_VARIANT_PATH_B = (
    "Путь Б - Комплексная консалтинговая программа «Возрождение малого бизнеса». Это не "
    "эквивалент 6 Программ полностью, а комплексный пакет из наиболее критичных модулей "
    "каждой из 6 Программ:<br/>"
    "• Бизнес-модель<br/>"
    "• Ценности бренда и Базовые ценности<br/>"
    "• Структура Развития бизнеса<br/>"
    "• Организационная структура<br/>"
    "• Ключевые показатели эффективности<br/>"
    "• Коучинговое управление персоналом<br/>"
    "Если бы Вы проходили эти Программы последовательно как отдельные полные Программы, "
    "это заняло бы порядка 73 недель. Комплексная программа даёт компактную версию тех же "
    "Системных элементов за 25 недель. Такой путь особенно подходит, если ситуация в "
    "бизнесе требует быстрой стабилизации по многим фронтам одновременно, а не "
    "углублённой проработки каждого Системного элемента."
)
ALT_VARIANT_BRIDGE = "Это открывает для Вас выбор между двумя способами, как двигаться дальше:"
ALT_VARIANT_QUESTIONS = [
    "Какой путь лучше подходит именно вам?",
    "Как превратить результаты диагностики в работающий план с конкретными сроками, "
    "ресурсами и списком дальнейших шагов?",
    "Мы с вами вместе ответим на эти и другие вопросы на онлайн-встрече.",
]

TRAFFIC = {
    100: colors.HexColor("#2E7D4F"), 80: colors.HexColor("#2E7D4F"),
    60: colors.HexColor("#FCA700"),
    40: colors.HexColor("#C0392B"), 20: colors.HexColor("#C0392B"), 0: colors.HexColor("#C0392B"),
}

styles = getSampleStyleSheet()
styles.add(ParagraphStyle("Body", parent=styles["Normal"], fontName="Roboto", fontSize=10.5,
                           leading=15, textColor=INK, spaceAfter=8))
styles.add(ParagraphStyle("StyleTitle", parent=styles["Normal"], fontName="Roboto-Black", fontSize=12.5,
                           leading=16, textColor=INK, spaceAfter=5))
styles.add(ParagraphStyle("BodyBold", parent=styles["Body"], fontName="Roboto-Bold"))
styles.add(ParagraphStyle("H1", parent=styles["Normal"], fontName="Roboto-Bold", fontSize=20,
                           leading=24, textColor=NAVY, spaceBefore=4, spaceAfter=10))
styles.add(ParagraphStyle("H2", parent=styles["Normal"], fontName="Roboto-Bold", fontSize=14,
                           leading=18, textColor=NAVY, spaceBefore=14, spaceAfter=8))
styles.add(ParagraphStyle("H3", parent=styles["Normal"], fontName="Roboto-Bold", fontSize=13,
                           leading=17, textColor=ORANGE, spaceBefore=10, spaceAfter=4))
styles.add(ParagraphStyle("Small", parent=styles["Body"], fontSize=8.5, leading=12,
                           textColor=colors.HexColor("#5B6B72")))
styles.add(ParagraphStyle("SmallBody", parent=styles["Body"], fontSize=9, leading=12.5))
styles.add(ParagraphStyle("CoverTitle", parent=styles["Normal"], fontName="Roboto-Bold", fontSize=16,
                           leading=20, textColor=NAVY, alignment=TA_CENTER))
styles.add(ParagraphStyle("CoverCompany", parent=styles["Normal"], fontName="Roboto-Bold", fontSize=15,
                           leading=19, textColor=WHITE, alignment=TA_CENTER))
styles.add(ParagraphStyle("CoverMeta", parent=styles["Normal"], fontName="Roboto", fontSize=12,
                           leading=17, textColor=colors.HexColor("#D8E4E8"), alignment=TA_CENTER))


class MarkerGlyph(Flowable):
    """Рисует символ-маркер (☑ / ● и т.п.) с точным позиционированием по
    фактическим метрикам глифа шрифта — не по номинальному ascent шрифта,
    который для мелких символов (точка, чекбокс) обычно намного выше
    реальных «чернил» и создаёт лишний зазор при выравнивании по верху."""
    def __init__(self, symbol, font_name, size, color, box_width, box_height, valign="TOP"):
        super().__init__()
        self.symbol = symbol
        self.font_name = font_name
        self.size = size
        self.color = color
        self.box_width = box_width
        self.box_height = box_height
        self.valign = valign

    def wrap(self, availWidth, availHeight):
        return self.box_width, self.box_height

    def draw(self):
        c = self.canv
        c.setFillColor(colors.HexColor(self.color) if isinstance(self.color, str) else self.color)
        c.setFont(self.font_name, self.size)
        # реальная верхняя/нижняя граница «чернил» первого символа строки (по font.glyf)
        y_max_pt, y_min_pt = _glyph_ink_bounds_pt(self.font_name, self.symbol, self.size)
        if self.valign == "TOP":
            baseline_y = self.box_height - y_max_pt
        elif self.valign == "BOTTOM":
            baseline_y = -y_min_pt
        else:  # MIDDLE
            ink_height = y_max_pt - y_min_pt
            baseline_y = (self.box_height - ink_height) / 2 - y_min_pt
        text_width = pdfmetrics.stringWidth(self.symbol, self.font_name, self.size)
        x = (self.box_width - text_width) / 2
        c.drawString(x, baseline_y, self.symbol)


_glyph_bounds_cache = {}


def _glyph_ink_bounds_pt(font_name, symbol, size):
    """Возвращает (yMax, yMin) в пунктах — реальную вертикальную протяжённость
    «чернил» первого символа строки relative к baseline, для данного шрифта/размера."""
    from fontTools.ttLib import TTFont as _FTFont
    key = (font_name, symbol[0])
    if key not in _glyph_bounds_cache:
        font_path = str(ASSETS / "fonts" / f"{font_name}.ttf") if font_name != "Roboto" else str(
            ASSETS / "fonts" / "Roboto-Regular.ttf")
        ft = _FTFont(font_path)
        cmap = ft.getBestCmap()
        upm = ft["head"].unitsPerEm
        glyph_name = cmap[ord(symbol[0])]
        glyf = ft["glyf"][glyph_name]
        _glyph_bounds_cache[key] = (glyf.yMax / upm, glyf.yMin / upm)
    y_max_frac, y_min_frac = _glyph_bounds_cache[key]
    return y_max_frac * size, y_min_frac * size


class TwoLineBanner(Flowable):
    """Плашка-заголовок с двумя строками (напр. название блока + ФИО клиента),
    как на образце «ЕСТЕСТВЕННОЕ СОЧЕТАНИЕ СТИЛЕЙ УПРАВЛЕНИЯ / ФИО»."""
    def __init__(self, line1, line2, width=None, height=17 * mm, bg_color=None):
        super().__init__()
        self.line1 = line1
        self.line2 = line2
        self.width = width
        self.height = height
        self.bg_color = bg_color or NAVY

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.bg_color)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=0)
        c.setFillColor(WHITE)
        c.setFont("Roboto-Bold", 12.5)
        c.drawCentredString(self.width / 2, self.height - 7 * mm, self.line1)
        c.setFont("Roboto", 10)
        c.drawCentredString(self.width / 2, self.height - 13 * mm, self.line2)


class SetPageFlag(Flowable):
    """Flowable нулевого размера — не влияет на вёрстку, только устанавливает
    флаг в момент отрисовки. Используется, чтобы сообщить onLaterPages,
    что СЛЕДУЮЩАЯ страница — особая (например, закрывающая, с другим фоном)."""
    def __init__(self, state_dict, key, value=True):
        super().__init__()
        self.state_dict = state_dict
        self.key = key
        self.value = value

    def wrap(self, availWidth, availHeight):
        return 0, 0

    def draw(self):
        self.state_dict[self.key] = self.value


class SectionBanner(Flowable):
    """Плашка-заголовок раздела (тёмная полоса с белым текстом), как в примере отчёта."""
    def __init__(self, text, width=None, height=13 * mm, bg_color=None):
        super().__init__()
        self.text = text
        self.width = width
        self.height = height
        self.bg_color = bg_color or NAVY

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.bg_color)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=0)
        c.setFillColor(WHITE)
        c.setFont("Roboto-Bold", 13)
        c.drawString(6 * mm, self.height / 2 - 4, self.text)


FOOTER_LOGO = str(ASSETS / "logos" / "Vertical_color_original.png")  # подлинный полноцветный оригинал (Vertical.svg — монохромная трассировка без цвета в самом файле, см. пояснение)
COVER_LOGO = str(ASSETS / "logos" / "Horizontal_clean.png")  # чистая версия без тэглайна, от пользователя (Horizontal_without_tagline.svg)
COVER_PHOTO = str(ASSETS / "Обложка.png")
PLATE_GRAY = colors.HexColor("#BFBFBF")

# Пропорции обложки — все размеры заданы в % от страницы (ширина/высота A4)
COVER_PLATE_WIDTH_PCT = 0.75      # ширина плашки — 75% ширины страницы
COVER_PLATE_HEIGHT_PCT = 0.28     # высота плашки — 28% высоты страницы
COVER_PLATE_TOP_PCT = 0.075       # отступ от верха страницы до плашки — 7.5% высоты страницы
COVER_LOGO_WIDTH_PCT = 0.28       # ширина логотипа — 28% ширины страницы
COVER_TITLE_LINES = [
    "АВТОРСКАЯ ОЦЕНКА", "СОСТОЯНИЯ БИЗНЕСА", "ПО СИСТЕМЕ", "«ВОЗРОЖДЕНИЕ БИЗНЕСА»",
]
COVER_SUBTITLE = "Отчёт о диагностике"


def draw_cover(canvas, doc, meta):
    """Полностью отрисовывает обложку абсолютными координатами — без
    flowable-вёрстки, чтобы расположение элементов было предсказуемым и
    воспроизводило макет «Макет_новой_обложки.png» с точностью до мм."""
    canvas.saveState()
    page_w, page_h = A4

    # 1. Фон — сплошной фирменный синий на всю страницу (без белой рамки по периметру)
    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, page_w, page_h, fill=1, stroke=0)

    # 2. Серая плашка с заголовком
    plate_w = page_w * COVER_PLATE_WIDTH_PCT
    plate_h = page_h * COVER_PLATE_HEIGHT_PCT
    plate_x = (page_w - plate_w) / 2
    plate_top_y = page_h - page_h * COVER_PLATE_TOP_PCT   # y от низа страницы (координаты canvas — снизу вверх)
    plate_bottom_y = plate_top_y - plate_h

    canvas.setFillColor(PLATE_GRAY)
    canvas.setStrokeColor(WHITE)
    canvas.setLineWidth(1.5)
    canvas.rect(plate_x, plate_bottom_y, plate_w, plate_h, fill=1, stroke=1)

    title_size = 28
    title_leading = 32
    subtitle_size = 20
    subtitle_leading = 24

    block_h = len(COVER_TITLE_LINES) * title_leading + 6 * mm + subtitle_leading
    block_top = plate_bottom_y + plate_h / 2 + block_h / 2  # верх блока текста, центрированного в плашке

    canvas.setFillColor(NAVY)
    canvas.setFont("Roboto-Black", title_size)
    y = block_top - title_leading * 0.8
    for line in COVER_TITLE_LINES:
        canvas.drawCentredString(page_w / 2, y, line)
        y -= title_leading

    y -= 6 * mm - (title_leading - subtitle_leading)
    canvas.setFont("Roboto", subtitle_size)
    canvas.drawCentredString(page_w / 2, y, COVER_SUBTITLE)

    # 3. Компания / ФИО / дата — под плашкой, белым, по центру
    name_size = 15  # компания и ФИО — одинаковый размер
    date_size = name_size - 2
    name_leading = 19
    date_leading = 16

    text_top = plate_bottom_y - (12 * mm + 3 + 2)
    canvas.setFillColor(WHITE)
    canvas.setFont("Roboto-Bold", name_size)
    canvas.drawCentredString(page_w / 2, text_top, meta["company"])
    canvas.setFont("Roboto-Bold", name_size)
    canvas.drawCentredString(page_w / 2, text_top - name_leading, meta["name"])
    canvas.setFont("Roboto", date_size)
    canvas.drawCentredString(page_w / 2, text_top - name_leading - date_leading, meta["diagnosis_date"])

    text_block_bottom = text_top - name_leading - date_leading - 4 * mm

    # 4. Фото — во всю ширину страницы, натуральные пропорции, БЕЗ логотипа поверх
    photo_w = page_w
    photo_h = photo_w * (855 / 1521)
    logo_w = page_w * COVER_LOGO_WIDTH_PCT
    logo_h = logo_w * (492 / 2400)
    bottom_strip = 6 * mm + logo_h + 6 * mm  # синяя полоса внизу под логотип (фото в неё не заходит)

    photo_bottom_y = bottom_strip
    # если между текстом и нижней синей полосой не хватает места — фото прижимается вверх под текстовый блок
    photo_top_y = photo_bottom_y + photo_h
    if photo_top_y > text_block_bottom:
        photo_top_y = text_block_bottom
        photo_bottom_y = photo_top_y - photo_h

    canvas.drawImage(COVER_PHOTO, 0, photo_bottom_y, width=photo_w, height=photo_h, mask="auto")

    # 5. Логотип — в правом нижнем углу обложки, НА СИНЕМ ФОНЕ (ниже фото, не поверх него)
    logo_x = page_w - logo_w - 14 * mm
    logo_y = (bottom_strip - logo_h) / 2
    canvas.drawImage(COVER_LOGO, logo_x, logo_y, width=logo_w, height=logo_h,
                      mask="auto", preserveAspectRatio=True)

    canvas.restoreState()


def draw_footer(canvas, doc, meta, page_state=None):
    """Футер внутренних страниц: логотип-иконка слева, текст (номер отчёта /
    дата диагностики / компания) отцентрован по ширине страницы, номер
    страницы — у самого правого края. Если page_state сигнализирует, что это
    закрывающая страница — сначала заливает всю страницу фирменным синим."""
    canvas.saveState()
    page_w, page_h = A4

    is_closing = bool(page_state and page_state.get("closing"))
    if is_closing:
        canvas.setFillColor(NAVY)
        canvas.rect(0, 0, page_w, page_h, fill=1, stroke=0)
        canvas.restoreState()
        return

    logo_h = 6 * mm * 2.5
    logo_w_img = logo_h * (1344 / 896)
    logo_x = doc.leftMargin
    logo_y = 4 * mm
    canvas.drawImage(FOOTER_LOGO, logo_x, logo_y, width=logo_w_img, height=logo_h,
                      mask="auto", preserveAspectRatio=True)

    line_y = logo_y + logo_h + 2 * mm
    canvas.setStrokeColor(colors.HexColor("#D8D2C4"))
    canvas.setLineWidth(0.4)
    canvas.line(doc.leftMargin, line_y, page_w - doc.rightMargin, line_y)

    canvas.setFont("Roboto", 7.5)
    canvas.setFillColor(colors.HexColor("#5B6B72"))
    footer_text = f'Отчёт № {meta["report_number"]}   |   Дата диагностики: {meta["diagnosis_date"]}   |   {meta["company"]}'
    text_baseline = logo_y + logo_h / 2 - 2.6
    canvas.drawCentredString(page_w / 2, text_baseline, footer_text)
    canvas.drawRightString(page_w - 8 * mm, text_baseline, str(canvas.getPageNumber()))

    canvas.restoreState()


def section_header(number, title):
    return [SectionBanner(f"{number}   {title}".upper()), Spacer(1, 6 * mm)]


def para(text, style="Body"):
    text = text.replace("\n\n", "<br/><br/>").replace("\n", "<br/>")
    return Paragraph(text, styles[style])


def header_with_body(header_text, body_text, header_style="H3", body_style="Body"):
    """H3-заголовок + следующий за ним контент, гарантированно на одной
    странице (KeepTogether) — без этого заголовок мог остаться внизу
    страницы в одиночестве, а сам контент "оторваться" на следующую
    (обнаружено на реальном Отчёте 05.08.2026, исправлено везде по файлу).
    body_text: либо готовый Flowable (Table и т.п.), либо текст — если текст
    содержит несколько абзацев (разделены \\n\\n), с заголовком группируется
    только ПЕРВЫЙ абзац (остальные добавляются в story отдельно, обычным
    потоком, вызывающей стороной) — чтобы не заставлять ReportLab держать
    вместе весь длинный блок целиком (что может создать куда более странные
    разрывы, чем исходная проблема)."""
    if hasattr(body_text, "wrap"):  # уже Flowable (Table, Paragraph, ...)
        return KeepTogether([para(header_text, header_style), body_text])
    return KeepTogether([para(header_text, header_style), para(body_text, body_style)])


def append_header_with_paragraphs(story, header_text, full_text, header_style="H3", body_style="Body"):
    """Как header_with_body(), но для текста из нескольких абзацев (\\n\\n) —
    заголовок группируется только с первым абзацем, остальные добавляются
    в story как обычно (см. docstring header_with_body)."""
    blocks = [b for b in full_text.split("\n\n") if b.strip()]
    if not blocks:
        story.append(para(header_text, header_style))
        return
    story.append(KeepTogether([para(header_text, header_style), para(blocks[0], body_style)]))
    for block in blocks[1:]:
        story.append(para(block, body_style))


def markdown_lite_to_flowables(text):
    """Простая конвертация текста из report_text_generator.py (## / ### / **bold**) во flowables."""
    blocks = [b.strip() for b in text.split("\n\n") if b.strip()]
    flow = []
    i = 0
    while i < len(blocks):
        block = blocks[i]
        if block.startswith("## "):
            flow.append(para(block[3:], "H2"))
        elif block.startswith("### "):
            # Заголовок группируется со следующим блоком, чтобы не остаться
            # в одиночестве внизу страницы при разбиении на страницы.
            if i + 1 < len(blocks) and not blocks[i + 1].startswith(("## ", "### ")):
                flow.append(KeepTogether([para(block[4:], "H3"), para(blocks[i + 1], "Body")]))
                i += 1
            else:
                flow.append(para(block[4:], "H3"))
        else:
            block = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', block)
            flow.append(para(block, "Body"))
        i += 1
    return flow


# ---------------------------------------------------------------------------
# Данные демонстрационного клиента (Стадия 3)
# ---------------------------------------------------------------------------

def build_demo_client(data):
    challenges = list(data["classic_challenges_top5_by_stage"].keys())
    scores = {ch: 4 for ch in challenges}
    for ch in ["Неясные Базовые ценности", "Культура компании не принимает изменения",
               "Отсутствие поддержки со стороны персонала", "Разрыв между руководством и сотрудниками",
               "Слабо проработанная бизнес-модель"]:
        scores[ch] = 8

    client = {
        "qualification": {
            "name": "Виктор Громов", "company": "ООО «Прибрежный Дом»",
            "diagnosis_date": "13.07.2026",  # в реальности — timestamp момента отправки опроса, не ручной ввод
            "report_number": get_next_report_number(),  # авто-инкремент, начиная с 000350 — см. get_next_report_number()
            "fte_a": 25, "fte_b": 0, "fte_c": 0, "fte_d": 0,
            "managers_actual": 1, "leaders_actual": 1,
            "employeesYearAgo": 18, "timeYears": 1, "timeMonths": 4,
            "years_in_business": 3,
        },
        "flow_a_dimensions": {
            "priority_spheres": ["Прибыль", "Люди", "Процессы"],  # не совпадает с целевым Стадии 3
            "builder_protector_ratio": "2:1",  # не совпадает с целевым 1:1
            "modality": data["rules_of_growth_targets"]["modality"]["3"],  # совпадает
            "management_styles": {"Основной": "Директивный", "Второстепенный": "Эталонный", "Дополнительный": "Коучинговый"},
            "management_styles_scores": {
                "Директивный": 16, "Эталонный": 15, "Коучинговый": 14,
                "Авторитетный": 6, "Демократический": 5, "Товарищеский": 4,
            },
            "three_leader_roles": data["rules_of_growth_targets"]["three_leader_roles"]["3"],  # совпадает
        },
        "immutable_rules_pct": {},
        "challenge_scores": scores,
        "section8_likert_by_kse": {
            "Организационная структура": [2], "Ценности бренда и Базовые ценности": [2],
            "Критерии роста бизнеса": [3], "Бизнес-модель": [3],
        },
    }
    rules_struct = data["immutable_rules"]["3"]
    for area, rules in rules_struct.items():
        pcts = []
        for i in range(len(rules)):
            pcts.append(20 if i == 0 else 80)
        client["immutable_rules_pct"][area] = pcts
    return client


def format_target_range(rng):
    """[3, 5] -> '3-5'; [1, 1] -> '1' (одиночное целевое значение)."""
    lo, hi = rng
    return str(lo) if lo == hi else f"{lo}-{hi}"


def format_zone_entry(full_stage):
    ze = full_stage.get("zone_entry")
    if not ze:
        return "-"
    return f'{ze["type"]} (до {ze["range"][1]} сотрудников)'


def format_zone_exit(full_stage):
    zx = full_stage.get("zone_exit")
    if not zx:
        return "-"
    return f'{zx["type"]} (нач. с {zx["range"][0]} сотрудников)'


def format_current_zone(stage):
    """stage — результат determine_zone() из scoring_algorithm (уже содержит zone_name/zone_type)."""
    if stage["zone_name"] == "Функциональная зона":
        return "Функциональная"
    return stage["zone_type"]


def find_critical_gaps(diagnose_result, data, threshold=50):
    stage_id = diagnose_result["стадия"]["stage_id"]
    result = []
    for r in diagnose_result["поток_а"]["failed_immutable_rules"]:
        if r["стадия_появления"] < stage_id and r["факт_%"] < threshold:
            result.append(r)
    return result


# ---------------------------------------------------------------------------
# Сборка Story (список flowables)
# ---------------------------------------------------------------------------

def build_story(client, result, data):
    story = []
    page_state = {"closing": False}
    company = client["qualification"]["company"]
    name = client["qualification"]["name"]
    stage = result["стадия"]
    stage_id = stage["stage_id"]
    s = str(stage_id)

    # ---------- Обложка ----------
    # Вся вёрстка обложки рисуется абсолютными координатами в draw_cover()
    # (передаётся как onFirstPage в doc.build) — здесь только запускаем страницу.
    story.append(Spacer(1, 1))
    story.append(PageBreak())

    # ---------- Стр. 2: методология + данные ----------
    story.append(SectionBanner("АВТОРСКАЯ ОЦЕНКА СОСТОЯНИЯ БИЗНЕСА"))
    story.append(Spacer(1, 6 * mm))
    story.append(para(
        "Авторская оценка состояния бизнеса по методологии системы «Возрождение бизнеса» "
        "использует десяток аналитических измерений, чтобы показать Вам, что происходит с "
        "Вашим бизнесом, и предоставить Вам конкретные идеи для приоритизации дальнейших действий."
    ))
    story.append(para(
        "Данный диагностический инструмент базируется на результатах 15-ти летних практических "
        "исследований более 1,500 компаний из самых разных отраслей."
    ))
    story.append(para(
        "Результаты диагностики, представленные в этом отчёте, позволят определить, какие "
        "Системные элементы отсутствуют в Вашем бизнесе. Вы также получите рекомендации, как "
        "адресовать основные вызовы, с которыми сталкивается Ваша организация. И Вы сможете "
        "понять, с чем ещё Ваш бизнес может столкнуться в процессе дальнейшего роста."
    ))
    story.append(Spacer(1, 8 * mm))

    story.append(SectionBanner("ВОЗРОЖДЕНИЕ ВАШЕГО БИЗНЕСА"))
    story.append(Spacer(1, 6 * mm))
    story.append(para(
        "Проведение оценки состояния Вашего бизнеса – это первый шаг к возрождению бизнеса по "
        "нашей авторской системе роста бизнеса."
    ))
    story.append(para(
        "Возрождение бизнеса – это система роста, которая определяет недостающие Системные "
        "элементы бизнеса и внедряет их в экосистему организации. Именно благодаря появлению "
        "недостающих Ключевых системных элементов экосистема бизнеса переходит в состояние "
        "устойчивости, жизнеспособности и динамичного развития. В результате бизнес становится "
        "высокопроизводительным и высокорентабельным."
    ))
    story.append(para(
        "Подобно законам живой природы, методология системы «Возрождение бизнеса» базируется "
        "на фундаментальных принципах построения человеческих организаций и поэтому применяется "
        "к любому бизнесу вне зависимости от его размера, отрасли или особенностей личности "
        "владельца."
    ))
    story.append(Spacer(1, 6 * mm))

    # Дата диагностики — НЕ вычисляется здесь. В реальном пайплайне это
    # timestamp, зафиксированный системой автоматически в момент отправки
    # опроса клиентом (клик «Завершить диагностику» в фронтенде), который
    # проходит через Make.com вместе с остальными ответами и передаётся сюда
    # как обычное поле client["qualification"]["diagnosis_date"].
    # date.today() здесь — временная заглушка только для демонстрационного
    # прогона этого скрипта, не для реального использования.
    diagnosis_date = client["qualification"].get("diagnosis_date")
    if not diagnosis_date:
        from datetime import date
        diagnosis_date = date.today().strftime("%d.%m.%Y")

    info_table = Table([
        ["Компания:", company],
        ["Клиент:", name],
        ["Дата диагностики:", diagnosis_date],
        ["Консультант:", "Игорь Баландин"],
        ["Email консультанта:", "ibalandin@fenix-lms.ru"],
    ], colWidths=[45 * mm, 100 * mm])
    info_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#9FB4BC")),
        ("TEXTCOLOR", (1, 0), (1, -1), WHITE),
        ("FONTNAME", (0, 0), (-1, -1), "Roboto"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(info_table)
    story.append(PageBreak())

    # ---------- Стр. 3: Содержание ----------
    story.append(SectionBanner("СОДЕРЖАНИЕ"))
    story.append(Spacer(1, 8 * mm))
    toc_items = [
        "1  Обзор компании", "2  Приоритетные сферы", "3  Коэффициент Строитель-Протектор",
        "4  Модальность", "5  Три роли лидера", "6  Стили управления",
        "7  Классические вызовы", "8  Уровень зрелости бизнеса", "9  Непреложные правила",
        "10  Ключевые системные элементы (справка)", "11  Ключевые системные элементы (рекомендация)",
        "12  Дальнейшие шаги",
    ]
    for item in toc_items:
        story.append(para(item))
    story.append(PageBreak())

    # ---------- Раздел 1 ----------
    story.extend(section_header(1, "Обзор компании"))
    fte = stage["fte"]
    full_stage = next(st for st in data["stages"]["stages"] if st["id"] == stage_id)
    TABLE_FONT = 13  # минимум Body(10.5) + 2
    fte_lo, fte_hi = full_stage["fte_range"]
    info_table = Table([
        ["Кол-во сотрудников (текущее)", str(fte)],
        ["Кол-во сотрудников (год назад)", str(client["qualification"]["employeesYearAgo"])],
        ["Кол-во Менеджеров", str(client["qualification"]["managers_actual"])],
        ["Кол-во Высших руководителей", str(client["qualification"]["leaders_actual"])],
        ["Сколько лет существует бизнес", str(client["qualification"]["years_in_business"])],
    ], colWidths=[100 * mm, 40 * mm], style=TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), TABLE_FONT),
        ("TEXTCOLOR", (0, 0), (0, -1), TEAL), ("FONTNAME", (0, 0), (0, -1), "Roboto-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    info_table.hAlign = "LEFT"
    story.append(KeepTogether([para("Предоставленная информация", "H3"), info_table]))
    story.append(Spacer(1, 4 * mm))
    stage_table = Table([
        ["Текущая Стадия", f'{stage_id} ({fte_lo} – {fte_hi} сотрудников)'],
        ["Название Стадии", stage["stage_name"]],
        ["Процент прохождения Стадии", f'{stage["percent_through_stage"]}%'],
        ["Целевое кол-во Менеджеров", format_target_range(full_stage["managers_range"])],
        ["Целевое кол-во Руководителей", format_target_range(full_stage["leaders_range"])],
        ["Зона входа", format_zone_entry(full_stage)],
        ["Текущая зона", format_current_zone(stage)],
        ["Зона выхода", format_zone_exit(full_stage)],
    ], colWidths=[100 * mm, 60 * mm], style=TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), TABLE_FONT),
        ("TEXTCOLOR", (0, 0), (0, -1), TEAL), ("FONTNAME", (0, 0), (0, -1), "Roboto-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    stage_table.hAlign = "LEFT"
    story.append(KeepTogether([para("Данные по Стадии роста", "H3"), stage_table]))
    story.append(Spacer(1, 6 * mm))
    story.append(KeepTogether([para("РЕЗЮМЕ", "H3"), para(data["stage_level_report_texts"][s]["РЕЗЮМЕ"])]))
    story.append(PageBreak())
    critical_gaps = find_critical_gaps(result, data)
    vyvody = rtg.render_section1_vyvody(result, company, data["stages"], has_critical_gaps=bool(critical_gaps))
    append_header_with_paragraphs(story, "ВЫВОДЫ", vyvody)
    story.append(PageBreak())

    # ---------- Разделы 2-6 (Поток А) ----------
    cd = client["flow_a_dimensions"]
    targets = data["rules_of_growth_targets"]

    RED = colors.HexColor("#C0392B")
    ROW_ODD = colors.HexColor("#CCD2D8")   # нечётные строки с данными (1-я, 3-я...)
    ROW_EVEN = colors.HexColor("#DCE0E4")  # чётные строки с данными (2-я, 4-я...)

    def _row_band_cmds(n_data_rows, first_row_index=1):
        cmds = []
        for i in range(n_data_rows):
            row_idx = first_row_index + i
            color = ROW_ODD if (i + 1) % 2 == 1 else ROW_EVEN
            cmds.append(("BACKGROUND", (0, row_idx), (-1, row_idx), color))
        return cmds

    def cmp_table(rows):
        t = Table(
            [["", "Текущее значение", "Целевое значение"]] + rows,
            colWidths=[55 * mm, 55 * mm, 55 * mm],
            style=TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), TABLE_FONT),
                ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"), ("TEXTCOLOR", (0, 0), (-1, 0), TEAL),
                ("FONTNAME", (0, 1), (0, -1), "Roboto-Bold"), ("TEXTCOLOR", (0, 1), (0, -1), TEAL),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LINEBELOW", (0, 0), (-1, 0), 0.6, TEAL),
            ]),
        )
        t.hAlign = "LEFT"
        return t

    # высота ячейки — измеряется по двухстрочному тексту «Рядовые сотрудники»
    # (самому длинному варианту переноса среди всех таблиц Разделов 2-5) и
    # затем применяется как единая высота строки во всех этих таблицах
    _measure_style = ParagraphStyle("TableCellMeasure", fontName="Roboto-Bold", fontSize=TABLE_FONT,
                                     leading=TABLE_FONT + 3, alignment=TA_CENTER, textColor=NAVY)
    _measure_style_red = ParagraphStyle("TableCellMeasureRed", parent=_measure_style, textColor=RED)
    _measure_p = Paragraph("Рядовые<br/>сотрудники", _measure_style)
    _measure_w, _measure_h = _measure_p.wrap(42 * mm, 1000)
    ROW_HEIGHT = _measure_h + 12  # + верх/низ паддинг 6+6pt

    def _wrapped_cell(text, red=False):
        # ВАЖНО: цвет текста нужно задавать здесь, внутри стиля самого
        # Paragraph, а НЕ через TEXTCOLOR в TableStyle — для ячеек-Paragraph
        # (в отличие от обычных строк) команда TEXTCOLOR не имеет эффекта,
        # цвет полностью определяется ParagraphStyle. Это было причиной
        # бага: несовпадения в разделе "Модальность" не подсвечивались
        # красным, хотя логика сравнения была верной.
        style = _measure_style_red if red else _measure_style
        return Paragraph(text.replace(" ", "<br/>", 1) if text == "Рядовые сотрудники" else text, style)

    def two_col_value_table(rows, headers=("Текущее значение", "Целевое значение")):
        """Таблица «значение/значение» в стиле Раздела 2: тёмно-бирюзовая шапка,
        чередующиеся строки (#CCD2D8/#DCE0E4), белые границы, несовпадения слева — красным."""
        table_rows = [[headers[0], headers[1]]] + [[cur, tgt] for cur, tgt in rows]
        style_cmds = [
            ("FONTNAME", (0, 0), (-1, -1), "Roboto-Bold"), ("FONTSIZE", (0, 0), (-1, -1), TABLE_FONT),
            ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("GRID", (0, 0), (-1, -1), 1, WHITE),
        ]
        style_cmds += _row_band_cmds(len(rows))
        for i, (cur, tgt) in enumerate(rows, start=1):
            style_cmds.append(("TEXTCOLOR", (0, i), (0, i), RED if cur != tgt else NAVY))
            style_cmds.append(("TEXTCOLOR", (1, i), (1, i), NAVY))
        row_heights = [None] + [ROW_HEIGHT] * len(rows)
        t = Table(table_rows, colWidths=[70 * mm, 70 * mm], rowHeights=row_heights, style=TableStyle(style_cmds))
        t.hAlign = "LEFT"
        return t

    def modality_table(client_modality, target_modality):
        """4-колоночная таблица «Текущая/Целевая комбинация ролей»: роль —
        уровень (текущий) — роль — уровень (целевой), несовпадения красным.
        «Рядовые сотрудники» переносится на 2 строки; высота всех строк
        (включая шапку) выровнена по этой ячейке."""
        display_level = {"Руководство": "Руководство", "Менеджеры": "Менеджеры", "Сотрудники": "Рядовые сотрудники"}
        inv_current = {role: level for level, role in client_modality.items()}
        inv_target = {role: level for level, role in target_modality.items()}
        roles_order = ["Доминирующая", "Поддерживающая", "Вспомогательная"]

        table_rows = [["Текущая комбинация ролей", "", "Целевая комбинация ролей", ""]]
        for role in roles_order:
            cur_level = display_level[inv_current[role]]
            tgt_level = display_level[inv_target[role]]
            mismatch = inv_current[role] != inv_target[role]
            table_rows.append([role, _wrapped_cell(cur_level, red=mismatch), role, _wrapped_cell(tgt_level)])

        style_cmds = [
            ("FONTNAME", (0, 0), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), TABLE_FONT),
            ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
            ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("SPAN", (0, 0), (1, 0)), ("SPAN", (2, 0), (3, 0)),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 1), (0, -1), 2), ("LEFTPADDING", (1, 1), (1, -1), 10),
            ("RIGHTPADDING", (2, 1), (2, -1), 2), ("LEFTPADDING", (3, 1), (3, -1), 10),
            ("GRID", (0, 0), (-1, -1), 1, WHITE),
        ]
        style_cmds += _row_band_cmds(len(roles_order))
        row_heights = [ROW_HEIGHT] * (len(roles_order) + 1)
        t = Table(table_rows, colWidths=[44 * mm, 41 * mm, 44 * mm, 41 * mm], rowHeights=row_heights,
                  style=TableStyle(style_cmds))
        t.hAlign = "LEFT"
        return t

    def management_styles_table(client_styles, target_styles):
        """4-колоночная таблица «Естественное/Целевое сочетание стилей» —
        по образцу таблицы Модальности (Раздел 4): подпись — стиль (текущий) —
        подпись — стиль (целевой), несовпадения красным. Шрифт уменьшен до
        11.5pt, чтобы длинные названия стилей не задевали границы ячеек."""
        STYLE_FONT = 11.5
        slots_order = ["Основной", "Второстепенный", "Дополнительный"]
        table_rows = [["Естественное сочетание стилей", "", "Целевое сочетание стилей", ""]]
        for slot in slots_order:
            table_rows.append([slot, client_styles[slot], slot, target_styles[slot]])

        style_cmds = [
            ("FONTNAME", (0, 0), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), STYLE_FONT),
            ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"), ("FONTSIZE", (0, 0), (-1, 0), TABLE_FONT),
            ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("SPAN", (0, 0), (1, 0)), ("SPAN", (2, 0), (3, 0)),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 1), (0, -1), 2), ("LEFTPADDING", (1, 1), (1, -1), 8),
            ("RIGHTPADDING", (2, 1), (2, -1), 2), ("LEFTPADDING", (3, 1), (3, -1), 8),
            ("FONTNAME", (1, 1), (1, -1), "Roboto-Bold"), ("FONTNAME", (3, 1), (3, -1), "Roboto-Bold"),
            ("TEXTCOLOR", (3, 1), (3, -1), NAVY),
            ("GRID", (0, 0), (-1, -1), 1, WHITE),
        ]
        style_cmds += _row_band_cmds(len(slots_order))
        for i, slot in enumerate(slots_order, start=1):
            mismatch = client_styles[slot] != target_styles[slot]
            style_cmds.append(("TEXTCOLOR", (1, i), (1, i), RED if mismatch else NAVY))
        row_heights = [ROW_HEIGHT] * (len(slots_order) + 1)
        t = Table(table_rows, colWidths=[40 * mm, 45 * mm, 40 * mm, 45 * mm], rowHeights=row_heights,
                  style=TableStyle(style_cmds))
        t.hAlign = "LEFT"
        return t

    def labeled_value_table(rows, label_header="", headers=("Текущее значение", "Целевое значение")):
        """Таблица «подпись / текущее / целевое» (Разделы 5-6): та же палитра и
        белые границы, что и в остальных таблицах Разделов 2-6."""
        table_rows = [[label_header, headers[0], headers[1]]] + rows
        style_cmds = [
            ("FONTNAME", (0, 0), (-1, -1), "Roboto-Bold"), ("FONTSIZE", (0, 0), (-1, -1), TABLE_FONT),
            ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("GRID", (0, 0), (-1, -1), 1, WHITE),
        ]
        style_cmds += _row_band_cmds(len(rows))
        for i, (label, cur, tgt) in enumerate(rows, start=1):
            style_cmds.append(("TEXTCOLOR", (1, i), (1, i), RED if cur != tgt else NAVY))
            style_cmds.append(("TEXTCOLOR", (2, i), (2, i), NAVY))
        row_heights = [None] + [ROW_HEIGHT] * len(rows)
        t = Table(table_rows, colWidths=[47 * mm, 47 * mm, 47 * mm], rowHeights=row_heights,
                  style=TableStyle(style_cmds))
        t.hAlign = "LEFT"
        return t

    def priority_spheres_table(client_ranked, target_ranked):
        return two_col_value_table(list(zip(client_ranked, target_ranked)),
                                    headers=("Текущая расстановка", "Целевая расстановка"))

    SECTION3_TITLE = "Коэффициент Строитель-Протектор"

    def _vyvody_with_fallback(new_result, old_text_block):
        """Новые функции возвращают None, если клиент уже на целевом
        значении — для этого случая используется старое (не переписанное
        Игорем) положительное предложение из stage_level_report_texts.json."""
        if new_result is not None:
            return new_result
        return rtg._first_two_sentences(old_text_block)[0]

    flow_a_sections = [
        ("Приоритетные сферы", 2,
         priority_spheres_table(cd["priority_spheres"], targets["priority_spheres"][s]),
         _vyvody_with_fallback(
             rtg.render_priority_spheres_vyvody(cd["priority_spheres"], targets["priority_spheres"][s],
                                                 s, data["priority_spheres_scenarios"]),
             data["stage_level_report_texts"][s]["Приоритетные сферы"]["выводы_обе_ветки"]),
         "Приоритетные сферы", "ЦЕЛЕВОЕ ЗНАЧЕНИЕ ДЛЯ ЭТОЙ СТАДИИ"),
        ("Строитель-Протектор", 3,
         two_col_value_table([(cd["builder_protector_ratio"], targets["builder_protector_ratio"][s])]),
         _vyvody_with_fallback(
             rtg.render_builder_protector_vyvody(cd["builder_protector_ratio"], targets["builder_protector_ratio"][s],
                                                  s, data["builder_protector_scenarios"]),
             data["stage_level_report_texts"][s]["Строитель-Протектор"]["выводы_обе_ветки"]),
         SECTION3_TITLE, "ЦЕЛЕВОЕ ЗНАЧЕНИЕ ДЛЯ ЭТОЙ СТАДИИ"),
        ("Модальность", 4,
         modality_table(cd["modality"], targets["modality"][s]),
         _vyvody_with_fallback(
             rtg.render_modality_vyvody(cd["modality"], targets["modality"][s],
                                         s, data["modality_scenarios"]),
             data["stage_level_report_texts"][s]["Модальность"]["выводы_обе_ветки"]),
         "Модальность", "ЦЕЛЕВАЯ КОМБИНАЦИЯ ДЛЯ ЭТОЙ СТАДИИ"),
        ("Три роли лидера", 5,
         labeled_value_table([[role, f'{cd["three_leader_roles"][role]}%', f'{targets["three_leader_roles"][s][role]}%']
                               for role in ["Визионер", "Менеджер", "Специалист"]]),
         rtg.render_leader_roles_vyvody(cd["three_leader_roles"], targets["three_leader_roles"][s],
                                         data["stage_level_report_texts"][s]["Три роли лидера"]["выводы_обе_ветки"]),
         "Три роли лидера", "ЦЕЛЕВОЕ ЗНАЧЕНИЕ ДЛЯ ЭТОЙ СТАДИИ"),
    ]

    for name, num, table, vyvody_text, display_title, subtitle_label in flow_a_sections:
        story.extend(section_header(num, display_title))
        block = data["stage_level_report_texts"][s][name]
        story.append(para(rtg.FLOW_A_INTRO_TEXTS[name]))
        story.append(Spacer(1, 3 * mm))
        story.append(table)
        story.append(Spacer(1, 4 * mm))
        story.append(header_with_body(subtitle_label, block["целевое_значение_объяснение"]))
        story.append(header_with_body("ВЫВОДЫ", vyvody_text))
        story.append(PageBreak())

    # ---------- Раздел 6: Стили управления (особая структура — см. пояснение в чате) ----------
    story.extend(section_header(6, "Стили управления"))
    story.append(para(rtg.FLOW_A_INTRO_TEXTS["Стили управления"]))
    story.append(Spacer(1, 4 * mm))
    story.append(TwoLineBanner("ЕСТЕСТВЕННОЕ СОЧЕТАНИЕ СТИЛЕЙ УПРАВЛЕНИЯ", client["qualification"]["name"],
                                bg_color=TEAL))
    story.append(Spacer(1, 4 * mm))

    ranking = rtg.render_natural_styles_ranking(cd["management_styles_scores"], cd["management_styles"])
    for style, score, label in ranking:
        title = f"{style} - {score} баллов" + (f" ({label})" if label else "")
        story.append(header_with_body(title, rtg.MANAGEMENT_STYLE_DESCRIPTIONS[style], header_style="StyleTitle"))
        story.append(Spacer(1, 2 * mm))

    styles6_block = data["stage_level_report_texts"][s]["Стили управления"]
    styles6_table = management_styles_table(cd["management_styles"], targets["management_styles"][s])
    story.append(styles6_table)
    story.append(Spacer(1, 4 * mm))
    story.append(header_with_body("ЦЕЛЕВОЕ СОЧЕТАНИЕ ДЛЯ ЭТОЙ СТАДИИ", styles6_block["целевое_значение_объяснение"]))
    story.append(header_with_body("ВЫВОДЫ", rtg.render_management_styles_vyvody(
        cd["management_styles"], targets["management_styles"][s], styles6_block)))
    story.append(PageBreak())
    story.extend(section_header(7, "Классические вызовы"))
    story.append(para(
        "Классические вызовы — это 5 основных проблем, с которыми бизнес обычно "
        "сталкивается на каждой Стадии роста. Всего на Стадиях роста существует 24 "
        "уникальных Классических вызова.\nВ отличие от других измерений Классические "
        "вызовы не имеют никаких целевых или идеальных значений или сочетаний. Знание и "
        "понимание Классических вызовов помогает организации быстрее выявлять проблемы, "
        "когда они возникают, и обеспечивают общую терминологию, которая определяет "
        "главные проблемы компании.",
        "SmallBody",
    ))
    story.append(Spacer(1, 3 * mm))

    RED_C00 = colors.HexColor("#C00000")
    challenges_by_stage = data["classic_challenges_top5_by_stage"]
    sorted_challenges = sorted(client["challenge_scores"].items(), key=lambda x: -x[1])
    top5_names = {ch for ch, _ in sorted_challenges[:5]}

    challenge_rows = [["Классический вызов", "Оценка", "Стадия роста"]]
    for ch, sc in sorted_challenges:
        stage_list = challenges_by_stage.get(ch, [])
        stage_text = ", ".join(str(x) for x in stage_list)
        challenge_rows.append([ch, str(sc), stage_text])

    t = Table(challenge_rows, colWidths=[110 * mm, 25 * mm, 35 * mm], repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#F5F1E8")]),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5), ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("ALIGN", (1, 0), (2, -1), "CENTER"),
    ]
    for i, (ch, sc) in enumerate(sorted_challenges, start=1):
        if ch in top5_names:
            style_cmds.append(("FONTNAME", (0, i), (1, i), "Roboto-Bold"))
            stage_list = challenges_by_stage.get(ch, [])
            if stage_id not in stage_list:
                style_cmds.append(("TEXTCOLOR", (2, i), (2, i), RED_C00))
                style_cmds.append(("FONTNAME", (2, i), (2, i), "Roboto-Bold"))
    t.setStyle(TableStyle(style_cmds))
    story.append(t)
    story.append(PageBreak())

    # ---------- Раздел 8: Уровень зрелости бизнеса ----------
    story.extend(section_header(8, "Уровень зрелости бизнеса"))
    time_text = f'{client["qualification"]["timeYears"]} год(а), {client["qualification"]["timeMonths"]} мес.'
    maturity_text = rtg.render_section8_maturity(result, stage_id, time_at_stage_text=time_text)
    for block in maturity_text.split("\n\n"):
        story.append(para(block))
    story.append(PageBreak())

    # ---------- Раздел 9: Непреложные правила ----------
    story.extend(section_header(9, "Непреложные правила"))
    story.append(para(
        "Непреложные правила – это очень мощное измерение методологии Стадии роста. "
        "Правила распределены по 6 основным областям бизнеса, потому что бизнес должен "
        "расти пропорционально в этих направлениях. Бизнес устойчиво растёт, если он "
        "выполняет Непреложные правила, относящиеся к его текущей Стадии роста. Правила "
        "имеют накопительный кумулятивный эффект – некоторые из них действуют ещё с более "
        "ранних Стадий.\n"
        "Для поддержания своего роста организация должна минимум на 80% соответствовать "
        "каждому Правилу на своей текущей Стадии. Для многих организаций, которые "
        "остановились в росте или даже откатились назад в своём развитии, причина их "
        "неудач может быть напрямую связана с неспособностью их лидеров обеспечить "
        "соблюдение Непреложных правил."
    ))
    story.append(Spacer(1, 4 * mm))

    RULE_TABLE_WIDTHS = [122 * mm, 23 * mm, 25 * mm]  # сумма = ширине плашки заголовка Раздела
    pct_header_style = ParagraphStyle("PctHeader", fontName="Roboto-Bold", fontSize=8.5,
                                       leading=10.5, textColor=WHITE, alignment=TA_CENTER)
    rule_cell_style = ParagraphStyle("RuleCell", fontName="Roboto", fontSize=8.5, leading=11,
                                      textColor=INK)
    R9_ROW_ODD = colors.HexColor("#CCD2D8")
    R9_ROW_EVEN = colors.HexColor("#E7EAED")

    rules_struct = data["immutable_rules"][s]
    # для блока ВЫВОДЫ: собираем % по Стадии появления правила и по области, кумулятивно
    pct_by_origin_stage = {}
    pct_by_area = {}
    for area, rules in rules_struct.items():
        pct_by_area[area] = []
        for i, rule in enumerate(rules):
            pct = client["immutable_rules_pct"][area][i]
            pct_by_origin_stage.setdefault(rule["стадия_появления"], []).append(pct)
            pct_by_area[area].append(pct)

    for area, rules in rules_struct.items():
        rows = [["Правило", "Стадия", Paragraph("%<br/>выполнения", pct_header_style)]]
        cell_colors = []
        for i, rule in enumerate(rules):
            pct = client["immutable_rules_pct"][area][i]
            rows.append([Paragraph(rule["текст"], rule_cell_style), str(rule["стадия_появления"]), f'{pct}%'])
            cell_colors.append(TRAFFIC[pct])
        rt = Table(rows, colWidths=RULE_TABLE_WIDTHS)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
            ("FONTNAME", (0, 1), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (0, 0), "LEFT"),
            ("ALIGN", (1, 0), (2, 0), "CENTER"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("ALIGN", (1, 0), (2, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 1, WHITE),
        ]
        for i in range(len(rules)):
            row_idx = i + 1
            band = R9_ROW_ODD if row_idx % 2 == 1 else R9_ROW_EVEN
            style_cmds.append(("BACKGROUND", (0, row_idx), (1, row_idx), band))
        for i, c in enumerate(cell_colors):
            style_cmds.append(("BACKGROUND", (2, i + 1), (2, i + 1), c))
            style_cmds.append(("TEXTCOLOR", (2, i + 1), (2, i + 1), WHITE))
        rt.setStyle(TableStyle(style_cmds))
        story.append(header_with_body(area, rt))
        story.append(Spacer(1, 3 * mm))

    if critical_gaps:
        story.append(Spacer(1, 4 * mm))
        story.append(header_with_body(
            "КРИТИЧЕСКИЕ УПУЩЕНИЯ",
            "Следующие Правила закрепились в Вашем бизнесе с предыдущих Стадий, но так и "
            "остаются практически невыполненными:"
        ))
        for g in critical_gaps:
            story.append(para(f'• {g["правило"]} — действует со Стадии {g["стадия_появления"]}, выполнено на {g["факт_%"]}%'))

    # ---------- Раздел 9: ВЫВОДЫ (% выполнения по Стадиям / по областям) ----------
    def traffic_color_for_avg(pct):
        if pct >= 80:
            return TRAFFIC[80]
        if pct >= 60:
            return TRAFFIC[60]
        return TRAFFIC[20]

    story.append(Spacer(1, 4 * mm))

    stage_rows = [["Стадия", "% выполнения"]]
    stage_style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("FONTNAME", (0, 1), (0, -1), "Roboto"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 1, WHITE),
    ]
    all_pcts = []
    for st_num in range(1, 8):
        vals = pct_by_origin_stage.get(st_num)
        row_idx = st_num
        band = R9_ROW_ODD if row_idx % 2 == 1 else R9_ROW_EVEN
        stage_style_cmds.append(("BACKGROUND", (0, row_idx), (0, row_idx), band))
        if vals:
            avg = round(sum(vals) / len(vals))
            all_pcts.extend(vals)
            stage_rows.append([f"Стадия {st_num}", f"{avg}%"])
            stage_style_cmds.append(("BACKGROUND", (1, row_idx), (1, row_idx), traffic_color_for_avg(avg)))
            stage_style_cmds.append(("TEXTCOLOR", (1, row_idx), (1, row_idx), WHITE))
        else:
            stage_rows.append([f"Стадия {st_num}", ""])
            stage_style_cmds.append(("BACKGROUND", (1, row_idx), (1, row_idx), band))
    total_avg = round(sum(all_pcts) / len(all_pcts)) if all_pcts else 0
    stage_rows.append(["ВСЕГО", f"{total_avg}%"])
    stage_style_cmds.append(("FONTNAME", (0, -1), (-1, -1), "Roboto-Bold"))
    stage_style_cmds.append(("BACKGROUND", (0, -1), (0, -1), colors.HexColor("#E7EAED")))
    stage_style_cmds.append(("BACKGROUND", (1, -1), (1, -1), traffic_color_for_avg(total_avg)))
    stage_style_cmds.append(("TEXTCOLOR", (1, -1), (1, -1), WHITE))
    stage_table = Table(stage_rows, colWidths=[40 * mm, 25 * mm], style=TableStyle(stage_style_cmds))
    stage_table.hAlign = "LEFT"

    area_rows = [["Область бизнеса", "% выполнения"]]
    area_style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("FONTNAME", (0, 1), (0, -1), "Roboto"),
        ("ALIGN", (1, 0), (1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 1, WHITE),
    ]
    for i, (area, vals) in enumerate(pct_by_area.items(), start=1):
        band = R9_ROW_ODD if i % 2 == 1 else R9_ROW_EVEN
        area_style_cmds.append(("BACKGROUND", (0, i), (0, i), band))
        avg = round(sum(vals) / len(vals)) if vals else 0
        area_rows.append([area, f"{avg}%"])
        area_style_cmds.append(("BACKGROUND", (1, i), (1, i), traffic_color_for_avg(avg)))
        area_style_cmds.append(("TEXTCOLOR", (1, i), (1, i), WHITE))
    area_table = Table(area_rows, colWidths=[70 * mm, 30 * mm], style=TableStyle(area_style_cmds))
    area_table.hAlign = "LEFT"

    two_col = Table([[stage_table, "", area_table]], colWidths=[65 * mm, 4 * mm, 100 * mm])
    two_col.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(header_with_body("ВЫВОДЫ", two_col))
    story.append(PageBreak())

    # ---------- Раздел 10: КСЭ справка ----------
    story.extend(section_header(10, "Ключевые системные элементы (справка)"))
    story.append(para(
        "Ключевые системные элементы бизнеса – это принципы, методы и инструменты, "
        "которые образуют ключевые структуры в организации, повышая устойчивость и "
        "динамический порядок в бизнесе. Значение каждого Системного элемента для "
        "эффективности его воздействия на бизнес зависит от Стадии роста."
    ))
    story.append(Spacer(1, 4 * mm))

    MARKER_FONT_SIZE = 19  # x2 от прежних 9.5pt
    MARKER_INFO = {
        "✓": {"symbol": "☑", "color": "#7F7F7F", "font": "DejaVuSans",
              "label": "Фундамент - элемент должен быть полностью внедрён и развёрнут в организации"},
        "●●●●●": {"symbol": "●●●●●", "color": "#1A7573", "font": "Roboto",
                   "label": "Ядро - элемент имеет высший приоритет к внедрению на данной Стадии"},
        "●●●": {"symbol": "●●●", "color": "#E9BD41", "font": "Roboto",
                 "label": "Надстройка - элемент может быть внедрён на данной Стадии в качестве 2-го приоритета"},
        "●": {"symbol": "●", "color": "#9A1C1F", "font": "Roboto",
               "label": "Элемент не является необходимым для внедрения на данной Стадии"},
    }

    def marker_paragraph(marker, size=MARKER_FONT_SIZE):
        info = MARKER_INFO[marker]
        style = ParagraphStyle(f"Marker_{marker}_{size}", fontName=info["font"], fontSize=size,
                                leading=size * 1.15, textColor=colors.HexColor(info["color"]),
                                alignment=TA_CENTER)
        return Paragraph(info["symbol"], style)

    kse_priority = data["kse_priority_by_stage"]["kse_priority_by_stage"]

    kse_table_rows = [["Ключевой системный элемент", "Приоритет на Стадии " + s]]
    kse_table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), TEAL), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
        ("FONTNAME", (0, 1), (0, -1), "Roboto"), ("FONTSIZE", (0, 0), (0, -1), 9.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 1, WHITE),
    ]
    for i, kse in enumerate(data["mapping_kse"]["kse_list"], start=1):
        marker = kse_priority[kse][s]["маркер"]
        band = colors.HexColor("#CCD2D8") if i % 2 == 1 else colors.HexColor("#E7EAED")
        kse_table_rows.append([kse, marker_paragraph(marker)])
        kse_table_style.append(("BACKGROUND", (0, i), (0, i), band))
        kse_table_style.append(("BACKGROUND", (1, i), (1, i), band))
    kse_table = Table(kse_table_rows, colWidths=[125 * mm, 45 * mm], style=TableStyle(kse_table_style))
    kse_table.hAlign = "LEFT"
    story.append(kse_table)
    story.append(Spacer(1, 3 * mm))

    legend_label_style = ParagraphStyle("LegendLabel", parent=styles["Body"], fontSize=7, leading=9.1, spaceAfter=0)
    marker_col_mm = 25
    marker_box_width = marker_col_mm * mm - 4  # минус паддинг колонки (2+2pt)
    marker_box_height = MARKER_FONT_SIZE * 1.15 - 1
    legend_rows = []
    for info in MARKER_INFO.values():
        marker_flowable = MarkerGlyph(
            info["symbol"], info["font"], MARKER_FONT_SIZE, info["color"],
            marker_box_width, marker_box_height, valign="MIDDLE",
        )
        legend_rows.append([marker_flowable, Paragraph(info["label"], legend_label_style)])
    legend_table = Table(legend_rows, colWidths=[marker_col_mm * mm, 145 * mm])
    legend_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (0, -1), "MIDDLE"),
        ("VALIGN", (1, 0), (1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (1, 0), (1, -1), "LEFT"),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING", (0, 0), (0, -1), 2), ("RIGHTPADDING", (0, 0), (0, -1), 2),
        ("LEFTPADDING", (1, 0), (1, -1), 6), ("RIGHTPADDING", (1, 0), (1, -1), 2),
    ]))
    legend_table.hAlign = "LEFT"
    story.append(legend_table)
    story.append(Spacer(1, 5 * mm))

    for kse in data["mapping_kse"]["kse_list"]:
        desc = data["kse_descriptions"][kse]
        story.append(KeepTogether([para(kse, "H3"), para(desc["описание"], "Small")]))
    story.append(PageBreak())

    # ---------- Раздел 11 ----------
    story.extend(section_header(11, "Ключевые системные элементы (рекомендация)"))
    story.append(para(
        "Ниже в приоритетном порядке приведены Ключевые системные элементы, которые "
        "отсутствуют или недостаточно внедрены в экосистеме Вашей организации. "
        "Внедрение этих Элементов необходимо для создания мощного волнового эффекта по "
        "всему бизнесу и его вывода на путь устойчивого роста."
    ))
    story.append(Spacer(1, 4 * mm))

    # сводная таблица: КСЭ, вошедшие в каждый уровень приоритетности
    TIER_TITLES = ["ФУНДАМЕНТ", "ЯДРО", "НАДСТРОЙКА"]
    TIER_KEYS = ["фундамент", "ядро", "надстройка"]
    TIER_SYMBOLS = [
        ("☑", "DejaVuSans", "#7F7F7F"),
        ("●●●●●", "Roboto", "#1A7573"),
        ("●●●", "Roboto", "#E9BD41"),
    ]
    grouped_kse = {"фундамент": [], "ядро": [], "надстройка": []}
    for row in result["приоритизация_ксэ"]:
        tier = rtg._normalize_tier(row["ярус"])
        if tier in grouped_kse:
            grouped_kse[tier].append(row["kse"])

    header_row = []
    for title, (symbol, font_name, color) in zip(TIER_TITLES, TIER_SYMBOLS):
        header_row.append(Paragraph(
            f'{title}<br/><font face="{font_name}" color="{color}">{symbol}</font>',
            ParagraphStyle("Tier11Header", fontName="Roboto-Bold", fontSize=11, leading=15,
                            textColor=WHITE, alignment=TA_CENTER),
        ))
    max_rows = max(len(grouped_kse[k]) for k in TIER_KEYS)
    summary_rows = [header_row]
    for i in range(max_rows):
        summary_rows.append([
            grouped_kse[k][i] if i < len(grouped_kse[k]) else "" for k in TIER_KEYS
        ])
    summary_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#C6AD9E")),
        ("FONTNAME", (0, 1), (-1, -1), "Roboto"), ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 1, WHITE),
    ]
    for i in range(1, max_rows + 1):
        band = colors.HexColor("#CCD2D8") if i % 2 == 1 else colors.HexColor("#E7EAED")
        summary_style.append(("BACKGROUND", (0, i), (-1, i), band))
    summary_table = Table(summary_rows, colWidths=[56.7 * mm] * 3, style=TableStyle(summary_style))
    summary_table.hAlign = "LEFT"
    story.append(summary_table)
    story.append(Spacer(1, 6 * mm))

    TIER_BANNER_COLORS = {
        "ФУНДАМЕНТ": colors.HexColor("#7F7F7F"),
        "ЯДРО": colors.HexColor("#1A7573"),
        "НАДСТРОЙКА": colors.HexColor("#E9BD41"),
    }
    kse_text = rtg.render_section11_kse_list(result, client["challenge_scores"], data)
    blocks = [b.strip() for b in kse_text.split("\n\n") if b.strip()]
    i = 0
    while i < len(blocks):
        block = blocks[i]
        if block.startswith("## "):
            title = block[3:].strip()
            story.append(SectionBanner(title, height=13 * mm - 4, bg_color=TIER_BANNER_COLORS.get(title)))
            story.append(Spacer(1, 4 * mm))
        elif block.startswith("### "):
            # Заголовок КСЭ группируется со следующим блоком (обоснование
            # "Данные диагностики показывают...") — без этого заголовок мог
            # остаться в одиночестве внизу страницы (обнаружено 05.08.2026).
            kse_title = block[4:]
            if i + 1 < len(blocks) and not blocks[i + 1].startswith(("## ", "### ")):
                story.append(header_with_body(kse_title, blocks[i + 1]))
                i += 1  # следующий блок уже использован
            else:
                story.append(para(kse_title, "H3"))
        elif block.startswith("→ Закрывает"):
            m = re.search(r'«(.+?)»', block)
            program_name = m.group(1) if m else ""
            story.append(para(f'Адресуется Консалтинговой программой «{program_name}»'))
        else:
            block = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', block)
            story.append(para(block, "Body"))
        i += 1
    story.append(PageBreak())

    # ---------- Раздел 12 ----------
    story.extend(section_header(12, "Дальнейшие шаги"))
    for block in rtg.render_section12_1_transition(bool(critical_gaps)).split("\n\n"):
        block = block.strip()
        if block:
            story.append(para(block, "Body"))

    grouped_s12 = {"фундамент": [], "ядро": [], "надстройка": []}
    for row in result["приоритизация_ксэ"]:
        tier = rtg._normalize_tier(row["ярус"])
        if tier in grouped_s12:
            grouped_s12[tier].append(row["kse"])

    programs_data = data["consulting_programs"]["individual_programs"]
    for tier_key in ("фундамент", "ядро", "надстройка"):
        tier_rows = grouped_s12[tier_key]
        if not tier_rows:
            continue
        title, tier_intro = rtg.TIER_HEADERS[tier_key]
        tier_intro = tier_intro.rstrip(".") + ":"

        table_rows = [["Ключевой системный элемент к внедрению", "Консалтинговая программа"]]
        for kse in tier_rows:
            table_rows.append([kse, programs_data[kse]["название_программы"]])
        tier_table_style = [
            ("BACKGROUND", (0, 0), (-1, 0), TIER_BANNER_COLORS.get(title, NAVY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), WHITE), ("FONTNAME", (0, 0), (-1, 0), "Roboto-Bold"),
            ("FONTNAME", (0, 1), (-1, -1), "Roboto"), ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("GRID", (0, 0), (-1, -1), 1, WHITE),
        ]
        for i in range(1, len(tier_rows) + 1):
            band = colors.HexColor("#CCD2D8") if i % 2 == 1 else colors.HexColor("#E7EAED")
            tier_table_style.append(("BACKGROUND", (0, i), (-1, i), band))
        tier_table = Table(table_rows, colWidths=[85 * mm, 85 * mm], style=TableStyle(tier_table_style))
        tier_table.hAlign = "LEFT"
        story.append(KeepTogether([
            SectionBanner(title, height=13 * mm - 4, bg_color=TIER_BANNER_COLORS.get(title)),
            Spacer(1, 4 * mm), para(tier_intro), Spacer(1, 3 * mm), tier_table,
        ]))
        story.append(Spacer(1, 5 * mm))

    # блок "Альтернативный вариант" — только если клиенту подходит бандл "Возрождение малого бизнеса"
    if rtg.render_section12_3_bundle_fork(result, data) is not None:
        story.append(Spacer(1, 4 * mm))

        bundle = data["consulting_programs"]["bundle_programs"]["Возрождение малого бизнеса"]
        bundle_kse = set(bundle["ксэ_покрывает"])
        plan_kse = {r["kse"] for r in result["приоритизация_ксэ"] if r["ярус"] != "низкий_приоритет"}
        matched_programs = sorted(
            data["consulting_programs"]["individual_programs"][k]["название_программы"]
            for k in (plan_kse & bundle_kse)
        )

        story.append(KeepTogether([
            SectionBanner("Альтернативный вариант", height=13 * mm - 4, bg_color=colors.HexColor("#703C65")),
            Spacer(1, 4 * mm), para(ALT_VARIANT_TEXT_INTRO),
        ]))
        programs_list_html = "<br/>".join(f'• {p}' for p in matched_programs)
        story.append(para(programs_list_html))
        story.append(Spacer(1, 3 * mm))

        PLATE_GRAY = colors.HexColor("#BFBFBF")
        plate_style = ParagraphStyle("AltPlateBody", parent=styles["Body"], spaceAfter=0)

        def plate(html_content):
            t = Table([[Paragraph(html_content, plate_style)]], colWidths=[169.8 * mm],
                      style=TableStyle([
                          ("BACKGROUND", (0, 0), (-1, -1), PLATE_GRAY),
                          ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                          ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                      ]))
            t.hAlign = "LEFT"
            return t

        story.append(plate(ALT_VARIANT_PATH_A))
        story.append(Spacer(1, 4 * mm))
        story.append(plate(ALT_VARIANT_PATH_B))
        story.append(Spacer(1, 4 * mm))

        story.append(para(ALT_VARIANT_BRIDGE))
        for q in ALT_VARIANT_QUESTIONS:
            story.append(para(f'<font face="Roboto-Black">{q}</font>'))

    story.append(SetPageFlag(page_state, "closing", True))
    story.append(PageBreak())

    # ---------- Закрывающая страница ----------
    story.append(Spacer(1, 60 * mm))
    closing_logo = Image(str(ASSETS / "logos" / "Vertical_real_transparent.png"),
                          width=60 * mm, height=60 * mm * 991 / 1483)
    closing_logo.hAlign = "CENTER"
    story.append(closing_logo)
    story.append(Spacer(1, 10 * mm))
    story.append(Paragraph(
        "Наша миссия — преумножать число устойчивых и растущих бизнесов.",
        ParagraphStyle("Mission", fontName="Roboto-Bold", fontSize=21, leading=26,
                        textColor=WHITE, alignment=TA_CENTER)
    ))
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("WWW.FENIX-LAB.RU", ParagraphStyle(
        "MissionUrl", fontName="Roboto-Bold", fontSize=11, textColor=ORANGE, alignment=TA_CENTER)))

    return story, page_state


def main():
    data = sa.load_data()
    with open(BASE / "data" / "stage_level_report_texts.json", encoding="utf-8") as f:
        data["stage_level_report_texts"] = json.load(f)
    with open(BASE / "data" / "consulting_programs.json", encoding="utf-8") as f:
        data["consulting_programs"] = json.load(f)
    with open(BASE / "data" / "priority_spheres_scenarios.json", encoding="utf-8") as f:
        data["priority_spheres_scenarios"] = json.load(f)
    with open(BASE / "data" / "builder_protector_scenarios.json", encoding="utf-8") as f:
        data["builder_protector_scenarios"] = json.load(f)
    with open(BASE / "data" / "modality_scenarios.json", encoding="utf-8") as f:
        data["modality_scenarios"] = json.load(f)

    client = build_demo_client(data)
    result = sa.diagnose(client, data)

    story, page_state = build_story(client, result, data)

    out_path = str(BASE / "Пример_отчёта_ФЕНИКС.pdf")
    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
        title="Полная оценка состояния бизнеса",
    )
    meta = {
        "report_number": client["qualification"]["report_number"],
        "diagnosis_date": client["qualification"]["diagnosis_date"],
        "company": client["qualification"]["company"],
        "name": client["qualification"]["name"],
    }
    # Обложка (стр. 1) рисуется полностью в draw_cover(). Со стр. 2 — футер
    # с логотипом, номером отчёта, датой и названием компании на каждой странице.
    doc.build(
        story,
        onFirstPage=partial(draw_cover, meta=meta),
        onLaterPages=partial(draw_footer, meta=meta, page_state=page_state),
    )
    print(f"PDF собран: {out_path}")


if __name__ == "__main__":
    main()
