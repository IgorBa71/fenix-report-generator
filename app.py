"""
app.py — веб-сервис на Render.com, оборачивающий scoring_algorithm.py и
pdf_report_builder.py в HTTP API, чтобы Make.com мог их вызывать (Make.com
исполняет только JavaScript, не Python — поэтому нужен этот отдельный слой).

Эндпоинт:
    POST /generate-report
    Тело запроса (JSON) — сырые ответы клиента из Опросника, в том же
    формате, в каком их хранит state.answers в Опроснике (см. Опросник_v1.html).

    Ответ (JSON):
        {
          "ok": true,
          "report_number": "000351",
          "pdf_base64": "...",              # PDF-файл, закодированный в base64
          "diagnose_result": {...}          # результат scoring_algorithm.diagnose()
                                             # — пригодится для Google-таблицы
        }

Все формулы агрегации ответов Опросника (Приоритетные сферы, Строитель-Протектор,
Модальность, Стили управления, соответствие Раздела 8 → 11 КСЭ) подтверждены
Игорем по оригинальным файлам-ключам проекта — см.
04_Связка_Make_com/Формулы_агрегации_ответов_Опросника.md в пакете передачи.
"""

import base64
import io
import json
import requests
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate
from functools import partial

import scoring_algorithm as sa
import pdf_report_builder as prb

app = Flask(__name__)

BASE = Path(__file__).parent


@app.after_request
def add_cors_headers(response):
    """Разрешаем запросы с любого источника (Опросник встроен в Tilda,
    домен там разный в зависимости от подключённого домена клиента,
    поэтому проще разрешить всем, чем поддерживать список доменов)."""
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return response


@app.route("/generate-report", methods=["OPTIONS"])
def generate_report_preflight():
    """Браузер сначала посылает OPTIONS-запрос (preflight) перед POST —
    нужно ответить на него отдельно, до основного маршрута."""
    return ("", 204)


# ---------------------------------------------------------------------------
# Перевод сырых ответов Опросника в формат, который ожидает scoring_algorithm
# ---------------------------------------------------------------------------

BUILDER_PROTECTOR_MAP = {
    "Абсолютно уверенное": "4:1",
    "Очень уверенное": "3:1",
    "Уверенное": "2:1",
    "Скорее уверенное": "3:2",
    "В равной степени уверенное и осмотрительное": "1:1",
    "Скорее осмотрительное": "2:3",
    "Осмотрительное": "1:2",
    "Очень осмотрительное": "1:3",
}

MANAGEMENT_STYLE_LETTER_TO_NAME = {
    0: "Авторитетный", 1: "Коучинговый", 2: "Товарищеский",
    3: "Демократический", 4: "Эталонный", 5: "Директивный",
}


def convert_priority_spheres(section1_answers, statements_data):
    """
    section1_answers: {текст_утверждения: ранг 1-3}
    statements_data: STATEMENTS.priority_spheres_map — {текст: сфера}
    Возвращает: ["Прибыль", "Люди", "Процессы"] — по убыванию суммы баллов.
    Подтверждено по файлу-ключу Приоритетные_сферы__ключ_к_Оценочному_опросу.xlsx.
    """
    sums = {"Люди": 0, "Прибыль": 0, "Процессы": 0}
    for text, sphere in statements_data.items():
        rank = section1_answers.get(text, 0)
        sums[sphere] += rank
    return sorted(sums.keys(), key=lambda s: sums[s], reverse=True)


def convert_builder_protector(section3_answer):
    return BUILDER_PROTECTOR_MAP.get(section3_answer, "1:1")


def convert_modality(section4_answers, modality_questions):
    """
    section4_answers: [значение_уровня, ...] — 9 элементов по порядку вопросов
    Группа 1 (0-2) -> Доминирующая, Группа 2 (3-5) -> Поддерживающая,
    Группа 3 (6-8) -> Вспомогательная. Внутри группы — мода (частый выбор).
    """
    roles = ["Доминирующая", "Поддерживающая", "Вспомогательная"]
    result = {}
    for group_idx in range(3):
        group_answers = section4_answers[group_idx * 3:(group_idx + 1) * 3]
        counts = {}
        for level in group_answers:
            counts[level] = counts.get(level, 0) + 1
        winner = max(counts, key=counts.get)
        result[winner] = roles[group_idx]
    return result


def convert_management_styles(section5_answers):
    """
    section5_answers: {1: {текст: балл 1-6}, 2: {...}, 3: {...}} — 3 группы по 6.
    Буква А-Е соответствует одному и тому же стилю во всех 3 группах
    (порядок как в исходных данных STATEMENTS.management_styles_statements).
    Возвращает (styles_dict, scores_dict) — второе нужно для отображения в PDF.
    """
    totals = {name: 0 for name in MANAGEMENT_STYLE_LETTER_TO_NAME.values()}
    for group_num in [1, 2, 3]:
        group = section5_answers.get(str(group_num)) or section5_answers.get(group_num) or {}
        for idx, (text, score) in enumerate(group.items()):
            if idx < 6:
                style_name = MANAGEMENT_STYLE_LETTER_TO_NAME[idx]
                totals[style_name] += score
    ranked = sorted(totals.items(), key=lambda kv: kv[1], reverse=True)
    styles = {
        "Основной": ranked[0][0],
        "Второстепенный": ranked[1][0],
        "Дополнительный": ranked[2][0],
    }
    return styles, totals


def convert_immutable_rules_pct(section7_answers, stage_id, data):
    """
    section7_answers: {"область::индекс": процент}
    Возвращает: {область: [pct1, pct2, ...]} в порядке правил для этой Стадии.
    """
    rules_struct = data["immutable_rules"][str(stage_id)]
    result = {}
    for area, rules in rules_struct.items():
        pct_list = []
        for idx in range(len(rules)):
            key = f"{area}::{idx}"
            pct_list.append(section7_answers.get(key, 0))
        result[area] = pct_list
    return result


SECTION8_STATEMENT_TO_KSE = {
    "Модель моего бизнеса была изначально осознанно выстроена с прицелом на устойчивость и прибыльность": "Бизнес-модель",
    "Все организационные элементы моей компании эффективно согласованы в интересах дальнейшего роста бизнеса": "Критерии роста бизнеса",
    "Мой бизнес последовательно практикует двустороннее общение между всеми сотрудниками и супервайзерами": "Коучинговое управление персоналом",
    "Мой бизнес всегда держит обещания, данные клиентам и своей команде": "Ценности бренда и Базовые ценности",
    "Мой бизнес ведёт вперёд сильная управленческая команда, все члены которой разделяют общее видение бизнеса и всегда понимают друг друга": "Сильная Управленческая команда",
    "Мой бизнес успешно приоритизирует ресурсы и инициативы в рамках всей организации, а не по отдельным подразделениям": "Комплексное планирование",
    "Мой бизнес стабильно генерирует растущие доходы благодаря выстроенным четким структурам на направлении Развития
