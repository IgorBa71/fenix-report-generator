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
import gzip
import io
import json
import os
import re
import smtplib
import uuid
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate
from functools import partial

import requests

import scoring_algorithm as sa
import pdf_report_builder as prb
from prodamus_hmac import prodamus_sign, prodamus_verify

# ── Яндекс.Метрика Measurement Protocol — серверное подтверждение оплаты ──
# Нужно один раз включить опцию Measurement Protocol в счётчике 110281435
# (Метрика → счётчик → Настройки → Безопасность и использование данных →
# Measurement Protocol) и сгенерировать токен там же. Токен хранится в
# переменной окружения на Render, а не в коде — YANDEX_MP_TOKEN.
YANDEX_METRIKA_COUNTER_ID = "110281435"
YANDEX_MP_TOKEN = os.environ.get("YANDEX_MP_TOKEN", "")
YANDEX_MP_COLLECT_URL = "https://mc.yandex.ru/collect"

# Скрипт консультации для Игоря (не для клиента!) — генерируется ДОПОЛНИТЕЛЬНО
# к клиентскому Отчёту, из тех же diagnose_result/client_response. Обёрнут в
# try/except в generate_report(): ошибка здесь никогда не должна ронять
# генерацию клиентского PDF-отчёта.
from script_adapter import adapt as adapt_script_data
from consultation_script_builder import build_script_sections
from consultation_script_pdf_builder import build_pdf as build_script_pdf

# Презентация для консультации (для Игоря, показывается клиенту на экране) —
# тоже генерируется ДОПОЛНИТЕЛЬНО, из тех же данных. Раньше требовала Node.js
# рядом с Python (pptxgenjs) — переписана целиком на python-pptx 06.08.2026,
# чтобы не тащить второй рантайм на Render. Обёрнута в try/except по той же
# причине, что и Скрипт — сборка Презентации никогда не должна ронять ответ
# с уже готовым Отчётом.
from pptx_data_export import export_pptx_data
from pptx_presentation_builder import new_presentation
from pptx_slides import (
    build_slide_cover, build_slide_agenda, build_slide_your_words, build_slide_reasons,
    build_slide_growth_rules, build_slide_kse_concept, build_slide_symptoms,
    build_slide_maturity_top5, build_slide_priority_chain,
    build_slide_cost_only, build_slide_full_contrast, build_slide_how_to_reach,
    build_slide_kse_list_repeat, build_slide_what_order, build_slide_house,
    build_slide_house_continuation, build_slide_how_to_implement,
    build_slide_what_is_program, build_slide_program_model, build_slide_program_model_practice,
    build_slide_programs_foundation, build_slide_programs_core, build_slide_programs_superstructure,
    build_slide_guarantee, build_slide_loyalty_program, build_slide_bundle_fork,
    build_slide_closing, build_slide_mission,
)

# ---------------------------------------------------------------------------
# 21.08.2026: отказ от Make.com для Сценария 1 ("Мгновенное уведомление мне +
# PDF на Диск"). Раньше здесь был URL вебхука Make, который дальше вручную
# раскладывал файлы по Google Диску, писал строку в Google Таблицу и слал
# письмо Игорю через Gmail (fenix.checkup.report@gmail.com) — три
# нероссийских сервиса на один шаг. Теперь письмо Игорю с тремя вложениями
# (Отчёт, Скрипт консультации, Презентация) отправляется НАПРЯМУЮ из этого
# приложения через SMTP Яндекс.Почты для домена (ящик report@fenix-lab.ru).
# Google Sheets/Drive для Сценария 1 больше не используются вообще.
#
# Требуются переменные окружения на Timeweb:
#   SMTP_HOST     — smtp.yandex.ru
#   SMTP_PORT     — 465 (SSL)
#   SMTP_LOGIN    — report@fenix-lab.ru
#   SMTP_PASSWORD — пароль ПРИЛОЖЕНИЯ (не обычный пароль от ящика!),
#                   создаётся в Яндекс 360 → Безопасность → Пароли приложений
# ---------------------------------------------------------------------------
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.yandex.ru")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
SMTP_LOGIN = os.environ.get("SMTP_LOGIN", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")

# Куда падает мгновенное уведомление с тремя вложениями (раньше — Gmail
# fenix.checkup.report@gmail.com, теперь можно указать любой ящик Игоря,
# включая тот же report@fenix-lab.ru или личную почту).
IGOR_NOTIFICATION_EMAIL = os.environ.get("IGOR_NOTIFICATION_EMAIL", "report@fenix-lab.ru")

# Логин/пароль для служебной панели /admin (План Б: ручной поиск заказа
# клиента и переотправка файлов, если авто-цепочка Сц1/Сц2 не сработала
# из-за сбоя на стороне Timeweb, например DDoS-атаки).
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")

# Проактивные уведомления Игорю в мессенджер MAX (бот "Феникс Алерты"),
# на случай сбоев в отправке Отчётов (Сц1/Сц2) — см. send_max_alert() ниже.
MAX_BOT_TOKEN = os.environ.get("MAX_BOT_TOKEN", "")
MAX_ALERT_CHAT_ID = os.environ.get("MAX_ALERT_CHAT_ID", "")

# ---------------------------------------------------------------------------
# Интеграция оплаты Prodamus (HMAC-подписанные ссылки, без клиентского
# JS-виджета "Единое окно" — тот не гарантирует корректное формирование
# счёта, что подтвердила сама поддержка Продамус). ВАЖНО: секретный ключ
# и адрес формы нужно свериться с личным кабинетом Продамус — ниже
# указаны значения из предыдущей сессии настройки, возможно устарели.
# ---------------------------------------------------------------------------
PRODAMUS_FORM_URL = "https://fenix-lab.payform.ru/"
# 19.08.2026: секретный ключ Продамуса перенесён из кода в переменную
# окружения (по той же схеме, что YANDEX_MP_TOKEN) — безопаснее хранить
# секреты только на хостинге, а не в открытом виде в GitHub-репозитории.
# Обязательно добавить переменную PRODAMUS_SECRET_KEY в настройках
# приложения на хостинге (Timeweb) — иначе платежи и вебхук перестанут
# работать (подпись будет пустой строкой).
PRODAMUS_SECRET_KEY = os.environ.get("PRODAMUS_SECRET_KEY", "")
# ORDERS_FILE (путь к файлу заказов) больше не используется — 19.08.2026
# заказы Prodamus перенесены в PostgreSQL, см. _load_orders/_save_orders
# ниже по файлу.

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "")
if not app.secret_key:
    # На случай, если переменную забыли задать — приложение не должно падать
    # на старте, но логин в админку тогда будет ненадёжным (сессии не переживут
    # рестарт контейнера). В переменных окружения Timeweb ОБЯЗАТЕЛЬНО задать
    # FLASK_SECRET_KEY — любую длинную случайную строку.
    print("ВНИМАНИЕ: FLASK_SECRET_KEY не задан в переменных окружения!", flush=True)
    app.secret_key = "insecure-dev-key-change-me"

BASE = Path(__file__).parent

# ---------------------------------------------------------------------------
# CORS: раньше браузер клиента стучался на Make.com (тот разрешает запросы
# с любого домена по умолчанию), теперь — напрямую на этот Flask-сервер с
# другого домена (fenix-lab.ru), поэтому без явных заголовков браузер сам
# блокирует ответ ещё до того, как код ниже успевает отработать.
# ---------------------------------------------------------------------------

ALLOWED_ORIGINS = {
    "https://fenix-lab.ru",
    "https://www.fenix-lab.ru",
}


@app.after_request
def add_cors_headers(response):
    origin = request.headers.get("Origin")
    if origin in ALLOWED_ORIGINS:
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return response


# ---------------------------------------------------------------------------
# Данные диагностики (вопросы, правила, симптомы) — раньше отдавались с
# igorba71.github.io/fenix-data/, но GitHub Pages оказался заблокирован без
# VPN у части российских пользователей (обнаружено 18.08.2026,
# net::ERR_CONNECTION_RESET на всех 5 файлах, воспроизводится и в домашней,
# и в мобильной сети, снимается VPN). Перенесено на api.fenix-lab.ru — этот
# домен уже проксируется через Cloudflare с отключённым TLS 1.3 именно ради
# обхода похожей блокировки самого Render (см. 07_Инфраструктура_и_доступы)
# и подтверждённо работает без VPN.
#
# Имена файлов слева (для истории/справки) — то, что раньше запрашивал
# Опросник по отдельности. Имена справа — реальные файлы в data/, те же
# самые, на которых работает scoring_algorithm.py — гарантирует
# идентичность данных 1-в-1 с бэкендом.
#
# 19.08.2026, попытка №1: challenge_symptoms.json (единый файл ~155 КБ
# несжатым) разбили на 3 части поменьше (по группам стадий 1-3/4-5/6-7),
# рассчитывая, что дело в объёме одного ответа. Результат неудачный:
# итоговое время загрузки Опросника выросло до 60-90 секунд, т.к. почти
# каждый из 7 последовательных запросов сначала "спотыкался" на ~10 сек
# (наш собственный таймаут) и проходил только со второй попытки — похоже,
# нестабильность именно в установке нового соединения до api.fenix-lab.ru
# (DNS/TLS-хендшейк), а не в объёме передаваемых данных. При таком раскладе
# больше запросов = кратно больше суммарного "штрафа" за нестабильность.
#
# 19.08.2026, попытка №2 (текущая): все 5 наборов данных объединены в один
# ответ по одному маршруту /data/bundle.json — теперь Опросник делает ОДИН
# запрос вместо 5-7, и "штраф" за нестабильное соединение оплачивается
# максимум один раз (плюс retry на этот один запрос), а не многократно.
# ---------------------------------------------------------------------------
DATA_DIR = BASE / "data"
_BUNDLE_FILES = {
    "stages": "stages.json",
    "immutable_rules": "immutable_rules.json",
    "statements": "statements.json",
    "challenge_symptoms": "challenge_symptoms_by_stage.json",
    "rog_targets": "rules_of_growth_targets.json",
}


@app.route("/data/bundle.json", methods=["GET"])
def serve_diagnostic_data_bundle():
    bundle = {}
    for key, real_name in _BUNDLE_FILES.items():
        path = DATA_DIR / real_name
        if not path.exists():
            return jsonify({"ok": False, "error": f"{real_name} not found on server"}), 404
        bundle[key] = json.loads(path.read_text(encoding="utf-8"))

    raw_text = json.dumps(bundle, ensure_ascii=False)
    compressed = gzip.compress(raw_text.encode("utf-8"))
    response = app.response_class(compressed, mimetype="application/json")
    response.headers["Content-Encoding"] = "gzip"
    response.headers["Content-Length"] = str(len(compressed))
    return response


@app.route("/generate-report", methods=["OPTIONS"])
def generate_report_options():
    # Браузер сначала посылает preflight-запрос OPTIONS без тела — на него
    # достаточно ответить пустым 204 с CORS-заголовками (добавляются выше
    # через after_request), настоящий POST последует отдельным запросом.
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


def normalize_phone(raw):
    """Приводит телефон к чистому виду без скобок/дефисов/пробелов
    (например, '+7(999) 999-99-99' -> '+79999999999'). Это нужно, чтобы:
    1) Google Sheets не принимал номер за формулу (даёт #ERROR! на значениях
       вида '+7(999)...', начинающихся с '+' и содержащих скобки);
    2) данные в Google Sheets/Make были аккуратными и единообразными,
       независимо от того, как именно телефон отформатирован в поле
       ввода на Тильде."""
    if not raw:
        return raw
    raw = raw.strip()
    plus = raw.startswith("+")
    digits = re.sub(r"\D", "", raw)
    return ("+" + digits) if plus else digits


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

    ВАЖНО: один и тот же уровень (Руководство/Менеджеры/Сотрудники) может
    оказаться "победителем" сразу в нескольких группах — это обычная,
    реальная ситуация в живых ответах клиентов, а не редкий крайний случай.
    Поэтому нельзя просто класть result[winner] = role: если один уровень
    выигрывает дважды, при таком наивном подходе одна из трёх ролей вообще
    пропадёт из результата (перетрётся), что ломает построение отчёта дальше
    (KeyError в pdf_report_builder.modality_table). Правильный алгоритм
    гарантирует, что итоговый словарь всегда содержит все 3 уровня и все
    3 роли — по одному разу каждый (полное однозначное соответствие).
    """
    roles = ["Доминирующая", "Поддерживающая", "Вспомогательная"]
    all_levels = ["Руководство", "Менеджеры", "Сотрудники"]

    # Считаем голоса за уровень отдельно в каждой из 3 групп.
    group_counts = []
    for group_idx in range(3):
        group_answers = section4_answers[group_idx * 3:(group_idx + 1) * 3]
        counts = {level: 0 for level in all_levels}
        for level in group_answers:
            if level in counts:
                counts[level] += 1
        group_counts.append(counts)

    result = {}
    assigned_levels = set()
    for group_idx in range(3):
        counts = group_counts[group_idx]
        # Среди ещё не занятых уровней выбираем тот, у кого больше всего
        # голосов в ЭТОЙ группе. При равенстве голосов — устойчивый порядок
        # по all_levels, чтобы результат был детерминированным.
        remaining = [lvl for lvl in all_levels if lvl not in assigned_levels]
        winner = max(remaining, key=lambda lvl: (counts[lvl], -all_levels.index(lvl)))
        result[winner] = roles[group_idx]
        assigned_levels.add(winner)
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
    "Мой бизнес стабильно генерирует растущие доходы благодаря выстроенным четким структурам на направлении Развития бизнеса": "Структура Развития бизнеса",
    "Мой бизнес отслеживает, измеряет и оценивает результаты своей деятельности": "Ключевые показатели эффективности",
    "В моём бизнесе создана устойчивая организационная структура, позволяющая организовывать работу, которая не зависит от отдельных сотрудников": "Организационная структура",
    "В моём бизнесе постоянно совершенствуются процессы, направленные на повышение качества и принимаемые всеми сотрудниками": "Базовые бизнес-процессы",
    "В моём бизнесе есть эффективная структура внутренних совещаний": "Структура рабочих совещаний",
}


def convert_section8_by_kse(section8_answers, statements_order=None, kse_list=None):
    """
    section8_answers: {текст_утверждения: балл 1-6}
    Соответствие "утверждение -> КСЭ" подтверждено Игорем по файлу-ключу
    СИСТЕМНЫЕ_ЭЛЕМЕНТЫ_БИЗНЕСА__Ключ_к_Оценочному_вопросу.xlsx — см.
    SECTION8_STATEMENT_TO_KSE выше.
    """
    result = {}
    for text, score in section8_answers.items():
        kse_name = SECTION8_STATEMENT_TO_KSE.get(text)
        if kse_name is not None:
            result[kse_name] = [score]
    return result


def build_client_response(payload, data, statements):
    """payload — то, что пришло в теле запроса (сырые ответы Опросника)."""
    q = payload["qualification"]
    management_styles, management_styles_scores = convert_management_styles(payload["section5"])

    client_dimensions = {
        "priority_spheres": convert_priority_spheres(
            payload["section1"], statements["priority_spheres_map"]),
        "builder_protector_ratio": convert_builder_protector(payload["section3"]),
        "modality": convert_modality(payload["section4"], statements["modality_questions"]),
        "management_styles": management_styles,
        "management_styles_scores": management_styles_scores,
        "three_leader_roles": {
            "Визионер": int(payload["section6"]["visionaire"]),
            "Менеджер": int(payload["section6"]["manager"]),
            "Специалист": int(payload["section6"]["specialist"]),
        },
    }

    client_response = {
        "qualification": {
            "name": q.get("name", ""),
            "company": q.get("company", ""),
            "email": q.get("email", ""),
            "phone": normalize_phone(q.get("phone", "")),
            "diagnosis_date": datetime.now().strftime("%d.%m.%Y"),
            "report_number": prb.get_next_report_number(),
            "fte_a": float(q["fte_a"]), "fte_b": float(q["fte_b"]),
            "fte_c": float(q["fte_c"]), "fte_d": float(q["fte_d"]),
            "managers_actual": int(q.get("managers", 0)),
            "leaders_actual": int(q.get("leaders", 0)),
            "employeesYearAgo": int(q.get("employeesYearAgo", 0)),
            "timeYears": int(q.get("timeYears", 0)),
            "timeMonths": int(q.get("timeMonths", 0)),
            "years_in_business": int(q.get("businessAge", 0)),
            # Эти 3 поля — прямой ввод клиента (Шаг 3 регистрации Опросника),
            # не вычисляются scoring_algorithm. Раньше не прокидывались сюда —
            # Скрипт консультации (consultation_script_builder.py) их читал
            # напрямую из payload в обход client_response. Прокидываем здесь,
            # чтобы client_response был самодостаточным для любого потребителя.
            "psychographic": q.get("psychographic", {}),
            "urgency": q.get("urgency", ""),
            "decisionMaker": q.get("decisionMaker", ""),
        },
        "flow_a_dimensions": client_dimensions,
        "challenge_scores": payload["section2"],
        "section8_likert_by_kse": None,  # заполним ниже, нужен stage/kse_list
        # Раздел 9 «Ваш взгляд на ситуацию» — тоже прямой ввод клиента (7
        # открытых вопросов), не часть scoring_algorithm. Пробрасываем как
        # есть, без изменений, для Скрипта консультации.
        "section9": payload.get("section9", {}),
    }

    # предварительный расчёт Стадии, чтобы знать порядок Непреложных правил
    stage_zone = sa.calculate_stage_zone(
        client_response["qualification"]["fte_a"],
        client_response["qualification"]["fte_b"],
        client_response["qualification"]["fte_c"],
        client_response["qualification"]["fte_d"],
        data,
    )
    stage_id = stage_zone["stage_id"]

    client_response["immutable_rules_pct"] = convert_immutable_rules_pct(
        payload["section7"], stage_id, data)

    client_response["section8_likert_by_kse"] = convert_section8_by_kse(payload["section8"])

    return client_response, q


# ---------------------------------------------------------------------------
# Эндпоинт
# ---------------------------------------------------------------------------

@app.route("/generate-report", methods=["POST"])
def generate_report():
    try:
        payload = request.get_json(force=True)

        data = sa.load_data()
        with open(BASE / "data" / "stage_level_report_texts.json", encoding="utf-8") as f:
            data["stage_level_report_texts"] = json.load(f)
        with open(BASE / "data" / "consulting_programs.json", encoding="utf-8") as f:
            data["consulting_programs"] = json.load(f)
        # Новые (06-07.08.2026) сценарии отклонений от целевых значений —
        # Приоритетные сферы / Коэффициент Строитель-Протектор / Модальность,
        # переписанные Игорем на все возможные случаи (см. build_deviation_
        # scenarios.py — там же комментарий про структуру).
        with open(BASE / "data" / "priority_spheres_scenarios.json", encoding="utf-8") as f:
            data["priority_spheres_scenarios"] = json.load(f)
        with open(BASE / "data" / "builder_protector_scenarios.json", encoding="utf-8") as f:
            data["builder_protector_scenarios"] = json.load(f)
        with open(BASE / "data" / "modality_scenarios.json", encoding="utf-8") as f:
            data["modality_scenarios"] = json.load(f)

        statements = _load_statements()

        client_response, q = build_client_response(payload, data, statements)
        result = sa.diagnose(client_response, data)

        # мета-данные для PDF (обложка/футер) — report_number уже присвоен
        # внутри build_client_response(), повторно счётчик не дёргаем
        report_number = client_response["qualification"]["report_number"]
        client_for_pdf = client_response

        story, page_state = prb.build_story(client_for_pdf, result, data)

        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf, pagesize=A4,
            leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
            title="Полная оценка состояния бизнеса",
        )
        meta = {
            "report_number": report_number,
            "diagnosis_date": client_for_pdf["qualification"]["diagnosis_date"],
            "company": client_for_pdf["qualification"]["company"],
            "name": client_for_pdf["qualification"]["name"],
        }
        doc.build(
            story,
            onFirstPage=partial(prb.draw_cover, meta=meta),
            onLaterPages=partial(prb.draw_footer, meta=meta, page_state=page_state),
        )
        pdf_bytes = buf.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode("ascii")

        # --- Скрипт консультации для Игоря (доп. к клиентскому Отчёту) ---
        # Обёрнуто в try/except: ошибка сборки Скрипта НЕ должна ронять ответ
        # с уже готовым клиентским Отчётом — это дополнительная, не
        # критичная для клиента часть.
        script_pdf_base64 = None
        try:
            adapted = adapt_script_data(result, client_response, data)
            sections = build_script_sections(**adapted)
            script_buf = io.BytesIO()
            build_script_pdf(sections, adapted["qualification"], adapted["stage_name"], script_buf)
            script_pdf_base64 = base64.b64encode(script_buf.getvalue()).decode("ascii")
        except Exception as e:
            print(f"Скрипт консультации НЕ собран (не критично для клиента): {e}")

        # --- Презентация для консультации (доп. к Отчёту и Скрипту) ---
        presentation_base64 = None
        try:
            pptx_data = export_pptx_data(result, client_response, data)
            pres = new_presentation()
            build_slide_cover(pres, pptx_data)
            build_slide_agenda(pres, pptx_data)
            build_slide_your_words(pres, pptx_data)
            build_slide_reasons(pres, pptx_data)
            build_slide_growth_rules(pres, pptx_data)
            build_slide_kse_concept(pres, pptx_data)
            build_slide_symptoms(pres, pptx_data)
            build_slide_maturity_top5(pres, pptx_data)
            build_slide_priority_chain(pres, pptx_data)
            build_slide_cost_only(pres, pptx_data)
            build_slide_full_contrast(pres, pptx_data)
            build_slide_how_to_reach(pres, pptx_data)
            build_slide_kse_list_repeat(pres, pptx_data)
            build_slide_what_order(pres, pptx_data)
            build_slide_house(pres, pptx_data)
            build_slide_how_to_implement(pres, pptx_data)
            build_slide_what_is_program(pres, pptx_data)
            build_slide_program_model(pres, pptx_data)
            build_slide_program_model_practice(pres, pptx_data)
            build_slide_house_continuation(pres, pptx_data)
            build_slide_programs_foundation(pres, pptx_data)
            build_slide_programs_core(pres, pptx_data)
            build_slide_programs_superstructure(pres, pptx_data)
            build_slide_bundle_fork(pres, pptx_data)
            build_slide_guarantee(pres, pptx_data)
            build_slide_loyalty_program(pres, pptx_data)
            build_slide_closing(pres, pptx_data)
            build_slide_mission(pres, pptx_data)
            pptx_buf = io.BytesIO()
            pres.save(pptx_buf)
            presentation_base64 = base64.b64encode(pptx_buf.getvalue()).decode("ascii")
        except Exception as e:
            print(f"Презентация НЕ собрана (не критично для клиента): {e}")

        forward_to_make(client_response, pdf_base64, script_pdf_base64, presentation_base64)

        # 21.08.2026: сохраняем ТОЛЬКО PDF Отчёта (не Скрипт, не Презентацию —
        # те для Игоря) и email клиента в заказ по order_id, если он пришёл
        # от Опросника. Это нужно планировщику отложенной отправки клиенту
        # через 23 часа после оплаты (send_delayed_report_to_clients) — он
        # смотрит именно в эти поля заказа, а не пересобирает отчёт заново.
        order_id = payload.get("order_id")
        if order_id:
            try:
                # 24.08.2026: было _load_orders()/_save_orders() — загрузка и
                # перезапись ВСЕЙ таблицы заказов ради одного order_id.
                order = _load_order(order_id)
                if order is not None:
                    order["report_pdf_base64"] = pdf_base64
                    # Скрипт и Презентация сохраняются ОТДЕЛЬНЫМИ полями и
                    # используются ТОЛЬКО в /admin (кнопка "прислать мне 3
                    # файла"). Автоматический Сц2 (send_delayed_report_to_clients)
                    # эти поля не читает вообще — клиенту они физически не
                    # могут уйти по ошибке.
                    order["script_pdf_base64"] = script_pdf_base64 or ""
                    order["presentation_base64"] = presentation_base64 or ""
                    order["report_number"] = report_number
                    order["client_email"] = q.get("email", "")
                    order["client_name"] = q.get("name", "")
                    _save_order(order_id, order)

                    # 22.08.2026: мгновенное письмо-подтверждение клиенту, сразу
                    # после прохождения опросника (не путать с Сц2 — тем, что
                    # уходит через 23ч с самим Отчётом). Цель — дать клиенту
                    # НАДЁЖНЫЙ канал связи с Игорем на случай проблем, и
                    # подтвердить, что email в заказе совпадает с тем, что
                    # клиент реально использует (в отличие от Планёрки или
                    # Продамус, где клиент мог указать другой email).
                    client_email = q.get("email", "")
                    client_name = q.get("name", "")
                    if client_email:
                        try:
                            send_email_smtp(
                                to_email=client_email,
                                subject="Вы прошли Оценочный опрос — дальнейшие шаги",
                                body=(
                                    f"Здравствуйте{', ' + client_name if client_name else ''}!\n\n"
                                    f"Вы прошли Полную оценку состояния бизнеса по системе "
                                    f"«Возрождение бизнеса». Ваши ответы приняты в обработку — "
                                    f"результаты Чек-апа в виде полного PDF-Отчёта (№{report_number}) "
                                    f"будут Вам отправлены на этот адрес в течение суток.\n\n"
                                    f"Если Вы уже запланировали онлайн-встречу с нашим бизнес-"
                                    f"консультантом для обсуждения результатов диагностики, то мы "
                                    f"скоро с Вами увидимся.\n\n"
                                    f"Если Вы ещё не запланировали онлайн-встречу, Вы можете сделать "
                                    f"это по ссылке {PLANERKA_BOOKING_URL}\n\n"
                                    f"Если в течение суток после прохождения Оценочного опроса Вы не "
                                    f"получили Отчёт (проверьте, пожалуйста, папку «Спам»), напишите "
                                    f"нам:\nreport@fenix-lab.ru\n\n"
                                    f"Ждём Вас на онлайн-встрече!\n\n"
                                    f"С уважением,\nЛаборатория бизнес лидерства «Феникс»"
                                ),
                            )
                        except Exception as e:
                            print(f"DEBUG generate-report: не удалось отправить письмо-подтверждение "
                                  f"клиенту {client_email!r}: {e}", flush=True)
                else:
                    print(f"DEBUG generate-report: order_id {order_id!r} передан, "
                          f"но не найден в БД заказов — Сц2 не сможет отправить клиенту", flush=True)
            except Exception as e:
                print(f"DEBUG generate-report: не удалось сохранить PDF в заказ {order_id!r}: {e}", flush=True)

        return jsonify({
            "ok": True,
            "report_number": report_number,
            "pdf_base64": pdf_base64,
            "script_pdf_base64": script_pdf_base64,  # None, если сборка не удалась — см. лог
            "presentation_base64": presentation_base64,  # None, если сборка не удалась — см. лог
            "diagnose_result": result,
        })

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


def forward_to_make(client_response, pdf_base64, script_pdf_base64=None, presentation_base64=None):
    """
    21.08.2026: несмотря на старое имя (осталось для минимальных правок
    вызывающего кода), функция больше НЕ обращается к Make.com. Вместо
    этого она сразу отправляет письмо Игорю (IGOR_NOTIFICATION_EMAIL) с
    тремя вложениями — Отчёт, Скрипт консультации, Презентация — напрямую
    через SMTP Яндекс.Почты. Обёрнуто в try/except: если отправка письма
    не удалась (например, нет сети до smtp.yandex.ru), клиент на сайте
    всё равно должен увидеть экран "Заключение" (PDF уже успешно создан
    и возвращается в ответе /generate-report) — потерю письма лучше
    разбирать отдельно по логам приложения, чем ронять весь ответ.

    script_pdf_base64 / presentation_base64: Скрипт консультации и
    Презентация — ОБА только для Игоря, клиенту не отправляются.
    """
    q = client_response["qualification"]
    report_number = q.get("report_number", "")
    name = q.get("name", "")
    company = q.get("company", "")
    email = q.get("email", "")
    phone = q.get("phone", "")
    diagnosis_date = q.get("diagnosis_date", "")

    subject = f"Чек-ап №{report_number} — {name or company or email}"
    body = (
        f"Новая диагностика пройдена.\n\n"
        f"Номер отчёта: {report_number}\n"
        f"Имя: {name}\n"
        f"Компания: {company}\n"
        f"Email: {email}\n"
        f"Телефон: {phone}\n"
        f"Дата: {diagnosis_date}\n\n"
        f"Во вложении: Отчёт"
        + (", Скрипт консультации" if script_pdf_base64 else "")
        + (", Презентация" if presentation_base64 else "")
        + "."
    )

    attachments = [("Otchet.pdf", pdf_base64)]
    if script_pdf_base64:
        attachments.append(("Skript_konsultacii.pdf", script_pdf_base64))
    if presentation_base64:
        attachments.append(("Prezentaciya.pptx", presentation_base64))

    try:
        send_email_smtp(
            to_email=IGOR_NOTIFICATION_EMAIL,
            subject=subject,
            body=body,
            attachments=attachments,
        )
        print(f"Уведомление отправлено на {IGOR_NOTIFICATION_EMAIL} (report_number={report_number})", flush=True)
        send_max_alert(f"✅ Новый Чек-ап №{report_number} — уведомление с файлами отправлено на почту")
    except Exception as e:
        import traceback
        print(f"Не удалось отправить уведомление Игорю по SMTP: {e}", flush=True)
        print(traceback.format_exc(), flush=True)
        send_max_alert(
            f"⚠️ Не пришло уведомление о новом Чек-апе (Сц1)\n"
            f"№{report_number}\n"
            f"Причина: {e}\n"
            f"Файлы сохранены в базе — можно достать через /admin"
        )


def send_max_alert(text):
    """Отправляет короткое сервисное сообщение Игорю в MAX (бот "Феникс
    Алерты"). Используется для проактивных уведомлений о сбоях в Сц1/Сц2
    и о "зависших" заказах (см. check_stuck_orders ниже) — план Б на случай
    недоступности Timeweb/почты в критичный момент.

    Намеренно не бросает исключение при неудаче (это сам по себе резервный
    канал уведомлений — если он недоступен, не должно ронять остальной код),
    просто печатает в лог, что не получилось отправить."""
    if not MAX_BOT_TOKEN or not MAX_ALERT_CHAT_ID:
        print("DEBUG send_max_alert: MAX_BOT_TOKEN/MAX_ALERT_CHAT_ID не заданы — уведомление не отправлено", flush=True)
        return
    try:
        resp = requests.post(
            f"https://platform-api2.max.ru/messages?chat_id={MAX_ALERT_CHAT_ID}",
            headers={"Authorization": MAX_BOT_TOKEN, "Content-Type": "application/json"},
            json={"text": text},
            timeout=15,
        )
        if resp.status_code >= 300:
            print(f"DEBUG send_max_alert: MAX API вернул {resp.status_code}: {resp.text[:300]}", flush=True)
    except Exception as e:
        print(f"DEBUG send_max_alert: не удалось отправить уведомление в MAX: {e}", flush=True)


def send_email_smtp(to_email, subject, body, attachments=None, from_email=None):
    """
    Отправляет письмо через SMTP Яндекс.Почты (порт 465, SSL).

    attachments: список кортежей (filename, base64_content). Тип вложения
    (PDF/PPTX) определяется по расширению в filename — content-type ставим
    универсальный application/octet-stream, чтобы не тащить лишнюю логику;
    почтовые клиенты сами разбираются по расширению файла.

    Бросает исключение при ошибке — вызывающий код сам решает, оборачивать
    ли в try/except (для Игоря — да, чтобы не ронять ответ клиенту; для
    отправки клиенту в будущем сценарии отложенной отправки — тоже да, с
    ретраем на следующем проходе планировщика).
    """
    if not SMTP_LOGIN or not SMTP_PASSWORD:
        raise RuntimeError("SMTP_LOGIN/SMTP_PASSWORD не заданы в переменных окружения")

    from_email = from_email or SMTP_LOGIN

    msg = MIMEMultipart()
    from email.utils import formataddr
    from email.header import Header
    msg["From"] = formataddr((str(Header("Лаборатория бизнес лидерства «Феникс»", "utf-8")), from_email))
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for filename, content_base64 in (attachments or []):
        if not content_base64:
            continue
        part = MIMEApplication(base64.b64decode(content_base64))
        part.add_header("Content-Disposition", "attachment", filename=filename)
        msg.attach(part)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
        server.login(SMTP_LOGIN, SMTP_PASSWORD)
        server.sendmail(from_email, [to_email], msg.as_string())


def _load_statements():
    with open(BASE / "data" / "statements.json", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Хранилище заказов Prodamus.
# 19.08.2026, миграция Render → Timeweb Cloud: сначала пробовали файл на
# /var/data (Render — Permission denied на Timeweb), затем /app/data_storage
# (доступен для записи, но подтверждено тестом — стирается при каждой
# пересборке контейнера, т.к. это не постоянный диск). В App Platform
# Timeweb нет опции подключения постоянного тома (Volume), поэтому переходим
# на управляемую БД PostgreSQL — данные не привязаны к конкретному
# контейнеру и переживают любые передеплои.
#
# Требуется переменная окружения DATABASE_URL вида:
#   postgresql://USER:PASSWORD@HOST:PORT/DBNAME
# (Timeweb обычно показывает готовую строку подключения в панели базы
# данных — можно скопировать как есть; если панель даёт только отдельные
# поля Host/Port/User/Password/DBName, собери строку по этому шаблону).
# ---------------------------------------------------------------------------
import psycopg
import psycopg.rows
from psycopg_pool import ConnectionPool

DATABASE_URL = os.environ.get("DATABASE_URL", "")

# 24.08.2026: раньше каждый вызов _get_db_connection() открывал новое
# соединение psycopg.connect(DATABASE_URL) — при SSL verify-full это полный
# TLS-хендшейк с проверкой сертификата на КАЖДЫЙ, даже самый мелкий запрос
# (например, /payment-status при поллинге с фронтенда каждые 1-2 сек).
# Обнаружено как причина зависаний 4-9 сек на переходе Шаг 3 → Опросник.
# Теперь используется пул соединений — TLS-хендшейк происходит один раз на
# каждое из нескольких соединений в пуле, дальше они переиспользуются.
_db_pool: ConnectionPool | None = None


def _get_pool() -> ConnectionPool:
    global _db_pool
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL не задана в переменных окружения")
    if _db_pool is None:
        # min_size/max_size невелики намеренно — конфигурация Timeweb сейчас
        # 1 CPU / 1 ГБ RAM, лишние простаивающие соединения к БД тоже
        # расходуют память. Если позже поднимете тариф — можно увеличить.
        _db_pool = ConnectionPool(DATABASE_URL, min_size=1, max_size=4, open=True)
    return _db_pool


def _get_db_connection():
    """Возвращает контекстный менеджер соединения из пула — использовать
    как и раньше: `with _get_db_connection() as conn: ...`."""
    return _get_pool().connection()


def _init_orders_table():
    """Создаёт таблицу заказов при первом запуске, если её ещё нет."""
    with _get_db_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS orders (
                    order_id TEXT PRIMARY KEY,
                    data JSONB NOT NULL,
                    updated_at TIMESTAMPTZ DEFAULT now()
                )
                """
            )
        conn.commit()


def _load_orders():
    with _get_db_connection() as conn:
        with conn.cursor(row_factory=psycopg.rows.dict_row) as cur:
            cur.execute("SELECT order_id, data FROM orders")
            rows = cur.fetchall()
    return {row["order_id"]: row["data"] for row in rows}


def _save_orders(orders):
    """Перезаписывает все заказы разом — сохраняем ту же сигнатуру функции,
    что была у файловой версии, чтобы не трогать вызывающий код.

    ВНИМАНИЕ: перебирает и перезаписывает КАЖДЫЙ заказ в таблице, даже если
    менялся только один. Годится для мест, где реально нужно сохранить
    несколько заказов разом, но НЕ используйте для горячих путей (частый
    поллинг/один заказ за раз) — там используйте _load_order/_save_order
    ниже, которые работают с одной строкой через WHERE order_id = %s."""
    with _get_db_connection() as conn:
        with conn.cursor() as cur:
            for order_id, order_data in orders.items():
                cur.execute(
                    """
                    INSERT INTO orders (order_id, data, updated_at)
                    VALUES (%s, %s, now())
                    ON CONFLICT (order_id)
                    DO UPDATE SET data = EXCLUDED.data, updated_at = now()
                    """,
                    (order_id, json.dumps(order_data, ensure_ascii=False)),
                )
        conn.commit()


def _load_order(order_id):
    """24.08.2026: точечный запрос ОДНОГО заказа по order_id вместо
    выгрузки всей таблицы (как делал _load_orders()). Используется в
    горячих путях — /payment-status, /verify-payform-redirect — которые
    дергаются поллингом с фронтенда много раз подряд. Возвращает dict с
    данными заказа или None, если заказ не найден."""
    with _get_db_connection() as conn:
        with conn.cursor(row_factory=psycopg.rows.dict_row) as cur:
            cur.execute("SELECT data FROM orders WHERE order_id = %s", (order_id,))
            row = cur.fetchone()
    return row["data"] if row else None


def _save_order(order_id, order_data):
    """24.08.2026: точечное сохранение ОДНОГО заказа вместо перезаписи всей
    таблицы (как делал _save_orders()). Пара к _load_order()."""
    with _get_db_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO orders (order_id, data, updated_at)
                VALUES (%s, %s, now())
                ON CONFLICT (order_id)
                DO UPDATE SET data = EXCLUDED.data, updated_at = now()
                """,
                (order_id, json.dumps(order_data, ensure_ascii=False)),
            )
        conn.commit()


def _count_orders():
    """24.08.2026: лёгкий COUNT(*) для логов вместо загрузки всей таблицы
    только ради len(orders)."""
    with _get_db_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM orders")
            (count,) = cur.fetchone()
    return count


def _find_orders_by_email(email):
    """24.08.2026: поиск заказов по email клиента через SQL-фильтр по полю
    JSONB (data->>'client_email'), а не загрузкой всей таблицы в Python и
    фильтрацией в памяти (как делала админ-панель раньше). Используется в
    /admin — там это разовый ручной поиск, но по мере роста базы полная
    загрузка стала бы всё медленнее, поэтому сразу делаем через WHERE.
    Возвращает список (order_id, data), тот же формат, что ждёт вызывающий
    код."""
    with _get_db_connection() as conn:
        with conn.cursor(row_factory=psycopg.rows.dict_row) as cur:
            cur.execute(
                "SELECT order_id, data FROM orders WHERE lower(data->>'client_email') = lower(%s)",
                (email,),
            )
            rows = cur.fetchall()
    return [(row["order_id"], row["data"]) for row in rows]


# Таблица должна существовать до первого запроса — создаём при старте
# приложения (если DATABASE_URL не задана, не падаем сразу, а только при
# первой реальной попытке чтения/записи заказов — см. _get_db_connection).
if DATABASE_URL:
    try:
        _init_orders_table()
    except Exception as _e:
        print(f"DEBUG: не удалось инициализировать таблицу orders при старте: {_e}", flush=True)


def get_diagnostic_price(stage_id):
    """Та же логика цены, что и в Опроснике (getDiagnosticPrice в JS) —
    держим синхронно на обеих сторонах."""
    return 9900 if int(stage_id) <= 3 else 14900


def get_product_name(stage_id):
    """Название товара для чека Prodamus зависит от диапазона Стадии."""
    if int(stage_id) <= 3:
        return "Полная оценка состояния бизнеса для Стадий 1-3"
    return "Полная оценка состояния бизнеса для Стадий 4-7"


@app.route("/create-payment-link", methods=["OPTIONS"])
def create_payment_link_options():
    return ("", 204)


@app.route("/create-payment-link", methods=["POST"])
def create_payment_link():
    """Принимает от Опросника данные клиента и Стадии, возвращает подписанную
    ссылку на оплату Prodamus (do=pay — сразу открывает страницу оплаты,
    без промежуточного шага получения короткой ссылки) вместе с order_id,
    который Опросник должен сохранить и передавать в /payment-status."""
    try:
        payload = request.get_json(force=True)
        stage_id = payload["stage_id"]
        price = get_diagnostic_price(stage_id)
        order_id = uuid.uuid4().hex[:12]

        # URL, куда Продамус вернёт клиента после оплаты. Раньше был жёстко
        # зашит как голый https://fenix-lab.ru/ — этого достаточно ТОЛЬКО
        # если оплата всегда проходит в отдельном попапе (тот сам закрывается
        # через /payment-status, реальный urlSuccess клиент почти не видит).
        # Но если попап по какой-то причине заблокирован браузером и Опросник
        # переходит на резервный путь (та же вкладка) — клиента нужно вернуть
        # именно на страницу с Опросником (там есть код восстановления
        # состояния, см. tryResumeFromPaymentRedirect() в Опроснике), а не на
        # случайную домашнюю страницу сайта, откуда прогресс было не вернуть
        # (обнаружено на реальном тесте 05.08.2026). Если фронтенд не передал
        # return_url (старая версия Опросника) — используем прежний адрес
        # как безопасный запасной вариант.
        return_url = payload.get("return_url") or "https://fenix-lab.ru/"

        data = {
            "order_id": order_id,
            "customer_phone": normalize_phone(payload.get("phone", "")),
            "customer_email": payload.get("email", ""),
            "products": [
                {
                    "name": get_product_name(stage_id),
                    "price": str(price),
                    "quantity": "1",
                }
            ],
            "customer_extra": payload.get("company", ""),
            "do": "pay",
            # ВАЖНО: переопределяем глобальные Success/Fail URL аккаунта
            # (которые настроены под другой продукт — LMS fenix-lms.ru) —
            # именно для этой платёжной ссылки. Без этого клиента после
            # оплаты редиректило на страницу логина LMS вместо того, чтобы
            # попап Опросника закрылся сам через /payment-status.
            "urlSuccess": return_url,
            "urlReturn": return_url,
        }
        data["signature"] = prodamus_sign(data, PRODAMUS_SECRET_KEY)

        from urllib.parse import urlencode
        query = urlencode(_flatten_for_query(data))
        payment_url = f"{PRODAMUS_FORM_URL}?{query}"

        # 24.08.2026: создание заказа — не требует загрузки всей таблицы
        # заранее (это НОВАЯ запись, не апдейт существующей).
        new_order = {
            "paid": False,
            "stage_id": stage_id,
            "price": price,
            "created_at": datetime.now().isoformat(),
            # UTM-метки клиента, пришедшие с Опросника — сохраняем вместе с
            # заказом, чтобы в момент реальной оплаты (webhook) можно было
            # отправить конверсию в Яндекс.Метрику с правильным источником
            # трафика, а не просто фактом "оплата произошла откуда-то".
            "utm_source": payload.get("utm_source", ""),
            "utm_medium": payload.get("utm_medium", ""),
            "utm_campaign": payload.get("utm_campaign", ""),
            "utm_content": payload.get("utm_content", ""),
            "utm_term": payload.get("utm_term", ""),
            "yandex_client_id": payload.get("yandex_client_id", ""),
        }
        _save_order(order_id, new_order)
        print(f"DEBUG create-payment-link: заказ {order_id!r} сохранён. "
              f"Всего заказов в БД теперь: {_count_orders()}", flush=True)

        return jsonify({"ok": True, "order_id": order_id, "payment_url": payment_url})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


def _flatten_for_query(d, parent_key=""):
    """Строит пары (ключ, значение) в стиле PHP http_build_query, включая
    вложенные словари/списки как products[0][name] и т.д."""
    items = []
    if isinstance(d, dict):
        for k, v in d.items():
            new_key = f"{parent_key}[{k}]" if parent_key else k
            items.extend(_flatten_for_query(v, new_key))
    elif isinstance(d, list):
        for i, v in enumerate(d):
            items.extend(_flatten_for_query(v, f"{parent_key}[{i}]"))
    else:
        items.append((parent_key, d))
    return items


def _unflatten_form(flat: dict) -> dict:
    """Разворачивает плоские ключи формы вида 'products[0][name]' обратно
    в такую же вложенную структуру (словари/списки), какая была на стороне
    Продамус при формировании подписи — иначе JSON для проверки подписи
    получится другим, и подпись не совпадёт, даже если данные верны."""
    root = {}
    for key, value in flat.items():
        if "[" not in key:
            root[key] = value
            continue
        first, rest = key.split("[", 1)
        parts = [first]
        remaining = "[" + rest
        while remaining.startswith("["):
            end = remaining.index("]")
            parts.append(remaining[1:end])
            remaining = remaining[end + 1:]

        node = root
        for i, part in enumerate(parts):
            if i == len(parts) - 1:
                node[part] = value
            else:
                if part not in node or not isinstance(node[part], dict):
                    node[part] = {}
                node = node[part]

    def _convert_sequential_to_lists(node):
        if isinstance(node, dict):
            converted = {k: _convert_sequential_to_lists(v) for k, v in node.items()}
            keys = list(converted.keys())
            if keys and all(k.isdigit() for k in keys):
                if sorted(keys, key=int) == [str(i) for i in range(len(keys))]:
                    return [converted[k] for k in sorted(keys, key=int)]
            return converted
        return node

    return _convert_sequential_to_lists(root)


def send_metrika_offline_conversion(client_id, order_id, price, utm=None):
    """Отправляет серверное подтверждение оплаты в Яндекс.Метрику через
    Measurement Protocol (https://mc.yandex.ru/collect). Не бросает
    исключений наружу — сбой отправки в Метрику не должен ронять обработку
    вебхука Продамус (деньги клиента важнее галочки в аналитике).

    Ограничение Metrika MP: событие можно приклеить к существующему визиту
    только если визит завершился менее 12 часов назад. Для нашего сценария
    (Опросник → оплата → вебхук — обычно минуты, редко часы) этого более
    чем достаточно; если клиент вернулся к оплате через день и больше —
    событие всё равно отправится и создаст НОВЫЙ визит с этим ClientID
    (см. документацию Метрики), просто не свяжется со старым визитом.
    """
    if not YANDEX_MP_TOKEN:
        print("DEBUG metrika MP: YANDEX_MP_TOKEN не задан — пропускаем отправку", flush=True)
        return
    if not client_id:
        print(f"DEBUG metrika MP: для заказа {order_id!r} нет yandex_client_id — "
              f"пропускаем отправку (клиент пришёл без JS-счётчика?)", flush=True)
        return
    try:
        params = {
            "tid": YANDEX_METRIKA_COUNTER_ID,
            "cid": client_id,
            "ms": YANDEX_MP_TOKEN,
            "t": "event",
            "ea": "diag_payment_confirmed_server",
            "dl": "https://fenix-lab.ru/online-check-up",  # реальный адрес приложения Чек-апа
            "ep_order_id": order_id,
            "ep_price": price,
        }
        if utm:
            for key in ("utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term"):
                if utm.get(key):
                    params[f"ep_{key}"] = utm[key]
        resp = requests.get(YANDEX_MP_COLLECT_URL, params=params, timeout=5)
        print(f"DEBUG metrika MP: заказ {order_id!r} отправлен, "
              f"HTTP {resp.status_code}", flush=True)
    except Exception as e:
        print(f"DEBUG metrika MP: ошибка отправки для заказа {order_id!r}: {e}", flush=True)


@app.route("/prodamus-webhook", methods=["POST"])
def prodamus_webhook():
    """Приём уведомления об оплате от Prodamus. URL этого эндпоинта нужно
    один раз прописать в личном кабинете Продамус (Настройки платёжной
    страницы → URL для уведомлений), а не передавать в каждой ссылке —
    сам Продамус подтверждает, что параметр urlNotification в ссылке
    поддерживается только для одной конкретной CMS (Advantshop)."""
    incoming_flat = request.form.to_dict()
    signature = request.headers.get("Sign", "")

    incoming = _unflatten_form(incoming_flat)

    # ВРЕМЕННОЕ ЛОГИРОВАНИЕ для диагностики: печатаем всё, что реально пришло,
    # чтобы точно увидеть названия полей (order_id / order_num и т.д.) в логах
    # Render. Убрать после того, как разберёмся с сопоставлением заказов.
    print("=== PRODAMUS WEBHOOK RAW ===", flush=True)
    print("incoming_flat:", incoming_flat, flush=True)
    print("incoming (unflattened):", incoming, flush=True)
    print("signature header:", signature, flush=True)
    print("=============================", flush=True)

    data_to_verify = {k: v for k, v in incoming.items() if k != "signature"}
    if not prodamus_verify(data_to_verify, PRODAMUS_SECRET_KEY, signature):
        return "signature incorrect", 400

    # ВАЖНО (подтверждено эмпирически по логам реального вебхука 24.07.2026):
    # при формировании ссылки на оплату наш order_id передаётся под ключом
    # order_id — Продамус принимает именно его как "ваш номер заказа".
    # Но в уведомлении об оплате Продамус кладёт этот же наш идентификатор
    # обратно под ключом order_num, а под order_id — уже СВОЙ внутренний
    # номер заказа (не наш). Поэтому сопоставлять с нашей базой
    # поиск заказа в БД нужен именно по order_num, а не по order_id.
    order_id = incoming.get("order_num", "")
    payment_status = incoming.get("payment_status", "")

    # 24.08.2026: было _load_orders() — загрузка ВСЕЙ таблицы (включая
    # печать списка всех order_id в БД, что тоже росло бы вместе с базой).
    # Заменено на точечный запрос одного заказа.
    order = _load_order(order_id)
    print(f"DEBUG webhook: order_id из вебхука = {order_id!r}", flush=True)
    print(f"DEBUG webhook: заказ найден в БД? {order is not None}", flush=True)
    if order is not None:
        order["paid"] = (payment_status == "success")
        order["paid_at"] = datetime.now().isoformat()
        order["raw_status"] = payment_status
        _save_order(order_id, order)
        print(f"DEBUG webhook: заказ {order_id!r} обновлён, paid={order['paid']}", flush=True)

        if order["paid"]:
            send_metrika_offline_conversion(
                client_id=order.get("yandex_client_id", ""),
                order_id=order_id,
                price=order.get("price", ""),
                utm={
                    "utm_source": order.get("utm_source", ""),
                    "utm_medium": order.get("utm_medium", ""),
                    "utm_campaign": order.get("utm_campaign", ""),
                    "utm_content": order.get("utm_content", ""),
                    "utm_term": order.get("utm_term", ""),
                },
            )
    else:
        print(f"DEBUG webhook: ЗАКАЗ {order_id!r} НЕ НАЙДЕН — обновление пропущено!", flush=True)

    return "success", 200


@app.route("/verify-payform-redirect", methods=["OPTIONS"])
def verify_payform_redirect_options():
    return ("", 204)


@app.route("/verify-payform-redirect", methods=["POST"])
def verify_payform_redirect():
    """Проверка подписи редиректа urlSuccess — позволяет открыть доступ сразу
    по возврату клиента со страницы оплаты, не дожидаясь вебхука (у вебхука
    подтверждена задержка ~59 сек на стороне Продамус). Опросник вызывает
    этот эндпоинт сразу после того, как Prodamus вернул клиента на urlSuccess
    с параметрами _payform_status/_payform_id/_payform_order_id/_payform_sign
    в URL редиректа.

    ВАЖНО (подтверждено поддержкой Продамус 01.08.2026): подпись здесь
    считается СТРОГО по трём полям _payform_status/_payform_id/
    _payform_order_id — в отличие от вебхука, где подписывается весь $_POST.
    Тот же секретный ключ, что и для вебхука (отдельного ключа под редирект
    не существует). payform_order_id — это наш собственный order_id (тот,
    что мы передавали как order_id при создании ссылки на оплату) — здесь,
    в отличие от вебхука, Продамус НЕ подменяет его своим внутренним номером.

    Вебхук (/prodamus-webhook) оставлен как есть — работает в фоне для
    подстраховки/синхронизации, на случай если этот эндпоинт по какой-то
    причине не был вызван (например, клиент закрыл вкладку сразу после
    оплаты, до того как успел отработать JS Опросника)."""
    try:
        payload = request.get_json(force=True)
        status = payload.get("_payform_status", "")
        payform_id = payload.get("_payform_id", "")
        payform_order_id = payload.get("_payform_order_id", "")
        signature = payload.get("_payform_sign", "")

        fields_to_verify = {
            "_payform_status": status,
            "_payform_id": payform_id,
            "_payform_order_id": payform_order_id,
        }
        if not prodamus_verify(fields_to_verify, PRODAMUS_SECRET_KEY, signature):
            print(f"DEBUG verify-payform-redirect: подпись НЕ совпала для "
                  f"order_id={payform_order_id!r}", flush=True)
            return jsonify({"ok": False, "valid": False, "error": "signature incorrect"}), 400

        order_id = payform_order_id
        paid = (status == "success")

        # 24.08.2026: было _load_orders()/_save_orders() — загрузка и
        # перезапись ВСЕЙ таблицы заказов ради проверки одного order_id.
        # Заменено на точечный запрос/апдейт одной строки — это и была
        # основная причина зависаний 4-9 сек на поллинге при переходе
        # Шаг 3 → Опросник (нагрузка росла вместе с числом заказов в БД).
        order = _load_order(order_id)
        if order is not None:
            if paid:
                order["paid"] = True
                order["paid_at"] = datetime.now().isoformat()
                order["raw_status"] = status
                order["confirmed_via"] = "redirect"
                _save_order(order_id, order)
            print(f"DEBUG verify-payform-redirect: заказ {order_id!r} проверен, "
                  f"paid={paid}", flush=True)
        else:
            print(f"DEBUG verify-payform-redirect: заказ {order_id!r} НЕ НАЙДЕН "
                  f"в базе заказов — подпись верна, но сопоставить с заказом "
                  f"не удалось", flush=True)

        return jsonify({"ok": True, "valid": True, "paid": paid})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/payment-status/<order_id>", methods=["GET"])
def payment_status(order_id):
    """Опросник опрашивает этот эндпоинт после открытия окна оплаты, чтобы
    узнать, подтвердил ли Продамус оплату, прежде чем открыть доступ
    к разделам диагностики.

    24.08.2026: было _load_orders() — загрузка ВСЕЙ таблицы заказов ради
    статуса одного order_id, при поллинге каждые 1-2 сек это и было
    основной причиной зависаний. Заменено на точечный запрос."""
    order = _load_order(order_id)
    if not order:
        return jsonify({"ok": False, "error": "order not found"}), 404
    return jsonify({"ok": True, "paid": order.get("paid", False)})


# ---------------------------------------------------------------------------
# 21.08.2026: замена Сценариев 2 и 3 Make.com ("Отложенная отправка клиенту
# через 23ч" и "Ручная переотправка PDF клиенту"). Раньше оба брали готовый
# PDF с Google Диска по File ID из Google Таблицы. Теперь PDF уже лежит в
# PostgreSQL (поле report_pdf_base64 заказа, см. /generate-report выше) —
# планировщик просто читает его оттуда и шлёт клиенту через тот же SMTP.
#
# КЛИЕНТУ УХОДИТ ТОЛЬКО ОТЧЁТ. Ни Скрипт консультации, ни Презентация сюда
# физически не попадают — они даже не читаются в этой функции, их просто
# нет в переменных ниже.
# ---------------------------------------------------------------------------
from datetime import timedelta

DELAYED_SEND_HOURS = 23

# Прямая ссылка на страницу записи на онлайн-встречу (та же Планёрка, что
# встроена виджетом на экране "Заключение" опросника) — используется в
# письмах клиенту (Сц2/Сц3), чтобы напомнить о встрече или дать возможность
# записаться повторно, если клиент забыл или передумал.
PLANERKA_BOOKING_URL = "https://planerka.app/meet/igor-balandin-sfkybs/polnaya-ocenka-sostoyaniya-biznesa"


def send_delayed_report_to_clients():
    """Раз в час (см. планировщик ниже) ищет в БД заказы, которые:
    - оплачены (paid = true)
    - отчёт уже сгенерирован (есть report_pdf_base64)
    - с момента оплаты прошло >= DELAYED_SEND_HOURS часов
    - клиенту ещё не отправлялось (нет client_sent_at)
    и отправляет каждому такому клиенту письмо ТОЛЬКО с PDF Отчёта."""
    try:
        orders = _load_orders()
    except Exception as e:
        print(f"DEBUG Сц2: не удалось загрузить заказы: {e}", flush=True)
        return

    now = datetime.now()
    sent_count = 0

    for order_id, order in orders.items():
        if not order.get("paid"):
            continue
        pdf_base64 = order.get("report_pdf_base64")
        if not pdf_base64:
            continue
        if order.get("client_sent_at"):
            continue  # уже отправлено раньше

        paid_at_raw = order.get("paid_at")
        if not paid_at_raw:
            continue
        try:
            paid_at = datetime.fromisoformat(paid_at_raw)
        except ValueError:
            continue

        if now - paid_at < timedelta(hours=DELAYED_SEND_HOURS):
            continue  # ещё рано

        client_email = order.get("client_email", "")
        if not client_email:
            print(f"DEBUG Сц2: у заказа {order_id!r} нет client_email — пропускаем", flush=True)
            continue

        report_number = order.get("report_number", "")
        client_name = order.get("client_name", "")

        try:
            send_email_smtp(
                to_email=client_email,
                subject=f"Ваш отчёт «Полная оценка состояния бизнеса» №{report_number}",
                body=(
                    f"Здравствуйте{', ' + client_name if client_name else ''}!\n\n"
                    f"Спасибо, что прошли диагностику бизнеса — во вложении ваш "
                    f"персональный Отчёт (№{report_number}) по нашей авторской "
                    f"аналитической модели «Возрождение бизнеса».\n\n"
                    f"Если вы уже записались на онлайн-встречу с нашим бизнес-"
                    f"консультантом — мы ждём вас в назначенное время, там подробно "
                    f"разберём результаты и наметим шаги для внедрения недостающих "
                    f"Ключевых системных элементов в вашем бизнесе.\n\n"
                    f"Если вы ещё не выбрали дату и время (или хотите перенести "
                    f"встречу) — сделать это можно здесь:\n"
                    f"{PLANERKA_BOOKING_URL}\n\n"
                    f"С уважением,\nЛаборатория бизнес лидерства «Феникс»"
                ),
                attachments=[("Otchet.pdf", pdf_base64)],
            )
            orders[order_id]["client_sent_at"] = now.isoformat()
            # 24.08.2026: было _save_orders(orders) — перезапись ВСЕЙ
            # таблицы ради одного изменённого заказа. Само сканирование
            # всех заказов раз в час — нормально, а вот перезапись всех
            # при каждой найденной отправке — нет.
            _save_order(order_id, orders[order_id])
            sent_count += 1
            print(f"DEBUG Сц2: Отчёт отправлен клиенту {client_email} (order_id={order_id!r})", flush=True)
            send_max_alert(f"✅ Отчёт отправлен клиенту {client_email} (№{report_number})")
        except Exception as e:
            import traceback
            print(f"DEBUG Сц2: не удалось отправить Отчёт клиенту для заказа {order_id!r}: {e}", flush=True)
            print(traceback.format_exc(), flush=True)
            send_max_alert(
                f"⚠️ Клиент не получил Отчёт (Сц2)\n"
                f"№{report_number}, {client_email}\n"
                f"Причина: {e}\n"
                f"Нужно отправить вручную через /admin"
            )

    if sent_count:
        print(f"DEBUG Сц2: всего отправлено писем клиентам за этот проход: {sent_count}", flush=True)


STUCK_ORDER_HOURS = 25  # на 2 часа больше нормы (23ч) — запас, чтобы не дублировать обычную задержку


def check_stuck_orders():
    """Отдельная проверка (тот же часовой планировщик): заказы, где отчёт
    готов, но клиенту не отправлено спустя STUCK_ORDER_HOURS часов — сигнал
    даже если явных ошибок в логах не было (например, само приложение было
    недоступно в момент, когда должен был сработать send_delayed_report_to_clients,
    и просто пропустило проход)."""
    try:
        orders = _load_orders()
    except Exception as e:
        print(f"DEBUG check_stuck_orders: не удалось загрузить заказы: {e}", flush=True)
        return

    now = datetime.now()
    for order_id, order in orders.items():
        if not order.get("paid") or not order.get("report_pdf_base64") or order.get("client_sent_at"):
            continue
        # уже предупреждали про этот заказ раньше — не дублируем на каждый проход
        if order.get("stuck_alert_sent"):
            continue
        paid_at_raw = order.get("paid_at")
        if not paid_at_raw:
            continue
        try:
            paid_at = datetime.fromisoformat(paid_at_raw)
        except ValueError:
            continue
        if now - paid_at < timedelta(hours=STUCK_ORDER_HOURS):
            continue

        client_email = order.get("client_email", "")
        report_number = order.get("report_number", "")
        send_max_alert(
            f"🔴 Клиент {client_email} не получил Отчёт больше {STUCK_ORDER_HOURS} часов — нужна проверка\n"
            f"№{report_number}, order_id: {order_id}\n"
            f"Проверь вручную через /admin"
        )
        orders[order_id]["stuck_alert_sent"] = now.isoformat()
        _save_order(order_id, orders[order_id])


@app.route("/resend-report/<order_id>", methods=["POST"])
def resend_report(order_id):
    """Замена Сценария 3 Make.com ("Ручная переотправка PDF клиенту").
    Игорь вызывает этот эндпоинт вручную (например, curl или Postman) с
    конкретным order_id, если клиент не получил письмо и просит прислать
    повторно. Клиенту уходит ТОЛЬКО Отчёт, как и в основном отложенном
    сценарии — Скрипт/Презентация здесь тоже не участвуют."""
    order = _load_order(order_id)
    if not order:
        return jsonify({"ok": False, "error": "order not found"}), 404

    pdf_base64 = order.get("report_pdf_base64")
    if not pdf_base64:
        return jsonify({"ok": False, "error": "report not generated for this order yet"}), 400

    client_email = order.get("client_email", "")
    if not client_email:
        return jsonify({"ok": False, "error": "no client_email stored for this order"}), 400

    report_number = order.get("report_number", "")
    client_name = order.get("client_name", "")

    try:
        send_email_smtp(
            to_email=client_email,
            subject=f"Ваш отчёт «Полная оценка состояния бизнеса» №{report_number} (повторно)",
            body=(
                f"Здравствуйте{', ' + client_name if client_name else ''}!\n\n"
                f"Направляем повторно ваш персональный Отчёт (№{report_number}) "
                f"по результатам диагностики бизнеса.\n\n"
                f"Если вы уже записались на онлайн-встречу с нашим бизнес-"
                f"консультантом — мы ждём вас в назначенное время. Если ещё "
                f"не выбрали дату и время (или хотите перенести встречу) — "
                f"сделать это можно здесь:\n"
                f"{PLANERKA_BOOKING_URL}\n\n"
                f"С уважением,\nЛаборатория бизнес лидерства «Феникс»"
            ),
            attachments=[("Otchet.pdf", pdf_base64)],
        )
        order["client_sent_at"] = datetime.now().isoformat()
        _save_order(order_id, order)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# Планировщик: проверяет заказы на отправку раз в час. Работает внутри того
# же процесса gunicorn — отдельный сервер/крон не нужен. При нескольких
# worker'ах (--workers 2, как сейчас на Timeweb) планировщик стартует в
# КАЖДОМ из них, что может привести к повторной отправке одного и того же
# письма — WERKZEUG_RUN_MAIN тут не поможет (это gunicorn, не werkzeug).
# Проверка client_sent_at в БД перед отправкой снижает риск дублей (после
# первой отправки одним worker'ом, второй увидит client_sent_at уже
# заполненным), но полностью не исключает гонку, если оба worker'а
# одновременно попадут в проход в одну и ту же секунду — с частотой раз в
# час и текущим объёмом заказов это практически исключено.
try:
    from apscheduler.schedulers.background import BackgroundScheduler

    _scheduler = BackgroundScheduler(timezone="UTC")
    _scheduler.add_job(send_delayed_report_to_clients, "interval", hours=1, id="send_delayed_report")
    _scheduler.add_job(check_stuck_orders, "interval", hours=1, id="check_stuck_orders")
    _scheduler.start()
    print("DEBUG: планировщик отложенной отправки клиентам запущен (раз в час)", flush=True)
except Exception as e:
    print(f"DEBUG: не удалось запустить планировщик отложенной отправки: {e}", flush=True)


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "fenix-report-generator"})


# ---------------------------------------------------------------------------
# 21.08.2026: служебная панель "План Б". Нужна на случай, если Timeweb
# временно недоступен (например, во время DDoS-атаки на инфраструктуру, как
# было 21.08.2026) в момент, когда должен был сработать Сц1 (уведомление
# Игорю) или Сц2 (отложенная отправка клиенту). Данные заказа и PDF-файлы к
# этому моменту уже сохранены в PostgreSQL (см. /generate-report) — эта
# страница просто даёт Игорю ручной доступ к ним по email клиента, без
# необходимости просить помощь с SQL-запросами.
# ---------------------------------------------------------------------------
from functools import wraps
from flask import session, redirect, url_for, request as flask_request


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return wrapper


ADMIN_LOGIN_PAGE = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">
<title>Вход — Феникс</title>
<style>body{{font-family:sans-serif;max-width:360px;margin:80px auto;padding:0 16px}}
input{{width:100%;padding:8px;margin:6px 0;box-sizing:border-box}}
button{{width:100%;padding:10px;background:#D5530B;color:#fff;border:none;border-radius:4px;cursor:pointer}}
.error{{color:#c00}}</style></head><body>
<h2>Служебная панель «Феникс»</h2>
{error_html}
<form method="post">
<input type="text" name="username" placeholder="Логин" required autofocus>
<input type="password" name="password" placeholder="Пароль" required>
<button type="submit">Войти</button>
</form></body></html>
"""


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    error_html = ""
    if flask_request.method == "POST":
        username = flask_request.form.get("username", "")
        password = flask_request.form.get("password", "")
        if ADMIN_USERNAME and ADMIN_PASSWORD and username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["admin_logged_in"] = True
            return redirect(url_for("admin_dashboard"))
        error_html = '<p class="error">Неверный логин или пароль</p>'
    return ADMIN_LOGIN_PAGE.format(error_html=error_html)


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.pop("admin_logged_in", None)
    return redirect(url_for("admin_login"))


ADMIN_DASHBOARD_PAGE = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">
<title>Панель — Феникс</title>
<style>body{{font-family:sans-serif;max-width:640px;margin:40px auto;padding:0 16px}}
input{{padding:8px;width:70%}}
button{{padding:8px 14px;background:#D5530B;color:#fff;border:none;border-radius:4px;cursor:pointer}}
table{{width:100%;border-collapse:collapse;margin-top:20px}}
td,th{{border:1px solid #ccc;padding:8px;text-align:left}}
.msg{{color:green}} .err{{color:#c00}}
form.inline{{display:inline}}</style></head><body>
<h2>Поиск заказа клиента</h2>
<form method="get">
<input type="text" name="email" placeholder="email клиента" value="{email_value}">
<button type="submit">Найти</button>
</form>
{message_html}
{results_html}
<form method="post" action="/admin/logout" style="margin-top:30px">
<button type="submit" style="background:#666">Выйти</button>
</form>
</body></html>
"""


@app.route("/admin", methods=["GET"])
@admin_required
def admin_dashboard():
    email = flask_request.args.get("email", "").strip().lower()
    message_html = ""
    results_html = ""

    if flask_request.args.get("msg"):
        results_html_msg = flask_request.args.get("msg")
        is_err = flask_request.args.get("err") == "1"
        message_html = f'<p class="{"err" if is_err else "msg"}">{results_html_msg}</p>'

    if email:
        # 24.08.2026: было _load_orders() + фильтрация в Python — загрузка
        # ВСЕЙ таблицы ради поиска по одному email. Заменено на SQL-фильтр
        # по полю data->>'client_email' (_find_orders_by_email).
        matches = _find_orders_by_email(email)
        if not matches:
            results_html = f"<p>Заказы с email <b>{email}</b> не найдены.</p>"
        else:
            rows = ""
            for order_id, o in matches:
                report_ready = "Да" if o.get("report_pdf_base64") else "Нет"
                sent_to_client = o.get("client_sent_at", "") or "не отправлялось"
                rows += f"""
                <tr>
                    <td>{order_id}</td>
                    <td>№{o.get('report_number', '—')}</td>
                    <td>{'Оплачен' if o.get('paid') else 'Не оплачен'}</td>
                    <td>Отчёт готов: {report_ready}</td>
                    <td>Клиенту: {sent_to_client}</td>
                    <td>
                        <form class="inline" method="post" action="/admin/send-client/{order_id}">
                            <button type="submit">Отправить Отчёт клиенту</button>
                        </form>
                        <form class="inline" method="post" action="/admin/send-me/{order_id}">
                            <button type="submit">Прислать мне 3 файла</button>
                        </form>
                    </td>
                </tr>"""
            results_html = f"<table><tr><th>order_id</th><th>№</th><th>Оплата</th><th>Отчёт</th><th>Отправка клиенту</th><th>Действия</th></tr>{rows}</table>"

    return ADMIN_DASHBOARD_PAGE.format(email_value=email, message_html=message_html, results_html=results_html)


@app.route("/admin/send-client/<order_id>", methods=["POST"])
@admin_required
def admin_send_client(order_id):
    """Кнопка «Отправить Отчёт клиенту» — уходит ТОЛЬКО PDF Отчёта, тот же
    текст письма, что и в автоматическом Сц2."""
    order = _load_order(order_id)
    if not order or not order.get("report_pdf_base64"):
        return redirect(url_for("admin_dashboard", email=order.get("client_email", "") if order else "",
                                 msg="Отчёт для этого заказа ещё не готов", err="1"))

    client_email = order.get("client_email", "")
    report_number = order.get("report_number", "")
    client_name = order.get("client_name", "")
    try:
        send_email_smtp(
            to_email=client_email,
            subject=f"Ваш отчёт «Полная оценка состояния бизнеса» №{report_number}",
            body=(
                f"Здравствуйте{', ' + client_name if client_name else ''}!\n\n"
                f"Направляем ваш персональный Отчёт (№{report_number}) по результатам "
                f"диагностики бизнеса.\n\n"
                f"Если вы уже записались на онлайн-встречу с нашим бизнес-консультантом "
                f"— мы ждём вас в назначенное время. Если ещё не выбрали дату и время "
                f"(или хотите перенести встречу) — сделать это можно здесь:\n"
                f"{PLANERKA_BOOKING_URL}\n\n"
                f"С уважением,\nЛаборатория бизнес лидерства «Феникс»"
            ),
            attachments=[("Otchet.pdf", order["report_pdf_base64"])],
        )
        order["client_sent_at"] = datetime.now().isoformat()
        _save_order(order_id, order)
        return redirect(url_for("admin_dashboard", email=client_email, msg="Отчёт отправлен клиенту"))
    except Exception as e:
        return redirect(url_for("admin_dashboard", email=client_email, msg=f"Ошибка отправки: {e}", err="1"))


@app.route("/admin/send-me/<order_id>", methods=["POST"])
@admin_required
def admin_send_me(order_id):
    """Кнопка «Прислать мне 3 файла» — Отчёт + Скрипт консультации +
    Презентация уходят на IGOR_NOTIFICATION_EMAIL, для подготовки к
    консультации, если исходное письмо Сц1 не дошло."""
    order = _load_order(order_id)
    client_email = order.get("client_email", "") if order else ""
    if not order or not order.get("report_pdf_base64"):
        return redirect(url_for("admin_dashboard", email=client_email, msg="Файлы для этого заказа ещё не готовы", err="1"))

    report_number = order.get("report_number", "")
    attachments = [("Otchet.pdf", order["report_pdf_base64"])]
    if order.get("script_pdf_base64"):
        attachments.append(("Skript_konsultacii.pdf", order["script_pdf_base64"]))
    if order.get("presentation_base64"):
        attachments.append(("Prezentaciya.pptx", order["presentation_base64"]))

    try:
        send_email_smtp(
            to_email=IGOR_NOTIFICATION_EMAIL,
            subject=f"[Ручная отправка] Чек-ап №{report_number} — {order.get('client_name', '') or client_email}",
            body=(
                f"Файлы по заказу {order_id} (email клиента: {client_email}), "
                f"отправлены вручную через /admin.\n\nВо вложении: "
                + ", ".join(name for name, _ in attachments) + "."
            ),
            attachments=attachments,
        )
        return redirect(url_for("admin_dashboard", email=client_email, msg="Файлы отправлены вам на почту"))
    except Exception as e:
        return redirect(url_for("admin_dashboard", email=client_email, msg=f"Ошибка отправки: {e}", err="1"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
