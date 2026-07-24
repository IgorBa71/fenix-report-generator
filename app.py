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
import uuid
from datetime import datetime
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

# URL вебхука Сценария 1 в Make.com ("Мгновенное уведомление мне + PDF на
# Диск") — после генерации PDF этот сервис сам отправляет туда готовый
# результат (report_number, контакты клиента, PDF в base64), а дальше Make
# кладёт файл на Google Диск, пишет строку в Google Таблицу и шлёт письмо.
MAKE_SC1_WEBHOOK_URL = "https://hook.us2.make.com/7xgojg4ccxh1opdd5d1xn8t52666p2ik"

# ---------------------------------------------------------------------------
# Интеграция оплаты Prodamus (HMAC-подписанные ссылки, без клиентского
# JS-виджета "Единое окно" — тот не гарантирует корректное формирование
# счёта, что подтвердила сама поддержка Продамус). ВАЖНО: секретный ключ
# и адрес формы нужно свериться с личным кабинетом Продамус — ниже
# указаны значения из предыдущей сессии настройки, возможно устарели.
# ---------------------------------------------------------------------------
PRODAMUS_FORM_URL = "https://fenix-lab.payform.ru/"
PRODAMUS_SECRET_KEY = "fd2f5514f3c5aa90aa6cd0ab0f7352f37ba23f810646057efd0ab1945ce0bc7f"
ORDERS_FILE = Path("/var/data/prodamus_orders.json")

app = Flask(__name__)

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
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
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
            "phone": q.get("phone", ""),
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
        },
        "flow_a_dimensions": client_dimensions,
        "challenge_scores": payload["section2"],
        "section8_likert_by_kse": None,  # заполним ниже, нужен stage/kse_list
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

        forward_to_make(client_response, pdf_base64)

        return jsonify({
            "ok": True,
            "report_number": report_number,
            "pdf_base64": pdf_base64,
            "diagnose_result": result,
        })

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


def forward_to_make(client_response, pdf_base64):
    """
    Отправляет готовый отчёт на вебхук Сценария 1 в Make.com — тот кладёт
    PDF на Google Диск, пишет строку в Google Таблицу и шлёт письмо-
    уведомление. Обёрнуто в try/except: если Make временно недоступен,
    клиент на сайте всё равно должен увидеть экран "Заключение" (PDF уже
    успешно создан) — потерю доставки лучше разбирать отдельно по логам
    Render, чем ронять весь ответ пользователю.
    """
    q = client_response["qualification"]
    payload = {
        "report_number": q["report_number"],
        "name": q.get("name", ""),
        "company": q.get("company", ""),
        "email": q.get("email", ""),
        "phone": q.get("phone", ""),
        "diagnosis_date": q["diagnosis_date"],
        "pdf_base64": pdf_base64,
    }
    try:
        resp = requests.post(MAKE_SC1_WEBHOOK_URL, json=payload, timeout=25)
        print(f"Переслано в Make (Сц1): статус {resp.status_code}")
    except Exception as e:
        print(f"Не удалось переслать отчёт в Make (Сц1): {e}")


def _load_statements():
    with open(BASE / "data" / "statements.json", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Хранилище заказов Prodamus (на постоянном диске /var/data — тот же, что
# используется для счётчика отчётов, — не сбрасывается при передеплое)
# ---------------------------------------------------------------------------

def _load_orders():
    if ORDERS_FILE.exists():
        return json.loads(ORDERS_FILE.read_text(encoding="utf-8"))
    return {}


def _save_orders(orders):
    ORDERS_FILE.parent.mkdir(parents=True, exist_ok=True)
    ORDERS_FILE.write_text(json.dumps(orders, ensure_ascii=False, indent=2), encoding="utf-8")


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

        data = {
            "order_id": order_id,
            "customer_phone": payload.get("phone", ""),
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
        }
        data["signature"] = prodamus_sign(data, PRODAMUS_SECRET_KEY)

        from urllib.parse import urlencode
        query = urlencode(_flatten_for_query(data))
        payment_url = f"{PRODAMUS_FORM_URL}?{query}"

        orders = _load_orders()
        orders[order_id] = {
            "paid": False,
            "stage_id": stage_id,
            "price": price,
            "created_at": datetime.now().isoformat(),
        }
        _save_orders(orders)

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
    # prodamus_orders.json нужно именно по order_num, а не по order_id.
    order_id = incoming.get("order_num", "")
    payment_status = incoming.get("payment_status", "")

    orders = _load_orders()
    if order_id in orders:
        orders[order_id]["paid"] = (payment_status == "success")
        orders[order_id]["paid_at"] = datetime.now().isoformat()
        orders[order_id]["raw_status"] = payment_status
        _save_orders(orders)

    return "success", 200


@app.route("/payment-status/<order_id>", methods=["GET"])
def payment_status(order_id):
    """Опросник опрашивает этот эндпоинт после открытия окна оплаты, чтобы
    узнать, подтвердил ли Продамус оплату, прежде чем открыть доступ
    к разделам диагностики."""
    orders = _load_orders()
    order = orders.get(order_id)
    if not order:
        return jsonify({"ok": False, "error": "order not found"}), 404
    return jsonify({"ok": True, "paid": order.get("paid", False)})


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "fenix-report-generator"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
