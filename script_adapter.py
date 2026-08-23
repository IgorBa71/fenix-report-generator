# -*- coding: utf-8 -*-
"""
script_adapter.py — переводит РЕАЛЬНЫЙ результат пайплайна (build_client_response()
+ sa.diagnose(), т.е. то же самое, что app.py кладёт в diagnose_result) в параметры,
которые ожидает build_full_script() из consultation_script_builder.py.

Ничего не пересчитывает заново — берёт готовые данные из diagnose_result и
data["consulting_programs"] (та же авторитетная структура цен/сроков, что и в
самом PDF-отчёте, report_text_generator.py), не выдумывает демо-цифры.
"""
import re

from report_text_generator import _normalize_tier, TIER_HEADERS
from consultation_script_builder import compute_bundle_fork

TIER_KEY_TO_RU_TITLE = {"фундамент": "Фундамент", "ядро": "Ядро", "надстройка": "Надстройка"}


# ---------------------------------------------------------------------------
# 1. Психографический тип — НЕ вычисляется алгоритмом (см. переписку в других
#    чатах проекта): это прямой ответ клиента, Шаг 3 регистрации Опросника,
#    ранжирование 4 утверждений. Должен приходить в payload["qualification"]
#    ["psychographic"] в формате {"Мыслитель": 1, ...}. build_client_response()
#    в текущем app.py его ЕЩЁ НЕ прокидывает в client_response — читаем из
#    сырого payload напрямую, пока это не исправлено в app.py.
# ---------------------------------------------------------------------------

def get_psychographic_ranks(client_response, payload=None):
    """Предпочитаем client_response (после исправления build_client_response()
    в app.py — см. app.py) — совместимость с payload оставлена на случай, если
    где-то ещё используется client_response старого формата."""
    ranks = client_response.get("qualification", {}).get("psychographic")
    if not ranks and payload:
        ranks = payload.get("qualification", {}).get("psychographic")
    if not ranks:
        raise ValueError(
            "Нет поля 'psychographic' (ранжирование Мыслитель/Делатель/"
            "Чувствователь/Наблюдатель) ни в client_response, ни в payload. "
            "Без него Скрипт собрать нельзя — это прямой ответ клиента с Шага "
            "3 регистрации Опросника, а не то, что вычисляет scoring_algorithm."
        )
    return ranks


# ---------------------------------------------------------------------------
# 2. Раздел 9 «Ваш взгляд на ситуацию» — тоже прямой ввод клиента, не часть
#    scoring_algorithm. Ожидаемые ключи (см. consultation_script_builder.py):
#    problem1/2/3, whyTheseChallenges, whyCantSolve, costOfInaction,
#    priceOfInaction, dreamOutcome.
# ---------------------------------------------------------------------------

SECTION9_KEYS = ("problem1", "problem2", "problem3", "whyTheseChallenges", "whyCantSolve",
                  "costOfInaction", "priceOfInaction", "dreamOutcome")


def get_section9(client_response, payload=None):
    section9 = client_response.get("section9") or (payload or {}).get("section9", {})
    missing = [k for k in SECTION9_KEYS if not section9.get(k)]
    if missing:
        raise ValueError(f"Не хватает полей Раздела 9: {missing}")
    return section9


# ---------------------------------------------------------------------------
# 3. Цена/длительность Программы для конкретной Стадии — парсим реальную
#    структуру data["consulting_programs"]["individual_programs"][kse].
# ---------------------------------------------------------------------------

def _price_for_stage(program, stage_id):
    price = program["цена"]
    if price["тип"] == "плоская":
        return price["значение"]
    for t in price["тарифы"]:
        if stage_id in t["стадии"]:
            return t["значение"]
    raise ValueError(f"Нет тарифа для Стадии {stage_id} в {program['название_программы']}")


def _weeks_for_stage(program, stage_id):
    """'13 недель (Стадии 1-3) / 15 недель (Стадии 4-5) и далее' -> число недель
    для конкретной Стадии (ищем сегмент, диапазон которого содержит stage_id).
    Если строка длительности не в этом формате (например, "12 месяцев" или
    "13-15 недель в зависимости от Стадии" — оба варианта реально встречаются
    в data/consulting_programs.json) — НЕ гадаем число, возвращаем None, и
    adapt() ниже подставит исходную строку как есть, без искажения."""
    duration = program["длительность"]
    segments = re.findall(r'(\d+)\s*недел\w*\s*\(Стадии\s*([\d\-,\s]+)\)', duration)
    for weeks_str, stages_str in segments:
        stage_nums = set()
        for part in stages_str.split(","):
            part = part.strip()
            if "-" in part:
                lo, hi = part.split("-")
                stage_nums.update(range(int(lo), int(hi) + 1))
            elif part:
                stage_nums.add(int(part))
        if stage_id in stage_nums:
            return int(weeks_str)
    return None


def format_rub(amount):
    return f'{amount:,.0f}'.replace(',', ' ') + ' руб.'


def _weeks_word(n):
    """Русское склонение 'неделя/недели/недель' по числу n."""
    n_abs = abs(n) % 100
    n1 = n_abs % 10
    if 11 <= n_abs <= 14:
        return "недель"
    if n1 == 1:
        return "неделя"
    if 2 <= n1 <= 4:
        return "недели"
    return "недель"


# ---------------------------------------------------------------------------
# 4. Основной адаптер
# ---------------------------------------------------------------------------

def adapt(diagnose_result, client_response, data, payload=None, top_n_part2=2):
    """payload теперь НЕОБЯЗАТЕЛЕН — с исправленным build_client_response() в
    app.py все нужные поля (psychographic/urgency/decisionMaker/section9) уже
    лежат в client_response. payload оставлен только как fallback на случай
    более старого client_response, собранного до этого исправления."""
    stage = diagnose_result["стадия"]
    stage_id = stage["stage_id"]
    stage_name_full = f'Стадия {stage_id}: {stage["stage_name"]}'

    flow_a = diagnose_result["поток_а"]
    flow_a_mismatches = flow_a["mismatches"]

    q = client_response["qualification"]
    qualification = {
        "name": q["name"],
        "company": q["company"],
        "diagnosis_date": q["diagnosis_date"],
        "report_number": q["report_number"],
        "urgency": q.get("urgency", ""),
        "decisionMaker": q.get("decisionMaker", ""),
    }

    psychographic_ranks = get_psychographic_ranks(client_response, payload)
    section9 = get_section9(client_response, payload)

    mapping = data["mapping_kse"]
    client_challenge_scores = client_response["challenge_scores"]
    area_deficit = flow_a["area_deficit"]

    # --- КСЭ, отсортированные по ярусам (фундамент -> ядро -> надстройка),
    # без "низкий_приоритет" — та же группировка, что в report_text_generator
    # (render_section11_kse_list/render_section12_2_plan), чтобы Дом на
    # слайде 14 и в Скрипте совпадал 1-в-1 с Отчётом. ---
    rows = [r for r in diagnose_result["приоритизация_ксэ"] if r["ярус"] != "низкий_приоритет"]
    grouped = {"фундамент": [], "ядро": [], "надстройка": []}
    for row in rows:
        grouped[_normalize_tier(row["ярус"])].append(row)

    all_kse_ordered = [row["kse"] for tier in ("фундамент", "ядро", "надстройка")
                        for row in grouped[tier]]

    # top_kse_rows — Часть 2 Раздела 4: до top_n_part2 приоритетных КСЭ из
    # ОСТАЛЬНЫХ 10 (не "Критерии роста бизнеса" — та обрабатывается в Части 1).
    top_kse_rows = [k for k in all_kse_ordered if k != "Критерии роста бизнеса"][:top_n_part2]

    # --- tier_programs: {"Фундамент": [...], "Ядро": [...], "Надстройка": [...]}
    # — реальные цены/сроки из data["consulting_programs"] для ЭТОЙ Стадии. ---
    programs_data = data["consulting_programs"]["individual_programs"]
    tier_programs = {}
    for tier_key in ("фундамент", "ядро", "надстройка"):
        tier_rows = grouped[tier_key]
        if not tier_rows:
            continue
        out = []
        for row in tier_rows:
            kse = row["kse"]
            program = programs_data[kse]
            outcomes = program["что_компания_получает"][:3]
            outcomes_text = "; ".join(o[0].lower() + o[1:] for o in outcomes)
            price = _price_for_stage(program, stage_id)
            weeks = _weeks_for_stage(program, stage_id)
            out.append({
                "name": kse,
                "outcomes": outcomes_text,
                "duration": f"{weeks} {_weeks_word(weeks)}" if weeks else program["длительность"],
                "price": format_rub(price),
            })
        tier_programs[TIER_KEY_TO_RU_TITLE[tier_key]] = out

    # --- bundle_fork: развилка "Путь А / Путь Б" (Комплексная программа
    # «Возрождение малого бизнеса», Стадии 1-3). ДО этого исправления (22.08.2026)
    # этот расчёт в Скрипте отсутствовал вообще — render_solution_section()
    # вызывалась без bundle_fork, падая на bundle_fork=None по умолчанию, из-за
    # чего блок пропускался ВСЕГДА, для любого клиента, даже когда пересечение
    # КСЭ было достаточным и в Презентации слайд 24 корректно строился. Логика
    # здесь зеркальна pptx_data_export.py:_build_bundle_fork — тот же источник
    # data["consulting_programs"]["bundle_programs"], тот же compute_bundle_fork(). ---
    bundle_fork = None
    bundle = data["consulting_programs"]["bundle_programs"].get("Возрождение малого бизнеса")
    if bundle and stage_id in bundle.get("ограничение_по_стадиям", [1, 2, 3]):
        bundle_kse_covered = set(bundle["ксэ_покрывает"])
        individual_weeks, individual_prices = {}, {}
        for kse in all_kse_ordered:
            prog = programs_data.get(kse)
            if not prog:
                continue
            individual_prices[kse] = _price_for_stage(prog, stage_id)
            weeks = _weeks_for_stage(prog, stage_id)
            if weeks:
                individual_weeks[kse] = weeks
        bundle_weeks_nums = re.findall(r'\d+', bundle["длительность"])
        bundle_weeks = int(bundle_weeks_nums[0]) if bundle_weeks_nums else None
        bundle_fork = compute_bundle_fork(
            all_kse_ordered=all_kse_ordered, stage_id=stage_id,
            bundle_kse_covered=bundle_kse_covered,
            bundle_weeks=bundle_weeks,
            bundle_price=bundle["цена"]["значение"],
            individual_program_weeks=individual_weeks,
            individual_program_prices=individual_prices,
        )

    return {
        "qualification": qualification,
        "stage_name": stage_name_full,
        "psychographic_ranks": psychographic_ranks,
        "section9": section9,
        "flow_a_mismatches": flow_a_mismatches,
        "top_kse_rows": top_kse_rows,
        "client_challenge_scores": client_challenge_scores,
        "area_deficit": area_deficit,
        "mapping": mapping,
        "flow_a": flow_a,
        "all_kse_ordered": all_kse_ordered,
        "tier_programs": tier_programs,
        "flow_b": diagnose_result["поток_б"],
        "stage_id": stage_id,
        "bundle_fork": bundle_fork,
    }
