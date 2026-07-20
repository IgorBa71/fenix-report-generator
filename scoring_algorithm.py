"""
Скоринг-алгоритм «Полная оценка состояния бизнеса»
=====================================================

Реализует логику из документа «Логика алгоритма» поверх структурированных
данных в /data. Точка входа — функция diagnose(client_response).

ВАЖНОЕ ОГРАНИЧЕНИЕ (см. итог в конце файла и финальный ответ в чате):
Для 6 измерений Потока А (Приоритетные сферы, Строитель-Протектор, Модальность,
Стили управления, Три роли лидера, Непреложные правила) этот модуль принимает
УЖЕ ГОТОВЫЕ значения клиента как вход. Формула «сырые ответы опроса разделов
1,3,4,5,6 -> итоговое значение измерения» (например, как из 3 групп ранжирования
статей получить порядок Прибыль/Люди/Процессы) в проекте пока не описана —
это отдельная задача, см. раздел "ОТКРЫТЫЙ ВОПРОС" внизу файла.
"""

import json
import os
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"


# ---------------------------------------------------------------------------
# 0. Загрузка данных
# ---------------------------------------------------------------------------

def load_data():
    files = [
        "stages", "rules_of_growth_targets", "mapping_kse",
        "classic_challenges_top5_by_stage", "immutable_rules",
        "kse_descriptions", "kse_priority_by_stage", "challenge_symptoms_by_stage",
        "consulting_programs", "zone_descriptions",
    ]
    data = {}
    for name in files:
        with open(DATA_DIR / f"{name}.json", encoding="utf-8") as f:
            data[name] = json.load(f)
    return data


# ---------------------------------------------------------------------------
# 1. Стадия, Зона, % прохождения
# ---------------------------------------------------------------------------

def compute_fte(fte_a, fte_b, fte_c, fte_d):
    """E = a + b + (c * d / 40), округление до целого."""
    return round(fte_a + fte_b + (fte_c * fte_d / 40))


def determine_stage(e, data):
    for stage in data["stages"]["stages"]:
        lo, hi = stage["fte_range"]
        if lo <= e <= hi:
            return stage
    # За пределами диапазона (>350) — считаем Стадией 7 без верхней границы
    if e > data["stages"]["stages"][-1]["fte_range"][1]:
        return data["stages"]["stages"][-1]
    return data["stages"]["stages"][0]


def determine_zone(e, stage):
    """Возвращает (тип_зоны, процент_прохождения_стадии)."""
    lo, hi = stage["fte_range"]
    pct = round(100 * (e - lo) / (hi - lo), 1) if hi > lo else 100.0

    if stage["zone_entry"] and e <= stage["zone_entry"]["range"][1]:
        zone_name, zone_type = "Зона входа", stage["zone_entry"]["type"]
    elif stage["zone_exit"] and e >= stage["zone_exit"]["range"][0]:
        zone_name, zone_type = "Зона выхода", stage["zone_exit"]["type"]
    else:
        zone_name, zone_type = "Функциональная зона", None

    return {"zone_name": zone_name, "zone_type": zone_type, "percent_through_stage": pct}


def calculate_stage_zone(fte_a, fte_b, fte_c, fte_d, data):
    e = compute_fte(fte_a, fte_b, fte_c, fte_d)
    stage = determine_stage(e, data)
    zone = determine_zone(e, stage)
    return {
        "fte": e,
        "stage_id": stage["id"],
        "stage_name": stage["name"],
        **zone,
    }


# ---------------------------------------------------------------------------
# 2. Поток А — Правила роста (+ кадровый разрыв Менеджеры/Руководители)
# ---------------------------------------------------------------------------

def evaluate_flow_a(stage_id, client_dimensions, immutable_rules_pct,
                     managers_actual, leaders_actual, data):
    """
    client_dimensions: словарь фактических значений клиента по 5 измерениям
        (кроме Непреложных правил, которые идут отдельно):
        {
          "priority_spheres": ["Прибыль","Люди","Процессы"],   # ранг клиента
          "builder_protector_ratio": "3:1",
          "modality": {"Руководство": "...", "Менеджеры": "...", "Сотрудники": "..."},
          "management_styles": {"Основной": "...", "Второстепенный": "...", "Дополнительный": "..."},
          "three_leader_roles": {"Визионер": 30, "Менеджер": 40, "Специалист": 30},
        }
    immutable_rules_pct: {area_name: [pct_rule1, pct_rule2, ...]}  # тот же
        порядок, что в immutable_rules.json для данной Стадии
    """
    targets = data["rules_of_growth_targets"]
    s = str(stage_id)
    mismatches = []

    if client_dimensions.get("priority_spheres") != targets["priority_spheres"].get(s):
        mismatches.append("Приоритетные сферы")
    if client_dimensions.get("builder_protector_ratio") != targets["builder_protector_ratio"].get(s):
        mismatches.append("Коэффициент Строитель-Протектор")
    if client_dimensions.get("modality") != targets["modality"].get(s):
        mismatches.append("Модальность")
    if client_dimensions.get("management_styles") != targets["management_styles"].get(s):
        mismatches.append("Стили управления")
    if client_dimensions.get("three_leader_roles") != targets["three_leader_roles"].get(s):
        mismatches.append("Три роли лидера")

    # Непреложные правила — % по каждому правилу, порог 80%
    rules_struct = data["immutable_rules"][s]
    failed_rules = []  # (область, текст_правила, факт%)
    area_deficit = {area: 0 for area in rules_struct}
    for area, rules in rules_struct.items():
        pct_list = immutable_rules_pct.get(area, [])
        for i, rule in enumerate(rules):
            pct = pct_list[i] if i < len(pct_list) else 0
            deficit = max(0, 80 - pct)
            area_deficit[area] += deficit
            if pct < 80:
                failed_rules.append({
                    "область": area,
                    "правило": rule["текст"],
                    "стадия_появления": rule["стадия_появления"],
                    "факт_%": pct,
                })

    if failed_rules:
        mismatches.append("Непреложные правила (минимум одно < 80%)")

    # --- триггер: КСЭ «Критерии роста бизнеса» ---
    # порог значимости: >=1 расхождение из 5 измерений ИЛИ любое правило <80%
    trigger_kriterii_rosta = (len(mismatches) >= 1)

    # --- 2.1 Кадровый разрыв Менеджеры/Руководители ---
    stage = next(st for st in data["stages"]["stages"] if st["id"] == stage_id)
    mgr_gap = max(0, stage["managers_range"][0] - managers_actual)
    ldr_gap = max(0, stage["leaders_range"][0] - leaders_actual)
    staffing_gap = mgr_gap > 0 or ldr_gap > 0

    trigger_org_structure = staffing_gap
    trigger_strong_mgmt_team = staffing_gap and stage_id >= 4

    return {
        "mismatches": mismatches,
        "mismatch_count": len(mismatches),
        "failed_immutable_rules": failed_rules,
        "area_deficit": area_deficit,  # используется дальше в Потоке В
        "managers_gap": mgr_gap,
        "leaders_gap": ldr_gap,
        "auto_triggers": {
            "Критерии роста бизнеса": trigger_kriterii_rosta,
            "Организационная структура": trigger_org_structure,
            "Сильная Управленческая команда": trigger_strong_mgmt_team,
        },
    }


# ---------------------------------------------------------------------------
# 3. Поток Б — Симптомы (24 Вызова) -> КСЭ + Уровень зрелости бизнеса
# ---------------------------------------------------------------------------

def evaluate_flow_b(stage_id, challenge_scores, data):
    """challenge_scores: {challenge_name: 0..10}"""
    mapping = data["mapping_kse"]["challenge_to_kse"]
    kse_list = data["mapping_kse"]["kse_list"]

    raw_signal = {kse: 0 for kse in kse_list}
    max_signal = {kse: 0 for kse in kse_list}
    for challenge, score in challenge_scores.items():
        for kse in mapping.get(challenge, []):
            raw_signal[kse] += score
            max_signal[kse] += 10

    signal_norm = {
        kse: round(100 * raw_signal[kse] / max_signal[kse], 1) if max_signal[kse] else 0.0
        for kse in kse_list
    }

    # --- Уровень зрелости бизнеса ---
    individual_top5 = sorted(challenge_scores, key=challenge_scores.get, reverse=True)[:5]
    top5_map = data["classic_challenges_top5_by_stage"]

    later_stage_hits, earlier_stage_hits = [], []
    for ch in individual_top5:
        stages_for_ch = top5_map.get(ch, [])
        if stage_id in stages_for_ch:
            continue  # вызов типичен для текущей Стадии клиента — не сигнал зрелости,
                      # даже если он также типичен и для других Стадий
        if any(s > stage_id for s in stages_for_ch):
            later_stage_hits.append(ch)
        if any(s < stage_id for s in stages_for_ch):
            earlier_stage_hits.append(ch)

    if later_stage_hits and not earlier_stage_hits:
        maturity_note = "опережает_стадию"
    elif earlier_stage_hits and not later_stage_hits:
        maturity_note = "отстаёт_от_стадии"
    elif later_stage_hits and earlier_stage_hits:
        maturity_note = "смешанная_картина"
    else:
        maturity_note = "соответствует_стадии"

    return {
        "raw_signal": raw_signal,
        "signal_norm": signal_norm,
        "individual_top5": individual_top5,
        "maturity": {
            "note": maturity_note,
            "later_stage_challenges": later_stage_hits,
            "earlier_stage_challenges": earlier_stage_hits,
        },
    }


# ---------------------------------------------------------------------------
# 4. Поток В — Области бизнеса -> КСЭ (пропорциональный бонус)
# ---------------------------------------------------------------------------

def evaluate_flow_c(area_deficit, data):
    """area_deficit приходит из evaluate_flow_a (Дефицит(Область), 0..80*N)."""
    mapping = data["mapping_kse"]["business_area_to_kse"]
    kse_list = data["mapping_kse"]["kse_list"]

    raw_bonus = {kse: 0 for kse in kse_list}
    max_bonus = {kse: 0 for kse in kse_list}

    # обратный маппинг: КСЭ -> список областей, которые на него ссылаются
    kse_to_areas = {kse: [] for kse in kse_list}
    for area, kses in mapping.items():
        for kse in kses:
            kse_to_areas[kse].append(area)

    for kse, areas in kse_to_areas.items():
        for area in areas:
            raw_bonus[kse] += area_deficit.get(area, 0)
            # максимум по области — если бы все правила в ней были на 0%
            # (это не хранится отдельно, поэтому берём area_deficit_max снаружи)

    return {"raw_bonus": raw_bonus, "kse_to_areas": kse_to_areas}


def compute_area_deficit_max(stage_id, data):
    """Максимально возможный Дефицит(Область) = кол-во правил * 80."""
    rules_struct = data["immutable_rules"][str(stage_id)]
    return {area: len(rules) * 80 for area, rules in rules_struct.items()}


# ---------------------------------------------------------------------------
# 5. Итоговая ярусная приоритизация КСЭ
# ---------------------------------------------------------------------------

def compute_priority_list(stage_id, flow_a, flow_b, flow_c, area_deficit_max, data):
    kse_list = data["mapping_kse"]["kse_list"]
    tiers = data["kse_priority_by_stage"]["kse_priority_by_stage"]
    s = str(stage_id)

    # нормализация бонуса Потока В к 0..100 на КСЭ
    kse_to_areas = flow_c["kse_to_areas"]
    bonus_norm = {}
    for kse in kse_list:
        areas = kse_to_areas[kse]
        max_possible = sum(area_deficit_max.get(a, 0) for a in areas)
        raw = flow_c["raw_bonus"][kse]
        bonus_norm[kse] = round(100 * raw / max_possible, 1) if max_possible else 0.0

    results = []
    for kse in kse_list:
        tier_info = tiers[kse][s]
        tier = tier_info["ярус"]

        # автотриггеры поднимают КСЭ до яруса "фундамент", если он ещё не там
        if flow_a["auto_triggers"].get(kse):
            if tier != "фундамент":
                tier = "фундамент (авто-триггер Потока А)"

        score = flow_b["signal_norm"][kse] + bonus_norm[kse]
        results.append({
            "kse": kse,
            "категория": data["kse_descriptions"][kse]["категория"],
            "маркер_стадии": tier_info["маркер"],
            "ярус": tier,
            "симптом_сигнал": flow_b["signal_norm"][kse],
            "бонус_областей": bonus_norm[kse],
            "скор": round(score, 1),
        })

    tier_rank = {
        "фундамент (авто-триггер Потока А)": 0,
        "фундамент": 0,
        "ядро": 1,
        "надстройка": 2,
        "низкий_приоритет": 3,
    }
    results.sort(key=lambda r: (tier_rank[r["ярус"]], -r["скор"]))
    return results


# ---------------------------------------------------------------------------
# 6. Раздел 8 — Лёгкость продажи (не влияет на приоритизацию)
# ---------------------------------------------------------------------------

def evaluate_sellability(section8_likert_by_kse):
    """section8_likert_by_kse: {kse_name: [likert_1_6, ...]} -> тег."""
    result = {}
    for kse, values in section8_likert_by_kse.items():
        avg = sum(values) / len(values) if values else 0
        if avg >= 5:
            tag = "Высокая"
        elif avg >= 3:
            tag = "Средняя"
        else:
            tag = "Низкая"
        result[kse] = {"среднее": round(avg, 2), "лёгкость_продажи": tag}
    return result


# ---------------------------------------------------------------------------
# 7. Оркестратор
# ---------------------------------------------------------------------------

def diagnose(client_response, data=None):
    if data is None:
        data = load_data()

    q = client_response["qualification"]
    stage_zone = calculate_stage_zone(q["fte_a"], q["fte_b"], q["fte_c"], q["fte_d"], data)
    stage_id = stage_zone["stage_id"]

    flow_a = evaluate_flow_a(
        stage_id,
        client_response["flow_a_dimensions"],
        client_response["immutable_rules_pct"],
        q["managers_actual"], q["leaders_actual"],
        data,
    )
    flow_b = evaluate_flow_b(stage_id, client_response["challenge_scores"], data)

    area_deficit_max = compute_area_deficit_max(stage_id, data)
    flow_c = evaluate_flow_c(flow_a["area_deficit"], data)

    priority_list = compute_priority_list(stage_id, flow_a, flow_b, flow_c, area_deficit_max, data)
    sellability = evaluate_sellability(client_response.get("section8_likert_by_kse", {}))

    return {
        "стадия": stage_zone,
        "поток_а": flow_a,
        "поток_б": {k: v for k, v in flow_b.items() if k != "raw_signal"},
        "приоритизация_ксэ": priority_list,
        "лёгкость_продажи": sellability,
    }


# ---------------------------------------------------------------------------
# Демонстрационный прогон на синтетических данных (Стадия 1)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    data = load_data()

    # синтетический клиент: 8 сотрудников, штат, без частичной занятости
    # намеренно НЕ выполняет Непреложные правила и не соответствует Модальности,
    # чтобы продемонстрировать срабатывание автотриггеров
    challenges = list(data["classic_challenges_top5_by_stage"].keys())
    demo_scores = {ch: 3 for ch in challenges}
    # искусственно поднимем несколько вызовов, связанных с Организационной структурой
    for ch in ["Внутренний хаос расшатывает бизнес", "Текучесть кадров", "Найм качественных сотрудников"]:
        demo_scores[ch] = 9

    demo_client = {
        "qualification": {
            "fte_a": 8, "fte_b": 0, "fte_c": 0, "fte_d": 0,
            "managers_actual": 0, "leaders_actual": 1,
        },
        "flow_a_dimensions": {
            "priority_spheres": ["Люди", "Прибыль", "Процессы"],  # не совпадает с целевым для Стадии 1
            "builder_protector_ratio": "1:1",  # не совпадает с целевым 4:1
            "modality": {"Руководство": "Доминирующая", "Менеджеры": "Поддерживающая", "Сотрудники": "Вспомогательная"},
            "management_styles": {"Основной": "Авторитетный", "Второстепенный": "Коучинговый", "Дополнительный": "Директивный"},
            "three_leader_roles": {"Визионер": 40, "Менеджер": 10, "Специалист": 50},
        },
        "immutable_rules_pct": {
            "Развитие бизнеса": [50, 60],
            "Бизнес-модель и планирование": [10, 40],
            "Финансы": [70, 50, 60],
            "Управление": [80, 40],
            "Операционная деятельность": [60],
            "Производственные отношения": [20],
        },
        "challenge_scores": demo_scores,
        "section8_likert_by_kse": {
            "Организационная структура": [5, 6],
            "Критерии роста бизнеса": [2, 3],
        },
    }

    result = diagnose(demo_client, data)

    print("=== СТАДИЯ ===")
    print(result["стадия"])
    print("\n=== ПОТОК А ===")
    print("Расхождения:", result["поток_а"]["mismatches"])
    print("Автотриггеры:", result["поток_а"]["auto_triggers"])
    print("Кадровый разрыв — менеджеры:", result["поток_а"]["managers_gap"],
          "| руководители:", result["поток_а"]["leaders_gap"])
    print("\n=== ЗРЕЛОСТЬ БИЗНЕСА (Поток Б) ===")
    print(result["поток_б"]["maturity"])
    print("Индивидуальный Топ-5:", result["поток_б"]["individual_top5"])
    print("\n=== ПРИОРИТИЗАЦИЯ КСЭ (топ-6) ===")
    for row in result["приоритизация_ксэ"][:6]:
        print(f"  [{row['ярус']:35s}] {row['kse']:40s} скор={row['скор']}")
    print("\n=== ЛЁГКОСТЬ ПРОДАЖИ ===")
    print(result["лёгкость_продажи"])
