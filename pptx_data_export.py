# -*- coding: utf-8 -*-
"""
pptx_data_export.py — переводит РЕАЛЬНЫЙ результат диагностики
(diagnose_result + client_response, тот же вход, что у script_adapter.py)
в единый JSON, из которого должна собираться Презентация (pptx_build/*.js).

ЭТО ШАГ 1 из 2 автоматизации Презентации:
  1. (этот файл) Python считает всё нужное для слайдов и пишет JSON.
  2. (следующий шаг, ещё не сделан) base.js должен НАЧАТЬ ЧИТАТЬ этот JSON
     вместо своих текущих захардкоженных demo-констант (client, section9,
     growthRules, evidenceChain, maturityTop5, tierPrograms, loyaltyProgram)
     — а также два слайда, у которых демо-данные захардкожены ЛОКАЛЬНО, в
     обход base.js вообще (new_13_kselist.js: priorityForThisClient;
     new_23_bundle_stub.js: demoIntersection/bundlePrograms/demoBundle*) —
     их тоже нужно будет переключить на этот JSON.

Все ключи в JSON — ровно те же имена, что сейчас экспортирует base.js, чтобы
переключение JS на JSON было максимально мéханическим (переименовывать
ничего не потребуется).
"""
import json
import re
from pathlib import Path

import scoring_algorithm as sa
from script_adapter import adapt as adapt_script_data
from consultation_script_builder import (
    get_primary_secondary_type, compute_bundle_fork,
    MATURITY_AHEAD, MATURITY_BEHIND, MATURITY_MIXED, MATURITY_MATCH,
)

BASE = Path(__file__).parent

GROWTH_RULE_NAMES = (
    "Приоритетные сферы",
    "Коэффициент Строитель-Протектор",
    "Модальность",
    "Стили управления",
    "Три роли лидера",
    "Непреложные правила",
)

TIER_RU = {"фундамент": "Фундамент", "ядро": "Ядро", "надстройка": "Надстройка"}

LOYALTY_DISCOUNT_SCHEDULE = [0, 10, 15, 20, 25]


def _build_growth_rules(flow_a):
    """[{name, ok}] x6 — для слайда 5. mismatches из flow_a может содержать
    'Непреложные правила (минимум одно < 80%)' — сверяем по началу строки,
    не по точному совпадению, т.к. формулировка чуть длиннее канонической."""
    mismatches = flow_a["mismatches"]
    return [
        {"name": name, "ok": not any(m.startswith(name) for m in mismatches)}
        for name in GROWTH_RULE_NAMES
    ]


def _build_evidence_chain(flow_a, flow_b, area_deficit):
    growth_rule_mismatches = [m.split(" (")[0] for m in flow_a["mismatches"]
                               if m.split(" (")[0] != "Непреложные правила"]
    challenges = flow_b["individual_top5"][:3]
    areas = [a for a, deficit in area_deficit.items() if deficit]
    return {
        "growthRuleMismatches": growth_rule_mismatches,
        "challenges": challenges,
        "areas": areas,
    }


def _build_maturity_top5(flow_b, stage_id, primary_type, data):
    top5_map = json.load(open(BASE / "data" / "classic_challenges_top5_by_stage.json", encoding="utf-8"))
    maturity = flow_b["maturity"]
    later = set(maturity["later_stage_challenges"])
    earlier = set(maturity["earlier_stage_challenges"])

    challenges = []
    for c in flow_b["individual_top5"]:
        typical_stages = top5_map.get(c, [])
        flagged = c in later or c in earlier
        stage_text = ("Стадия " if len(typical_stages) == 1 else "Стадии ") + \
                     ", ".join(str(s) for s in typical_stages) if typical_stages else "—"
        challenges.append({"text": c, "typicalStage": stage_text, "flagged": flagged})

    note = maturity["note"]
    if note == "опережает_стадию":
        conclusion = MATURITY_AHEAD[primary_type].split("\n")[-1]  # последний абзац — сам вывод
    elif note == "отстаёт_от_стадии":
        conclusion = MATURITY_BEHIND[primary_type].split("\n")[-1]
    elif note == "смешанная_картина":
        conclusion = MATURITY_MIXED[primary_type].split("\n")[-1]
    else:
        conclusion = MATURITY_MATCH[primary_type].split("\n")[-1]

    return {"challenges": challenges, "conclusion": conclusion}


def _build_loyalty_program(all_kse_ordered, tier_programs):
    price_by_kse = {p["name"]: int("".join(ch for ch in p["price"] if ch.isdigit()))
                     for programs in tier_programs.values() for p in programs}
    items = []
    for i, kse in enumerate(all_kse_ordered):
        price = price_by_kse.get(kse)
        if price is None:
            continue
        discount = LOYALTY_DISCOUNT_SCHEDULE[min(i, len(LOYALTY_DISCOUNT_SCHEDULE) - 1)]
        items.append({"name": kse, "price": price, "discount": discount})
    return {"items": items}


def _build_bundle_fork(result, data, stage_id, all_kse_ordered):
    consulting = data["consulting_programs"]
    bundle = consulting["bundle_programs"].get("Возрождение малого бизнеса")
    if not bundle or stage_id not in bundle.get("ограничение_по_стадиям", [1, 2, 3]):
        return None
    bundle_kse_covered = set(bundle["ксэ_покрывает"])
    individual_weeks, individual_prices = {}, {}
    for kse in all_kse_ordered:
        prog = consulting["individual_programs"].get(kse)
        if not prog:
            continue
        price = prog["цена"]
        if price["тип"] == "плоская":
            individual_prices[kse] = price["значение"]
        else:
            for t in price["тарифы"]:
                if stage_id in t["стадии"]:
                    individual_prices[kse] = t["значение"]
        segs = re.findall(r'(\d+)\s*недел\w*\s*\(Стадии\s*([\d\-,\s]+)\)', prog["длительность"])
        for weeks_str, stages_str in segs:
            nums = set()
            for part in stages_str.split(","):
                part = part.strip()
                if "-" in part:
                    lo, hi = part.split("-")
                    nums.update(range(int(lo), int(hi) + 1))
                elif part:
                    nums.add(int(part))
            if stage_id in nums:
                individual_weeks[kse] = int(weeks_str)

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
    if bundle_fork is None:
        return None
    return {
        "intersection": bundle_fork["intersection"],
        "bundlePrograms": sorted(bundle_kse_covered),
        "individualWeeks": bundle_fork["individual_weeks_sum"],
        "individualPrice": bundle_fork["individual_price_sum"],
        "bundleWeeks": bundle_fork["bundle_weeks"],
        "bundlePrice": bundle_fork["bundle_price"],
    }


def _tier_programs_for_slides(tier_programs):
    """Слайды (new_19/20/22_*.js) ждут outcomes как СПИСОК строк (для
    буллетов) — а script_adapter.py собирает их в одну строку через '; '
    (для прозы в Отчёте/Скрипте). Разбиваем обратно для Презентации."""
    out = {}
    for tier, programs in tier_programs.items():
        out[tier] = [
            {**p, "outcomes": [o.strip() for o in p["outcomes"].split(";") if o.strip()]}
            for p in programs
        ]
    return out


def export_pptx_data(diagnose_result, client_response, data, payload=None, top_n_slide13=5):
    """Главная функция — вызывать с теми же аргументами, что и
    script_adapter.adapt(). Возвращает dict, готовый к json.dump()."""
    adapted = adapt_script_data(diagnose_result, client_response, data, payload)
    primary_type, _ = get_primary_secondary_type(adapted["psychographic_ranks"])

    flow_a = adapted["flow_a"]
    flow_b = diagnose_result["поток_б"]
    stage = diagnose_result["стадия"]
    stage_id = stage["stage_id"]

    q = client_response["qualification"]

    return {
        "client": {
            "name": q["name"],
            "company": q["company"],
            "stageId": stage_id,
            "stageName": stage["stage_name"],
            "date": q["diagnosis_date"],
        },
        "section9": adapted["section9"],
        "growthRules": _build_growth_rules(flow_a),
        "evidenceChain": _build_evidence_chain(flow_a, flow_b, adapted["area_deficit"]),
        "maturityTop5": _build_maturity_top5(flow_b, stage_id, primary_type, data),
        "allKseOrdered": adapted["all_kse_ordered"],
        "topKseForSlide13": adapted["all_kse_ordered"][:top_n_slide13],
        "tierPrograms": _tier_programs_for_slides(adapted["tier_programs"]),
        "loyaltyProgram": _build_loyalty_program(adapted["all_kse_ordered"], adapted["tier_programs"]),
        "bundleFork": _build_bundle_fork(diagnose_result, data, stage_id, adapted["all_kse_ordered"]),
    }


if __name__ == "__main__":
    # Тот же демо-payload, что и в run_real_pipeline_test.py — самостоятельно,
    # без exec() чужого файла (там был __file__, ломающий такой трюк).
    from app import build_client_response, _load_statements

    data = sa.load_data()
    with open(BASE / "data" / "stage_level_report_texts.json", encoding="utf-8") as f:
        data["stage_level_report_texts"] = json.load(f)
    with open(BASE / "data" / "consulting_programs.json", encoding="utf-8") as f:
        data["consulting_programs"] = json.load(f)
    statements = _load_statements()
    mapping_kse = json.load(open(BASE / "data" / "mapping_kse.json", encoding="utf-8"))
    immutable_rules = json.load(open(BASE / "data" / "immutable_rules.json", encoding="utf-8"))
    challenges = list(mapping_kse["challenge_to_kse"].keys())
    areas_stage3 = immutable_rules["3"]

    priority_texts = list(statements["priority_spheres_map"].keys())
    section1 = {t: (i % 3) + 1 for i, t in enumerate(priority_texts)}
    section2 = {c: 2 for c in challenges}
    section2.update({"Текучесть кадров": 9, "Разрыв между руководством и сотрудниками": 8,
                      "Не внедряются нужные системы": 7})
    section3 = "Скорее уверенное"
    levels_cycle = ["Руководство", "Менеджеры", "Сотрудники"]
    section4 = [levels_cycle[i % 3] for i in range(9)]
    section5 = {}
    for group_num, texts in statements["management_styles_statements"].items():
        section5[group_num] = {t: (6 - idx) for idx, t in enumerate(texts)}
    section6 = {"visionaire": 7, "manager": 5, "specialist": 3}
    section7 = {}
    for area, rules in areas_stage3.items():
        for idx in range(len(rules)):
            low = area in ("Управление", "Производственные отношения")
            section7[f"{area}::{idx}"] = 20 if low else 100
    section8 = {t: 3 for t in statements["section8_statements"]}
    qualification_raw = {
        "name": "Виктор Громов", "company": "ООО «Прибрежный Дом»",
        "email": "viktor.gromov@example.com", "phone": "+7 (999) 123-45-67",
        "fte_a": 15, "fte_b": 5, "fte_c": 8, "fte_d": 25,
        "managers": 2, "leaders": 1, "employeesYearAgo": 20, "timeYears": 4, "timeMonths": 0,
        "businessAge": 6,
        "psychographic": {"Мыслитель": 1, "Наблюдатель": 2, "Делатель": 3, "Чувствователь": 4},
        "urgency": "Уже сейчас, готов(а) действовать", "decisionMaker": "Я лично, единолично",
    }
    section9_raw = {
        "problem1": "Менеджеры среднего звена постоянно уходят",
        "problem2": "Я лично согласовываю почти каждое решение",
        "problem3": "Команда не понимает, куда мы движемся",
        "whyTheseChallenges": "Рос на интуиции",
        "whyCantSolve": "Не хватает времени",
        "costOfInaction": "Выгорю окончательно",
        "priceOfInaction": "Потерянная прибыль минимум 3-4 млн в год",
        "dreamOutcome": "Компания работает без меня",
    }
    payload = {"qualification": qualification_raw, "section1": section1, "section2": section2,
               "section3": section3, "section4": section4, "section5": section5,
               "section6": section6, "section7": section7, "section8": section8,
               "section9": section9_raw}

    client_response, q = build_client_response(payload, data, statements)
    result = sa.diagnose(client_response, data)

    pptx_data = export_pptx_data(result, client_response, data, payload)

    out_path = BASE / "pptx_data.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(pptx_data, f, ensure_ascii=False, indent=2)
    print(f"Сохранено: {out_path}")
    print(f"Ключи: {list(pptx_data.keys())}")
