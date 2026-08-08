"""
Генерация блока «ВЫВОДЫ» Раздела 1 «Обзор компании».

Финальный шаблон — 6 частей, из них 2 всегда показываются (Стадия+Зона,
закрывающая строка), 1 показывается по ветке (несоответствия Потока А:
есть/нет), 3 условные (Кадровый разрыв, тизер Зрелости, тизер Критических
упущений — только если применимо).

Источники:
- Части 1, 3 (обе ветки), 6 — из официального шаблона Игоря (Стадии 2-7,
  Стадия 1 — из примера отчёта), решение сохранить упоминание программы
  «Критерии роста бизнеса» в Части 3 — подтверждено (детерминированная связь
  с автотриггером Потока А, не «спойлер» персонализированной рекомендации).
- Части 2, 4, 5 — новые блоки, добавленные в ходе проектирования (п.2.1
  документа логики; Раздел 8 и Раздел 9 соответственно).
"""

import json
import re
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"

MEASUREMENT_DISPLAY_NAMES = {
    "Приоритетные сферы": "Приоритетные сферы",
    "Коэффициент Строитель-Протектор": "Коэффициент Строитель-Протектор",
    "Модальность": "Модальность",
    "Стили управления": "Стили управления",
    "Три роли лидера": "Три роли лидера",
    "Непреложные правила (минимум одно < 80%)": "Непреложные правила",
}

# Родительный падеж названий Стадий (для конструкции "Стадия X, которая называется Стадия {X в род. падеже}")
STAGE_NAME_GENITIVE = {
    1: "Стартапа",
    2: "Подъёма",
    3: "Делегирования",
    4: "Профессиональная",  # прилагательное, падеж не меняется
    5: "Интеграции",
    6: "Стратегическая",  # прилагательное, падеж не меняется
    7: "Образа будущего",
}

# Родительный падеж названий Зон (для конструкции "через зону {тип в род. падеже}")
ZONE_TYPE_GENITIVE = {
    "Наводнение": "Наводнения",
    "Аэротруба": "Аэротрубы",
}


def load_zone_descriptions():
    with open(DATA_DIR / "zone_descriptions.json", encoding="utf-8") as f:
        return json.load(f)


def render_zone_sentence(stage_zone, zone_descriptions):
    """Часть 1 (окончание) — предложение про Зону."""
    if stage_zone["zone_name"] == "Функциональная зона":
        return zone_descriptions["Функциональная"]
    zone_type = stage_zone["zone_type"]  # "Наводнение" или "Аэротруба"
    zone_type_gen = ZONE_TYPE_GENITIVE[zone_type]
    return f'Компания проходит через зону {zone_type_gen}. {zone_descriptions[zone_type]}'


def _quote_company_name(name):
    """Не дублирует кавычки, если название компании уже в кавычках (напр. 'ООО «Ромашка»')."""
    if '«' in name or '"' in name:
        return name
    return f'«{name}»'


def render_stage_and_zone(company_name, stage_zone, zone_descriptions):
    """Часть 1 — всегда показывается."""
    stage_id = stage_zone["stage_id"]
    stage_name_gen = STAGE_NAME_GENITIVE[stage_id]
    zone_sentence = render_zone_sentence(stage_zone, zone_descriptions)
    quoted_name = _quote_company_name(company_name)
    return (
        f'{quoted_name} — это бизнес на Стадии {stage_id}, которая называется '
        f'Стадия {stage_name_gen}. {zone_sentence}'
    )


def _format_range(lo, hi):
    return f'{lo} человек' if lo == hi else f'{lo}-{hi} человек'


def render_staffing_gap(flow_a, stage):
    """Часть 2 — условная, только если есть кадровый разрыв."""
    lines = []
    if flow_a["managers_gap"] > 0:
        lo, hi = stage["managers_range"]
        lines.append(
            f'на Стадии {stage["id"]} целевой диапазон Менеджеров — {_format_range(lo, hi)}, '
            f'у Вас сейчас на {flow_a["managers_gap"]} меньше минимального значения'
        )
    if flow_a["leaders_gap"] > 0:
        lo, hi = stage["leaders_range"]
        lines.append(
            f'целевой диапазон Высших руководителей — {_format_range(lo, hi)}, '
            f'у Вас сейчас на {flow_a["leaders_gap"]} меньше минимального значения'
        )
    if not lines:
        return None
    body = "; ".join(lines)
    # формулировка зависит от того, есть ли реальные несоответствия Потока А ниже —
    # иначе получается противоречие («причины несоответствий», а несоответствий нет)
    if flow_a["mismatches"]:
        tail = 'Это может быть одной из причин несоответствий, которые Вы увидите ниже.'
    else:
        tail = 'Это стоит держать в фокусе внимания при дальнейшем росте компании.'
    return f'Дополнительно нужно отметить: {body}. {tail}'


def render_flow_a_mismatches(flow_a, stage_id):
    """Часть 3 — две ветки: есть несоответствия / нет."""
    mismatches = flow_a["mismatches"]
    if not mismatches:
        return (
            f'Организация полностью соответствует целевым значениям Стадии {stage_id} '
            f'по всем измерениям Правил роста — это редкий и сильный результат, '
            f'требующий отдельного внимания в остальной части диагностики.'
        )
    names = [MEASUREMENT_DISPLAY_NAMES.get(m, m) for m in mismatches]
    if len(names) == 1:
        list_part = f'по измерению: {names[0]}'
    else:
        list_part = f'по следующим измерениям: {", ".join(names)}'
    return (
        f'У организации выявлены несоответствия {list_part}. Консалтинговая программа '
        f'«Критерии роста бизнеса» позволит устранить эти несоответствия и выйти на '
        f'целевые значения Стадии {stage_id}.'
    )


def render_maturity_teaser(maturity_note):
    """Часть 4 — условная, только если зрелость не 'соответствует_стадии'."""
    if maturity_note == "соответствует_стадии":
        return None
    return (
        'В Разделе 8 Вы увидите дополнительный и, возможно, неожиданный вывод о том, '
        'насколько зрелость Вашего бизнеса соответствует его текущему размеру.'
    )


def render_critical_gaps_teaser(has_critical_gaps):
    """Часть 5 — условная, только если в Разделе 9 есть блок «Критические упущения»."""
    if not has_critical_gaps:
        return None
    return (
        'Отдельно стоит обратить внимание на Раздел 9 — там перечислены требования, '
        'которые остаются невыполненными на протяжении нескольких Стадий подряд.'
    )


CLOSING_LINE = (
    'Внедрение в организации недостающих Ключевых системных элементов, которые '
    'выявлены по итогам данной диагностики, позволит бизнесу успешно продолжить '
    'дальнейший рост.'
)


# ---------------------------------------------------------------------------
# Раздел 12. Дальнейшие шаги
# ---------------------------------------------------------------------------

SECTION12_TRANSITION_MAIN = (
    'Список выше показывает, какие Ключевые системные элементы отсутствуют или '
    'недостаточно развиты в экосистеме Вашего бизнеса. Но порядок, в котором Вы '
    'будете их внедрять, имеет не меньшее значение, чем сам список.\n\n'
    'Представьте экосистему бизнеса как дом. Элементы с пометкой "Фундамент" — то, '
    'без чего невозможно опираться на остальное; они должны быть заложены в первую '
    'очередь, вне зависимости от того, кажутся ли они Вам сейчас самыми острыми. '
    'Элементы с пометкой "Ядро" — несущие стены; их можно начинать только после того, как '
    'Фундамент выдержит нагрузку. Элементы с пометкой "Надстройка" — то, что достраивается, '
    'когда основа уже устойчива.\n\n'
    'Попытка внедрить более поздний уровень раньше предыдущего не ускоряет '
    'результат — она создаёт видимость порядка без реальной устойчивости и чаще '
    'всего требует переделки, когда Фундамент всё-таки даёт о себе знать.'
)

SECTION12_TRANSITION_CRITICAL_GAPS = (
    'Обратите внимание: в Разделе 9 Вы уже видели, что часть Непреложных правил не '
    'выполняется на протяжении нескольких Стадий подряд. Это не случайность — как '
    'правило, именно застарелые, годами не устраняемые пробелы совпадают с тем, '
    'что оказывается в Фундаменте плана ниже.'
)

SECTION12_TRANSITION_CLOSING = (
    'Внедрение Ключевых системных элементов происходит в виде Консалтинговых программ.'
)

CTA_TEXT = (
    'Какой путь лучше подходит именно Вам?\n\n'
    'Как превратить его в работающий план с конкретными сроками и ресурсами именно '
    'для Вашей компании?\n\n'
    'Обсудим эти и все другие вопросы на онлайн-встрече'
)


def render_section12_1_transition(has_critical_gaps):
    if has_critical_gaps:
        return (SECTION12_TRANSITION_MAIN + "\n\n" + SECTION12_TRANSITION_CRITICAL_GAPS
                + "\n\n" + SECTION12_TRANSITION_CLOSING)
    return SECTION12_TRANSITION_MAIN + "\n\n" + SECTION12_TRANSITION_CLOSING


def render_section12_2_plan(diagnose_result, data):
    """План внедрения по ярусам — короче Раздела 11: без 'почему', только
    название + 1-2 позитивных пункта + программа."""
    programs = data["consulting_programs"]["individual_programs"]
    rows = [r for r in diagnose_result["приоритизация_ксэ"] if r["ярус"] != "низкий_приоритет"]

    grouped = {"фундамент": [], "ядро": [], "надстройка": []}
    for row in rows:
        grouped[_normalize_tier(row["ярус"])].append(row)

    blocks = []
    for tier_key in ("фундамент", "ядро", "надстройка"):
        tier_rows = grouped[tier_key]
        if not tier_rows:
            continue
        title, tier_intro = TIER_HEADERS[tier_key]
        tier_intro = tier_intro.rstrip(".") + ":"
        blocks.append(f'## {title}\n\n{tier_intro}')

        for row in tier_rows:
            kse = row["kse"]
            program = programs[kse]
            outcomes = program["что_компания_получает"][:2]
            outcomes_text = "; ".join(o[0].lower() + o[1:] for o in outcomes)
            blocks.append(
                f'**{kse}** — {outcomes_text}. — Программа «{program["название_программы"]}»'
            )

    return "\n\n".join(blocks)


def _parse_early_stage_weeks(duration_text):
    """Первое число в строке длительности = длительность для ранних Стадий
    (Игорь последовательно перечисляет диапазоны от ранних Стадий к поздним)."""
    m = re.search(r'\d+', duration_text)
    return int(m.group()) if m else None


def render_section12_3_bundle_fork(diagnose_result, data, min_intersection=4):
    """Возвращает текст развилки с бандлом, или None, если условия показа не выполнены."""
    stage_id = diagnose_result["стадия"]["stage_id"]
    if stage_id not in (1, 2, 3):
        return None

    bundle = data["consulting_programs"]["bundle_programs"]["Возрождение малого бизнеса"]
    bundle_kse = set(bundle["ксэ_покрывает"])
    plan_kse = {r["kse"] for r in diagnose_result["приоритизация_ксэ"] if r["ярус"] != "низкий_приоритет"}

    intersection = sorted(plan_kse & bundle_kse)
    if len(intersection) < min_intersection:
        return None

    programs = data["consulting_programs"]["individual_programs"]
    sum_weeks = sum(_parse_early_stage_weeks(programs[k]["длительность"]) for k in intersection)
    bundle_weeks = _parse_early_stage_weeks(bundle["длительность"])

    n = len(intersection)
    total = len(bundle_kse)

    return (
        f'В Вашем плане выше {n} из {total} Системных элементов совпадают с Консалтинговой программой '
        f'«Возрождение малого бизнеса» — комплексным пакетом, который объединяет '
        f'наиболее критичные модули сразу нескольких Программ в одном формате именно для '
        f'малого бизнеса.\n\n'
        f'Это открывает 2 способа, как Вы можете двигаться дальше:\n\n'
        f'Путь А — полные Программы по отдельности (см. список выше). Каждая '
        f'Программа даёт максимальную глубину проработки своего Системного элемента — от '
        f'диагностики конкретно Вашей ситуации до полного внедрения и обкатки в '
        f'реальной работе.\n\n'
        f'Путь Б — Комплексная программа «Возрождение малого бизнеса». Не эквивалент прохождения '
        f'всех {n} Программ полностью — это комплексный набор наиболее '
        f'критичных модулей каждой Программы. Если бы Вы проходили эти '
        f'Программы в полном объёме, это заняло '
        f'бы порядка {sum_weeks} недель. А Комплексная программа "Возрождение малого бизнеса" даёт сокращённую версию тех же '
        f'Системных элементов за {bundle_weeks} недель. Такой путь особенно подходит, если '
        f'ситуация в бизнесе требует быстрой стабилизации по многим фронтам '
        f'одновременно, а не глубокой проработки каждого Системного элемента.\n\n'
        f'Какой путь лучше подходит именно Вам?\n\n'
        f'Как превратить его в работающий план с конкретными сроками и ресурсами именно для Вашей компании?\n\n'
        f'Обсудим эти и все другие вопросы на онлайн-встрече.'
    )


def render_section12(diagnose_result, data, has_critical_gaps=False):
    parts = [
        render_section12_1_transition(has_critical_gaps),
        render_section12_2_plan(diagnose_result, data),
    ]
    bundle_text = render_section12_3_bundle_fork(diagnose_result, data)
    if bundle_text:
        parts.append(bundle_text)
    else:
        parts.append(CTA_TEXT)
    return "\n\n".join(parts)


# ---------------------------------------------------------------------------
# Раздел 11. Ключевые системные элементы (рекомендация)
# ---------------------------------------------------------------------------

TIER_HEADERS = {
    "фундамент": (
        "ФУНДАМЕНТ",
        "Это база, без которой любые дальнейшие усилия неустойчивы. Начинать нужно "
        "здесь — вне зависимости от того, что кажется более «горящим» прямо сейчас."
    ),
    "ядро": (
        "ЯДРО",
        "Несущая конструкция бизнеса. Внедряется после того, как Фундамент заложен, "
        "и формирует основной каркас управляемости и роста."
    ),
    "надстройка": (
        "НАДСТРОЙКА",
        "Достраивается поверх устойчивого Ядра — усиливает и оттачивает то, что уже "
        "работает."
    ),
}


def _normalize_tier(tier_raw):
    """'фундамент (авто-триггер Потока А)' -> 'фундамент', остальные без изменений."""
    return "фундамент" if tier_raw.startswith("фундамент") else tier_raw


def _find_symptom_challenges(kse, challenge_scores, mapping, threshold=6, limit=3):
    """Вызовы, которые указывают на этот КСЭ и получили высокую оценку клиента."""
    linked = [c for c, kses in mapping["challenge_to_kse"].items() if kse in kses]
    scored = [(c, challenge_scores.get(c, 0)) for c in linked]
    scored = [(c, s) for c, s in scored if s >= threshold]
    scored.sort(key=lambda x: -x[1])
    return [c for c, s in scored[:limit]]


def _find_deficit_areas(kse, area_deficit, mapping):
    """Области бизнеса, просевшие (deficit > 0) и указывающие на этот КСЭ."""
    linked = [a for a, kses in mapping["business_area_to_kse"].items() if kse in kses]
    return [a for a in linked if area_deficit.get(a, 0) > 0]


def _render_kse_justification(row, challenge_scores, area_deficit, mapping, flow_a):
    """Абзац 'Почему именно этот Элемент' — три ветки: авто-триггер Правил роста, авто-триггер
    кадрового разрыва, обычный (симптомы + Области).

    Важно: проверяем флаг flow_a["auto_triggers"] напрямую, а не текст яруса —
    если КСЭ и так уже был в фундаменте по маркеру самой Стадии, триггер не меняет
    текст яруса видимым образом, но сам факт срабатывания всё равно есть и должен
    быть показан клиенту.
    """
    kse = row["kse"]
    triggered = flow_a["auto_triggers"].get(kse, False)

    if triggered and kse == "Критерии роста бизнеса":
        return (
            'Этот Элемент добавлен в список не по анализу симптомов, а потому что в '
            'Разделах 2-6 у Вас выявлены несоответствия базовым Правилам роста для Вашей '
            'Стадии. «Критерии роста бизнеса» — это Ключевой системный элемент, который отвечает именно за '
            'внедрение системы этих Правил в бизнесе.'
        )

    if triggered and kse in ("Организационная структура", "Сильная Управленческая команда"):
        gap_parts = []
        if flow_a["managers_gap"] > 0:
            gap_parts.append("Менеджеров")
        if flow_a["leaders_gap"] > 0:
            gap_parts.append("Высших руководителей")
        gap_text = " и ".join(gap_parts) if gap_parts else "управленческого состава"
        return (
            f'Этот Элемент добавлен в список не по анализу симптомов, а потому что в '
            f'Разделе 1 у Вас выявлен кадровый разрыв по количеству {gap_text} — без '
            f'этого Элемента невозможно системно закрыть такой разрыв.'
        )

    # обычный случай — симптомы + просевшие Области
    challenges = _find_symptom_challenges(kse, challenge_scores, mapping)
    areas = _find_deficit_areas(kse, area_deficit, mapping)

    parts = []
    if challenges:
        pronoun_verb = "показал" if len(challenges) == 1 else "показали"
        parts.append(f'{_join_ru(challenges)} {pronoun_verb} наибольшую тяжесть в Ваших ответах')
    if areas:
        pronoun_verb2 = "просела" if len(areas) == 1 else "просели"
        parts.append(f'по Непреложным правилам {pronoun_verb2} {_join_ru(areas)}')

    if not parts:
        return 'Этот Элемент относится к приоритетным для Вашей текущей Стадии роста.'
    return "Почему именно этот Элемент: " + "; также ".join(parts) + "."


def render_section11_kse_list(diagnose_result, client_challenge_scores, data):
    """
    client_challenge_scores: сырые баллы 0-10 клиента по 24 Вызовам (вход, не
        часть diagnose_result — тот хранит только нормализованный сигнал)
    """
    mapping = data["mapping_kse"]
    programs = data["consulting_programs"]["individual_programs"]
    flow_a = diagnose_result["поток_а"]
    area_deficit = flow_a["area_deficit"]

    rows = [r for r in diagnose_result["приоритизация_ксэ"] if r["ярус"] != "низкий_приоритет"]

    grouped = {"фундамент": [], "ядро": [], "надстройка": []}
    for row in rows:
        grouped[_normalize_tier(row["ярус"])].append(row)

    blocks = []
    for tier_key in ("фундамент", "ядро", "надстройка"):
        tier_rows = grouped[tier_key]
        if not tier_rows:
            continue
        title, tier_intro = TIER_HEADERS[tier_key]
        blocks.append(f'## {title}\n\n{tier_intro}')

        for row in tier_rows:
            kse = row["kse"]
            program = programs[kse]
            outcomes = program["что_компания_получает"][:2]
            outcomes_text = "; ".join(o[0].lower() + o[1:] for o in outcomes)
            justification = _render_kse_justification(
                row, client_challenge_scores, area_deficit, mapping, flow_a
            )
            block = (
                f'### {kse}\n\n'
                f'{justification}\n\n'
                f'Внедрение этого Элемента даёт: {outcomes_text}.\n\n'
                f'→ Закрывает Консалтинговая программа «{program["название_программы"]}»'
            )
            blocks.append(block)

    return "\n\n".join(blocks)


# ---------------------------------------------------------------------------
# Раздел 8. Уровень зрелости бизнеса
# ---------------------------------------------------------------------------

def _join_ru(items):
    """Соединяет список строк в русское перечисление: 'A', 'A и B', 'A, B и C'."""
    items = list(items)
    if not items:
        return ""
    if len(items) == 1:
        return f'«{items[0]}»'
    quoted = [f'«{i}»' for i in items]
    return ", ".join(quoted[:-1]) + f' и {quoted[-1]}'


def render_section8_maturity(diagnose_result, stage_id, time_at_stage_text=None):
    """
    stage_id: текущая Стадия клиента
    time_at_stage_text: строка вида "2 года 3 месяца" из поля 10 квалификации,
        или None, если клиент не указал / оставил пустым
    """
    flow_b = diagnose_result["поток_б"]
    top5 = flow_b["individual_top5"]
    maturity = flow_b["maturity"]
    note = maturity["note"]
    later = maturity["later_stage_challenges"]
    earlier = maturity["earlier_stage_challenges"]

    top5_list = "\n".join(f'{i+1}. {c}' for i, c in enumerate(top5))
    intro = (
        f'На основе Ваших ответов мы определили индивидуальный Топ-5 вызовов именно '
        f'Вашей компании — не общих для Стадии {stage_id}, а тех, что реально беспокоят '
        f'Вас сильнее всего прямо сейчас:\n\n{top5_list}'
    )

    if note == "опережает_стадию":
        is_single = len(later) == 1
        pronoun = "который" if is_single else "которые"
        verb = "проявляется" if is_single else "проявляются"
        interp = (
            f'Среди них — {_join_ru(later)}, {pronoun} по нашим исследованиям типично '
            f'{verb} на более поздних Стадиях роста, а не на Стадии {stage_id}, где '
            f'сейчас находится Ваш бизнес.\n\n'
            f'Это означает, что Ваш бизнес функционально более зрелый, чем предполагает '
            f'его текущий размер — Вы уже прошли через типичные сложности этой Стадии. Но '
            f'именно поэтому объективно «застряли» на ней не из-за незрелости, а из-за '
            f'отсутствия конкретных Ключевых системных элементов (см. Раздел 11). Здесь '
            f'требуются точечные, но безотлагательные действия.'
        )
    elif note == "отстаёт_от_стадии":
        is_single = len(earlier) == 1
        pronoun = "который" if is_single else "которые"
        verb = "характерен" if is_single else "характерны"
        interp = (
            f'Среди них — {_join_ru(earlier)}, {pronoun} по нашим исследованиям типично '
            f'{verb} для более ранних Стадий роста, а не для Стадии {stage_id}, на '
            f'которой сейчас находится Ваш бизнес.\n\n'
            f'Это означает, что функциональная зрелость Вашего бизнеса отстаёт от его '
            f'фактического размера — компания выросла быстрее, чем успела выстроить '
            f'внутреннюю систему под этот масштаб. Отсюда может возникать ощущение, что '
            f'бизнес «больше, чем Вы можете унести». Здесь тоже требуются безотлагательные '
            f'действия.'
        )
    elif note == "смешанная_картина":
        interp = (
            f'Среди них одновременно есть и вызовы, типичные для более ранних Стадий '
            f'({_join_ru(earlier)}), и вызовы, типичные для более поздних Стадий '
            f'({_join_ru(later)}), чем текущая Стадия {stage_id} Вашего бизнеса.\n\n'
            f'Это довольно редкое и по-своему показательное сочетание: в одних '
            f'направлениях бизнес уже перерос свой формальный размер, а в других — ещё не '
            f'дозрел до него. Оба этих сигнала указывают на одну и ту же причину — '
            f'отсутствие конкретных Ключевых системных элементов (см. Раздел 11), а не '
            f'общую незрелость или временные трудности роста.'
        )
    else:  # соответствует_стадии
        interp = (
            f'Ваш индивидуальный Топ-5 полностью укладывается в типичную картину для '
            f'Стадии {stage_id} — серьёзных сигналов о том, что зрелость бизнеса опережает '
            f'или отстаёт от его размера, не обнаружено. Это хороший знак: трудности, с '
            f'которыми Вы сталкиваетесь, ожидаемы для Вашего масштаба, и системная работа '
            f'над Ключевыми элементами (Раздел 11) должна дать предсказуемый результат.'
        )

    paragraphs = [intro, interp]

    # кросс-проверка времени на Стадии — только фиксация факта, без сравнения с
    # "типичным сроком" (такого бенчмарка у нас нет в данных)
    if time_at_stage_text and note != "соответствует_стадии":
        paragraphs.append(
            f'Дополнительный факт для контекста: в компании сохраняется текущее '
            f'количество сотрудников уже {time_at_stage_text}.'
        )

    return "\n\n".join(paragraphs)


def render_section1_vyvody(diagnose_result, company_name, stages_data, has_critical_gaps=False):
    """
    Собирает финальный текст блока ВЫВОДЫ из результата diagnose() (scoring_algorithm.py).

    has_critical_gaps: bool — передаётся отдельно, вычисляется в логике Раздела 9
        (правило не с текущей Стадии + % < 50), в diagnose_result сейчас не входит.
    """
    zone_descriptions = load_zone_descriptions()
    stage_zone = diagnose_result["стадия"]
    flow_a = diagnose_result["поток_а"]
    stage = next(s for s in stages_data["stages"] if s["id"] == stage_zone["stage_id"])

    paragraphs = [render_stage_and_zone(company_name, stage_zone, zone_descriptions)]

    gap_text = render_staffing_gap(flow_a, stage)
    if gap_text:
        paragraphs.append(gap_text)

    paragraphs.append(render_flow_a_mismatches(flow_a, stage_zone["stage_id"]))

    maturity_text = render_maturity_teaser(diagnose_result["поток_б"]["maturity"]["note"])
    if maturity_text:
        paragraphs.append(maturity_text)

    gaps_text = render_critical_gaps_teaser(has_critical_gaps)
    if gaps_text:
        paragraphs.append(gaps_text)

    paragraphs.append(CLOSING_LINE)

    return "\n\n".join(paragraphs)


# ---------------------------------------------------------------------------
# Разделы 2-6 (Поток А) — вступительные тексты, таблицы «текущее/целевое»,
# и резолверы блока ВЫВОДЫ (выбор нужной ветки шаблона по факту диагностики).
# Тексты вступлений и веток взяты из Шаблон_Отчёта__Стадия_3.pdf (не зависят
# от Стадии, поэтому не хранятся в stage_level_report_texts.json).
# ---------------------------------------------------------------------------

MANAGEMENT_STYLE_DESCRIPTIONS = {
    "Эталонный": (
        "Лидер, использующий Эталонный стиль, устанавливает крайне высокие стандарты "
        "производительности и личным примером показывает, как им соответствовать. Он "
        "требует большого усердия, а к низким результатам относится без сочувствия. Он "
        "нанимает целеустремлённых высокопрофессиональных сотрудников с высокой мотивацией, "
        "хорошо ладит с такими сотрудниками и выжимает из них максимум."
    ),
    "Директивный": (
        "Основные черты такого стиля – это Влияние, Нацеленность на результат, Инициатива. "
        "Директивный стиль управления лучше всего помогает вывести компанию из кризиса, "
        "пройти глубокую реорганизацию бизнеса или решить вопросы с проблемными "
        "сотрудниками. Если вдруг нужно тушить пожар, то требуется директивный лидер."
    ),
    "Коучинговый": (
        "Коучинговый лидер занимается развитием сотрудников. Он действует под девизом: "
        "«Попробуй вот это!» Такой лидер помогает сотрудникам ставить перед собой "
        "долгосрочные профессиональные цели, раскрыть индивидуальный потенциал сотрудника и "
        "разработать план достижения цели. Он уверен, что в каждом есть потенциал и главное "
        "– это грамотно его раскрыть. Такой лидер регулярно даёт обратную связь и делает "
        "подсказки, он готов мириться с сиюминутными неудачами, если это способствует "
        "долгосрочному обучению сотрудника и дальнейшему достижению успеха."
    ),
    "Авторитетный": (
        "Авторитетный лидер мобилизует и вдохновляет сотрудников на воплощение своих "
        "замыслов для реализации образа будущего организации. Когда коллективная задача "
        "формулируется с точки зрения большого образа будущего, люди понимают, что то, что "
        "они делают, действительно, имеет значение. Яркий энтузиазм является отличительной "
        "чертой этого стиля. Этот стиль может сильно влиять на эмоциональный климат "
        "организации и трансформировать дух организации."
    ),
    "Демократический": (
        "Демократический лидер достигает консенсуса на основе коллективных решений. Такой "
        "лидер старается вовлечь сотрудников в принятие любых решений. Для него важно, чтобы "
        "все или хотя бы абсолютное большинство согласились с тем, что он предлагает. Такой "
        "лидер знает, как подавить конфликт и создать чувство гармонии. Он силён в "
        "укреплении доверия и уважения."
    ),
    "Товарищеский": (
        "При товарищеском стиле лидер фокусируется на создании гармонии в команде и "
        "построении эмоциональных связей с её членами. Такой лидер видит в сотрудниках "
        "индивидуальность и меньше делает упор на выполнение задач и достижение целей. Такой "
        "лидер создаёт очень высокий уровень лояльности и укрепляет межличностные отношения."
    ),
}


def render_natural_styles_ranking(scores, top3):
    """Возвращает список (название_стиля, баллы, ярлык_или_None) в порядке убывания
    баллов — для блока «Естественное сочетание стилей управления»."""
    top3_by_style = {v: k for k, v in top3.items()}  # {"Директивный": "Основной", ...}
    ranked = sorted(scores.items(), key=lambda kv: -kv[1])
    return [(style, score, top3_by_style.get(style)) for style, score in ranked]


FLOW_A_INTRO_TEXTS = {
    "Приоритетные сферы": (
        "Вся деятельность организации относится к трём основным сферам: Люди, Прибыль и "
        "Процессы. На конкретной Стадии роста этим трём сферам нужно уделять внимание в "
        "строго определённом порядке. Каждая из этих сфер всегда важна. Но в зависимости "
        "от Вашей Стадии роста приоритетная расстановка этих сфер меняется."
    ),
    "Строитель-Протектор": (
        "Коэффициент Строитель-Протектор служит для определения общего организационного "
        "мышления в отношении уровней уверенности и осмотрительности в компании. Он "
        "выражается в виде математического значения – т.е. коэффициента – и показывает "
        "отношение уверенности к осмотрительности, которые в целом испытывает организация. "
        "Строители – это те, кто принимают новые возможности. Они готовы идти на риск и "
        "поддерживают развитие и рост. Они любят совершенствовать привычные методы работы. "
        "Они – словно педаль газа организации. Протекторы – это те, кого нужно убеждать "
        "принять новые бизнес возможности. Они всегда осмотрительны, не любят рисковать и "
        "предпочитают не менять привычный порядок и уклад. Они словно педаль тормоза "
        "организации и предпочитают замедлить темп изменений.\n\n"
        "Для того чтобы успешно развиваться, бизнес должен эффективно сбалансировать эти два "
        "типа мышления в соответствии с той Стадией роста, к которой он относится. Потому что "
        "обе крайности убивают бизнес. Если у Вас в компании слишком много Строителей, они "
        "могут продолжать давить на газ, даже когда компания будет на грани срыва в пропасть. "
        "Если у Вас слишком много Протекторов, они в конце концов могут задушить компанию, "
        "встав ногой во весь рост на педаль тормоза."
    ),
    "Модальность": (
        "Модальность — это роль, которую каждый уровень сотрудников в организации играет в "
        "отношении всей организации. Три уровня сотрудников: Высшее руководство, Менеджеры "
        "и Рядовые сотрудники. На каждой Стадии роста каждый из трёх уровней выполняет "
        "определённую роль. Существует три модальности: Доминирующая (оказывающая основное "
        "влияние), Поддерживающая (помогающая в достижении целей) и Вспомогательная "
        "(облегчающая достижение целей). Несоответствие целевой комбинации ролей, особенно "
        "в отношении Доминирующей роли, может существенно повлиять на способность "
        "организации поддерживать развитие и рост."
    ),
    "Три роли лидера": (
        "Измерение «Три роли лидера» определяет, какую долю времени и энергии владелец "
        "бизнеса затрачивает на выполнение трёх основных функций лидера в организации — "
        "Визионера, Менеджера и Специалиста. На каждой Стадии роста бизнесу нужно, чтобы "
        "владелец разное время посвящал каждой из этих ролей."
    ),
    "Стили управления": (
        "Для каждой Стадии роста существует целевое (идеальное) сочетание трёх стилей "
        "управления: первостепенного (основного), второстепенного и третьестепенного "
        "(дополнительного). Сочетание стилей управления может кардинально меняться от "
        "Стадии к Стадии, поэтому естественное сочетание стилей, которое хорошо работало на "
        "одной Стадии, может быть совсем неэффективным на другой."
    ),
}


def _first_two_sentences(text_block):
    """Возвращает (совпадает_текст, несовпадает_текст) — первые два
    предложения шаблона, до первого 'ЕСЛИ' или до конца, если условий нет."""
    parts = text_block.split(". ИЛИ ", 1)
    match_stmt = parts[0].strip().rstrip(".") + "."
    rest = parts[1] if len(parts) > 1 else ""
    # отделяем второе предложение (несоответствие) от возможного "ЕСЛИ ...:" или detail
    m = re.match(r"^(.*?\.)\s*(.*)$", rest, re.S)
    if m:
        mismatch_stmt = m.group(1).strip()
        detail = m.group(2).strip()
    else:
        mismatch_stmt = rest.strip()
        detail = ""
    return match_stmt, mismatch_stmt, detail


def render_priority_spheres_vyvody(client_val, target_val, stage_id, priority_spheres_scenarios):
    """client_val/target_val: списки сфер по приоритету, напр.
    ["Прибыль","Люди","Процессы"]. priority_spheres_scenarios: dict из
    data/priority_spheres_scenarios.json — 5 сценариев отклонения на Стадию
    (6-й, целевой, здесь не нужен — для него используется старый text_block,
    см. вызов ниже)."""
    if client_val == target_val:
        return None  # вызывающая сторона подставит старый match_stmt
    key = "|".join(client_val)
    scenario = priority_spheres_scenarios[str(stage_id)].get(key)
    if scenario:
        return scenario
    # защита: если вдруг перестановка не найдена (не должно случаться —
    # все 5 неверных перестановок на каждую Стадию покрыты) — не падаем
    return "Текущая расстановка приоритетов отличается от целевой для этой Стадии."


def render_builder_protector_vyvody(client_val, target_val, stage_id, builder_protector_scenarios):
    """client_val/target_val: строки-отношения вида '2:1'.
    builder_protector_scenarios: dict из data/builder_protector_scenarios.json
    — 2 сценария на Стадию ('below'/'above')."""
    if client_val == target_val:
        return None

    def _ratio_value(s):
        try:
            a, b = s.split(":")
            return float(a) / float(b)
        except Exception:
            return None

    cv, tv = _ratio_value(client_val), _ratio_value(target_val)
    stage_scenarios = builder_protector_scenarios[str(stage_id)]
    if cv is not None and tv is not None and cv < tv:
        return stage_scenarios["below"]
    if cv is not None and tv is not None and cv > tv:
        return stage_scenarios["above"]
    return "Текущее значение Коэффициента отличается от целевого для этой Стадии."


def render_modality_vyvody(client_val, target_val, stage_id, modality_scenarios):
    """client_val/target_val: {уровень: роль}, напр.
    {"Руководство":"Доминирующая", "Менеджеры":"Поддерживающая",
    "Сотрудники":"Вспомогательная"}. modality_scenarios: dict из
    data/modality_scenarios.json — 5 сценариев на Стадию, ключ —
    'Уровень_на_Доминирующей|Уровень_на_Поддерживающей|Уровень_на_Вспомогательной'."""
    if client_val == target_val:
        return None
    role_to_level = {role: level for level, role in client_val.items()}
    try:
        key = "|".join([role_to_level["Доминирующая"], role_to_level["Поддерживающая"],
                         role_to_level["Вспомогательная"]])
    except KeyError:
        return "Текущая комбинация ролей отличается от целевой для этой Стадии."
    scenario = modality_scenarios[str(stage_id)].get(key)
    if scenario:
        return scenario
    return "Текущая комбинация ролей отличается от целевой для этой Стадии."


def render_leader_roles_vyvody(client_val, target_val, text_block):
    match_stmt, mismatch_stmt, detail = _first_two_sentences(text_block)
    if client_val == target_val:
        return match_stmt

    has_any_marker = bool(re.search(r"ЕСЛИ (МНОГО|МАЛО) (СПЕЦИАЛИСТА|МЕНЕДЖЕРА|ВИЗИОНЕРА):", detail))
    if not has_any_marker:
        # текст этой Стадии не размечен по ролям явно — единый блок,
        # относящийся к самому частому для неё несоответствию (см. пояснение в чате)
        return f"{mismatch_stmt} {detail.strip()}" if detail.strip() else mismatch_stmt

    marker_defs = [
        ("МНОГО", "СПЕЦИАЛИСТА", "Специалист", lambda c, t: c > t),
        ("МАЛО", "СПЕЦИАЛИСТА", "Специалист", lambda c, t: c < t),
        ("МНОГО", "МЕНЕДЖЕРА", "Менеджер", lambda c, t: c > t),
        ("МАЛО", "МЕНЕДЖЕРА", "Менеджер", lambda c, t: c < t),
        ("МНОГО", "ВИЗИОНЕРА", "Визионер", lambda c, t: c > t),
        ("МАЛО", "ВИЗИОНЕРА", "Визионер", lambda c, t: c < t),
    ]
    parts = [mismatch_stmt]
    for qty, role_word, role_key, cond in marker_defs:
        pattern = (rf"ЕСЛИ {qty} {role_word}:\s*(.*?)"
                   r"(?:ЕСЛИ (?:МНОГО|МАЛО) (?:СПЕЦИАЛИСТА|МЕНЕДЖЕРА|ВИЗИОНЕРА):|$)")
        m = re.search(pattern, detail, re.S)
        if m and cond(client_val.get(role_key, 0), target_val.get(role_key, 0)):
            fragment = m.group(1).strip()
            if fragment and fragment not in parts:
                parts.append(fragment)
    if len(parts) == 1:
        return mismatch_stmt
    return " ".join(parts)


def render_management_styles_vyvody(client_val, target_val, text_block):
    """ВЫВОДЫ для Раздела 6. Для каждого реального расхождения — лишний стиль
    у клиента, недостающий целевой стиль, или стиль есть в обоих сочетаниях,
    но на разных позициях (недооценён/переоценён) — берёт готовый текст из
    text_block["выводы_по_стилям"], если он написан для этой Стадии. Для
    расхождений, для которых текст ещё не написан, — честная нейтральная
    фраза без домысливания причин."""
    match_stmt, mismatch_stmt, _ = _first_two_sentences(text_block["выводы_обе_ветки"])
    if client_val == target_val:
        return match_stmt

    RANK = {"Основной": 1, "Второстепенный": 2, "Дополнительный": 3}
    client_slot_by_style = {style: slot for slot, style in client_val.items()}
    target_slot_by_style = {style: slot for slot, style in target_val.items()}

    client_top3 = set(client_val.values())
    target_top3 = set(target_val.values())
    excess = sorted(client_top3 - target_top3)
    missing = sorted(target_top3 - client_top3)
    shared_misplaced = sorted(
        style for style in (client_top3 & target_top3)
        if client_slot_by_style[style] != target_slot_by_style[style]
    )

    library = text_block.get("выводы_по_стилям", {})
    excess_lib = library.get("excess", {})
    missing_lib = library.get("missing", {})
    excess_pairs_lib = library.get("excess_pairs", {})
    misplaced_lib = library.get("misplaced", {})
    under_lib = misplaced_lib.get("under", {})
    over_lib = misplaced_lib.get("over", {})

    paragraphs = [mismatch_stmt]
    remaining_excess = set(excess)
    uncovered_excess, uncovered_missing, uncovered_misplaced = [], [], []

    # сначала — авторские тексты для классических парных сценариев (например,
    # «Директивный+Авторитетный» на Стадии 3), если оба стиля пары лишние одновременно
    for pair_key, text in excess_pairs_lib.items():
        pair = set(pair_key.split("+"))
        if pair <= remaining_excess:
            paragraphs.append(text)
            remaining_excess -= pair

    for style in sorted(remaining_excess):
        if style in excess_lib:
            paragraphs.append(excess_lib[style])
        else:
            uncovered_excess.append(style)
    for style in missing:
        if style in missing_lib:
            paragraphs.append(missing_lib[style])
        else:
            uncovered_missing.append(style)
    for style in shared_misplaced:
        client_slot = client_slot_by_style[style]
        target_slot = target_slot_by_style[style]
        underranked = RANK[client_slot] > RANK[target_slot]  # у клиента слабее целевого
        lib = under_lib if underranked else over_lib
        if style in lib:
            paragraphs.append(lib[style].format(client_slot=client_slot, target_slot=target_slot))
        else:
            uncovered_misplaced.append((style, client_slot, target_slot))

    if uncovered_excess or uncovered_missing or uncovered_misplaced:
        fact_parts = []
        if uncovered_excess:
            fact_parts.append(f'в Вашем сочетании есть {", ".join(uncovered_excess)}, чего нет в целевом сочетании')
        if uncovered_missing:
            fact_parts.append(f'в целевом сочетании есть {", ".join(uncovered_missing)}, чего нет в Вашем сочетании')
        if uncovered_misplaced:
            mp_txt = "; ".join(f'{style} у Вас {cs}, а должен быть {ts}' for style, cs, ts in uncovered_misplaced)
            fact_parts.append(f'по позициям в сочетании: {mp_txt}')
        paragraphs.append(f'Кроме того: {"; ".join(fact_parts)}.')

    return "\n\n".join(paragraphs)
