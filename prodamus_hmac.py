"""
prodamus_hmac.py — реализация алгоритма подписи Prodamus (HMAC-SHA256).

Алгоритм подтверждён ЭМПИРИЧЕСКИ 24.07.2026 на демо-форме Продамус
(https://demo.payform.ru, публичный тестовый секретный ключ из их же
документации) — сгенерированная этим кодом подписанная ссылка была
успешно принята сервером Продамус и вернула рабочую короткую ссылку
на оплату, в том числе с кириллицей в названии товара.

Шаги алгоритма (по документации Продамус + проверено на практике):
1. Все значения (включая вложенные, включая числа) приводятся к строкам.
2. Словарь рекурсивно сортируется по ключам в алфавитном порядке на всех
   уровнях вложенности (списки не сортируются, порядок элементов в них
   сохраняется как есть — сортируются только словари/объекты).
3. Результат сериализуется в JSON:
   - юникод (кириллица) НЕ экранируется (сырой UTF-8, как есть) —
     это ключевое отличие от JSON-кодирования "по умолчанию" в PHP;
   - символ "/" экранируется как "\\/";
   - без пробелов между элементами (компактный формат).
4. HMAC-SHA256 от этой JSON-строки с секретным ключом формы, результат —
   в шестнадцатеричном виде (hex digest).

ВАЖНО: секретный ключ используется тот, что указан в настройках
конкретной платёжной формы (личный кабинет Продамус → Настройки).
"""

import json
import hmac
import hashlib


def _stringify(value):
    """Рекурсивно приводит все значения к строкам."""
    if isinstance(value, dict):
        return {k: _stringify(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_stringify(v) for v in value]
    if isinstance(value, bool):
        return "1" if value else "0"
    return str(value)


def _sort_recursive(value):
    """Рекурсивно сортирует словари по ключам на всех уровнях вложенности."""
    if isinstance(value, dict):
        return {k: _sort_recursive(value[k]) for k in sorted(value.keys())}
    if isinstance(value, list):
        return [_sort_recursive(v) for v in value]
    return value


def prodamus_sign(data: dict, secret_key: str) -> str:
    """Строит подпись Prodamus для словаря data. НЕ включай в data сам
    ключ 'signature' — подпись считается ДО его добавления, добавь
    результат этой функции в data['signature'] уже после вызова."""
    prepared = _stringify(data)
    prepared = _sort_recursive(prepared)
    json_str = json.dumps(prepared, ensure_ascii=False, separators=(",", ":"))
    json_str = json_str.replace("/", "\\/")
    return hmac.new(
        secret_key.encode("utf-8"), json_str.encode("utf-8"), hashlib.sha256
    ).hexdigest()


def prodamus_verify(data: dict, secret_key: str, signature: str) -> bool:
    """Проверяет подпись входящего вебхука. data — тело запроса (без поля
    signature, если оно там оказалось), signature — значение из заголовка
    'Sign'. Возвращает True, если подпись подлинная."""
    if not signature:
        return False
    expected = prodamus_sign(data, secret_key)
    return hmac.compare_digest(expected, signature)
