# fenix-report-generator

Веб-сервис на Flask, оборачивающий Python-алгоритм диагностики бизнеса
(`scoring_algorithm.py`) и генератор PDF-отчёта (`pdf_report_builder.py`)
в HTTP API — чтобы Make.com мог их вызывать (Make.com исполняет только
JavaScript, не Python).

## Деплой на Render.com

1. Создать новый репозиторий на GitHub, залить туда все файлы этой папки
2. В Render: New → Web Service → подключить этот репозиторий
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `gunicorn app:app`
5. После деплоя Render выдаст URL вида `https://fenix-report-generator.onrender.com`

## Проверка после деплоя

```
GET  https://<ваш-render-url>/           → {"status": "ok", ...}
POST https://<ваш-render-url>/generate-report   → генерация отчёта
```

См. `app.py` — в шапке файла описан точный формат JSON, который нужно
отправить в `/generate-report`, и что вернётся в ответ.

## Важно — непроверенные предположения

Отмечены прямо в коде (`app.py`, docstring в начале файла):
1. Формула агрегации "Приоритетные сферы" — предположение по аналогии,
   нужно подтвердить у Игоря.
2. Маппинг утверждений Раздела 8 на 11 КСЭ — по порядковому совпадению,
   нужно сверить с оригинальным файлом-ключом проекта.
