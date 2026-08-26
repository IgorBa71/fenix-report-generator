FROM python:3.14-slim

WORKDIR /code

# Системные зависимости: build-essential и libpq-dev нужны для сборки
# psycopg (даже несмотря на [binary]-вариант, некоторые окружения всё
# равно требуют заголовки libpq на этапе установки); curl — для
# возможных проверок доступности внутри контейнера; postgresql-client —
# даёт утилиту pg_dump для ежедневного бэкапа БД (26.08.2026, Dockhost —
# без managed PostgreSQL, бэкапы делаем сами из кода приложения).
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    build-essential \
    libpq-dev \
    postgresql-client \
    ca-certificates \
    && update-ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# 26.08.2026: библиотека requests по умолчанию проверяет HTTPS-сертификаты
# через свой собственный набор (пакет certifi), а не через системное
# хранилище ОС. Обнаружено на реальном тесте: сервер MAX API
# (platform-api2.max.ru) не полностью присылает цепочку сертификатов —
# certifi это ловит как ошибку, а более полное системное хранилище ОС
# (используется, например, отправкой email через smtplib) — нет. Эта
# переменная заставляет requests использовать системное хранилище вместо
# своего собственного, для ВСЕХ вызовов requests в приложении (Метрика,
# Яндекс.Диск, MAX) — универсальное решение, а не патч под один сервис.
ENV REQUESTS_CA_BUNDLE=/etc/ssl/certs/ca-certificates.crt

COPY requirements.txt /code
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt

COPY . /code

EXPOSE 8000

# Та же команда запуска, что настроена на Timeweb (--workers 1 —
# результат сегодняшней правки под ограничение по памяти).
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8000", "--timeout", "120", "--workers", "1", "--threads", "4"]
