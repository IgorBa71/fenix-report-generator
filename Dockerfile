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

# 26.08.2026: заставляем requests использовать системное хранилище
# сертификатов вместо своего собственного (certifi) — на случай если
# у каких-то серверов, к которым мы обращаемся, чуть отличается набор
# доверенных цепочек. Основная причина проблемы с MAX API устранена ниже
# отдельно (добавлением конкретного сертификата Минцифры) — эта строка
# просто дополнительная подстраховка.
ENV REQUESTS_CA_BUNDLE=/etc/ssl/certs/ca-certificates.crt

COPY requirements.txt /code
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt

# 26.08.2026: сервер MAX API (platform-api2.max.ru) с 2026 года использует
# российский сертификат Минцифры (Russian Trusted Root/Sub CA) — он не
# входит ни в обычное системное хранилище ОС, ни в certifi. Без него
# requests не может проверить HTTPS-соединение к MAX и падает с
# CERTIFICATE_VERIFY_FAILED. Файлы сертификатов лежат в /code/certs
# (скачаны с официального источника Минцифры gu-st.ru), устанавливаем их
# в системное хранилище доверенных сертификатов.
COPY certs/russian_trusted_root_ca.pem /usr/local/share/ca-certificates/russian_trusted_root_ca.crt
COPY certs/russian_trusted_sub_ca.pem /usr/local/share/ca-certificates/russian_trusted_sub_ca.crt
RUN update-ca-certificates

COPY . /code

EXPOSE 8000

# Та же команда запуска, что настроена на Timeweb (--workers 1 —
# результат сегодняшней правки под ограничение по памяти).
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8000", "--timeout", "120", "--workers", "1", "--threads", "4"]
