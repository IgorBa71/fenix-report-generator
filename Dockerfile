FROM python:3.14-slim

WORKDIR /code

# Системные зависимости: build-essential и libpq-dev нужны для сборки
# psycopg (даже несмотря на [binary]-вариант, некоторые окружения всё
# равно требуют заголовки libpq на этапе установки); curl — для
# возможных проверок доступности внутри контейнера.
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    build-essential \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /code
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt

COPY . /code

EXPOSE 8000

# Та же команда запуска, что настроена на Timeweb (--workers 1 —
# результат сегодняшней правки под ограничение по памяти).
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8000", "--timeout", "120", "--workers", "1", "--threads", "4"]
