FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp \
    MALLOC_TRIM_THRESHOLD_=65536

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-crosextra-carlito \
        fonts-liberation2 \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /usr/share/doc/* /usr/share/man/* /usr/share/info/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && rm -rf /root/.cache

COPY . .

RUN mkdir -p /app/uploads /app/output

EXPOSE 10000

CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 1 --timeout 210 --access-logfile - --error-logfile - app:app"]
