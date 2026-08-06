FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
        libglib2.0-0 \
        libgomp1 \
        fonts-dejavu-core \
        tesseract-ocr \
        tesseract-ocr-rus \
        tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

# Индекс пакетов можно подменить на зеркало, если pypi.org недоступен/медленный:
#   docker compose build --build-arg PIP_INDEX_URL=https://<зеркало>/simple
# Таймаут и число повторов подняты: дефолтные 15 с рвут соединение на медленном канале.
ARG PIP_INDEX_URL=https://pypi.org/simple
ARG PIP_TRUSTED_HOST=""
ENV PIP_INDEX_URL=${PIP_INDEX_URL} \
    PIP_DEFAULT_TIMEOUT=120 \
    PIP_RETRIES=10

COPY requirements.txt .
# vendor/ — заранее скачанные .whl для офлайн-сборки (когда сервер не видит pypi.org).
# Если каталог непустой — ставим только из него, в сеть не ходим вообще.
# Пустой (только .gitkeep) — обычная установка из индекса.
COPY vendor ./vendor
RUN if [ -n "$(ls -A vendor 2>/dev/null | grep -v '^\.gitkeep$')" ]; then \
        echo ">>> Офлайн-установка из vendor/" && \
        pip install --no-index --find-links=vendor -r requirements.txt; \
    elif [ -n "$PIP_TRUSTED_HOST" ]; then \
        pip install --trusted-host "$PIP_TRUSTED_HOST" -r requirements.txt; \
    else \
        pip install -r requirements.txt; \
    fi

COPY . .

RUN mkdir -p uploads outputs && \
    [ -f gigachat_settings.json ] || echo '{}' > gigachat_settings.json

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
