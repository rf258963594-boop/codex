FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONPATH=/app/app \
    APP_HOST=0.0.0.0 \
    APP_PORT=8088 \
    SOFFICE_PATH=/usr/bin/soffice \
    LIBREOFFICE_PROFILE_DIR=/tmp/libreoffice-profile \
    HOME=/tmp

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        fontconfig \
        fonts-dejavu \
        fonts-liberation \
        fonts-noto-core \
        fonts-noto-cjk \
        fonts-noto-cjk-extra \
        libreoffice-common \
        libreoffice-core \
        libreoffice-writer \
        poppler-utils \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY templates ./templates
COPY README.md README_DEPLOY.md DEPLOYMENT.md RENDERING.md ./

RUN mkdir -p \
        app/data \
        app/uploads \
        app/generated \
        app/data/signatures \
        /tmp/libreoffice-profile \
    && chmod -R 755 /app

EXPOSE 8088

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8088/login', timeout=3).read()" || exit 1

CMD ["python", "app/server.py"]
