FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    APP_HOST=0.0.0.0 \
    APP_PORT=8088 \
    SOFFICE_PATH=/usr/bin/soffice \
    LIBREOFFICE_PROFILE_DIR=/tmp/libreoffice-profile

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-core \
        libreoffice-common \
        fontconfig \
        fonts-liberation \
        fonts-noto-core \
        fonts-noto-cjk \
        fonts-noto-cjk-extra \
        poppler-utils \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY outputs ./outputs
COPY README.md RENDERING.md ./

RUN mkdir -p app/data app/uploads app/generated outputs/.lo-profile-codex /tmp/libreoffice-profile

EXPOSE 8088

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8088/login', timeout=3).read()" || exit 1

CMD ["python", "app/server.py"]
