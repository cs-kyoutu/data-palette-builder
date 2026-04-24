# syntax=docker/dockerfile:1.7
FROM python:3.11.9-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8002

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY backend ./backend
COPY frontend ./frontend
COPY skills ./skills
COPY templates ./templates

RUN mkdir -p backend/uploads backend/output \
    && groupadd --system app \
    && useradd --system --gid app --uid 10001 --home-dir /app --shell /usr/sbin/nologin app \
    && chown -R app:app /app

USER app

EXPOSE 8002

CMD ["sh", "-c", "uvicorn backend.app:app --host 0.0.0.0 --port ${PORT}"]
