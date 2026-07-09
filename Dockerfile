FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PYTHONUNBUFFERED=1

# Use shell form so $PORT actually gets expanded — Render injects this
# env var at runtime and it can differ from a hardcoded value
CMD gunicorn --bind 0.0.0.0:$PORT main:app