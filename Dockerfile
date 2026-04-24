FROM python:3.11-slim

# Install Chrome + required system libs
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget \
        gnupg \
        ca-certificates \
    && wget -qO - https://dl.google.com/linux/linux_signing_key.pub \
        | gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg \
    && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] http://dl.google.com/linux/chrome/deb/ stable main" \
        > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PYTHONUNBUFFERED=1
ENV FLASK_ENV=production
ENV SELENIUM_HEADLESS=true

# Single worker because session state (chrome drivers, oauth store) is in-memory.
# Threads give us concurrency within the worker.
# Long timeout because Selenium batch operations can run for minutes.
CMD gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 4 --timeout 300 backend.app:app
