# Playwright's official image ships Chromium + all OS deps, matched to our
# playwright version so browser launches work out of the box on Render/Railway.
FROM mcr.microsoft.com/playwright/python:v1.60.0-jammy

WORKDIR /app

# Install Python deps first for better layer caching.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# in-memory job state is shared across requests, so run a single worker with
# multiple threads. --timeout 0 keeps long-lived SSE streams from being killed.
ENV PORT=5000
ENV ANALYSIS_CONCURRENCY=2
ENV CRAWL_MAX_PAGES=1000
ENV SSE_HEARTBEAT_SECONDS=20
EXPOSE 5000
CMD gunicorn -k gthread -w 1 --threads 8 --timeout 0 -b 0.0.0.0:${PORT} app:app
