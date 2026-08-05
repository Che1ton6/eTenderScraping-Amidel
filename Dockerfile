FROM python:3.11-slim

# ── System dependencies ───────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    # Chrome for Selenium
    wget gnupg ca-certificates \
    # Tesseract OCR (used by pytesseract)
    tesseract-ocr \
    # Fonts and misc
    fonts-liberation libglib2.0-0 libnss3 libx11-6 libxcb1 libxcomposite1 \
    libxcursor1 libxdamage1 libxext6 libxfixes3 libxi6 libxrandr2 libxrender1 \
    libxss1 libxtst6 libgbm1 libasound2 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libpango-1.0-0 libpangocairo-1.0-0 \
    && wget -q -O /tmp/chrome.deb \
       https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
    && apt-get install -y /tmp/chrome.deb \
    && rm /tmp/chrome.deb \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# ── Python dependencies ───────────────────────────────────────────────────────
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── App source ────────────────────────────────────────────────────────────────
COPY *.py ./
COPY config.json ./

# ── Output volume (data written here at runtime) ──────────────────────────────
VOLUME ["/app/data"]

# ── Equation file lives inside the data mount so it persists across runs ─────
ENV EQUATION_FILE_PATH=/app/data/RFQ_and_ICT_Equation.xlsx

# ── Entry point ───────────────────────────────────────────────────────────────
ENTRYPOINT ["python", "_run_headless.py"]
