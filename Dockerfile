FROM python:3.11-slim

# System deps: Chrome + Tesseract + fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget gnupg ca-certificates \
    tesseract-ocr \
    fonts-liberation libglib2.0-0 libnss3 libx11-6 libxcb1 libxcomposite1 \
    libxcursor1 libxdamage1 libxext6 libxfixes3 libxi6 libxrandr2 libxrender1 \
    libxss1 libxtst6 libgbm1 libasound2 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libpango-1.0-0 libpangocairo-1.0-0 \
    && wget -q -O /tmp/chrome.deb \
       https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
    && apt-get install -y /tmp/chrome.deb \
    && rm /tmp/chrome.deb \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY *.py ./
COPY config.json ./

VOLUME ["/app/data"]

ENV EQUATION_FILE_PATH=/app/data/RFQ_and_ICT_Equation.xlsx
ENV PORT=8000

EXPOSE 8000

# Long timeout: a full scrape can take 5-10+ minutes; the sync HTTP request
# from the Logic App is held open for its duration.
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-8000} --workers 1 --threads 4 --timeout 1800 app:app"]
