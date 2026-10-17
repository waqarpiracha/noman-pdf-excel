FROM python:3.12-slim

# tesseract OCR engine + English & Urdu language data
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-eng tesseract-ocr-urd \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY app/ ./app/
COPY server.py extract.py format_excel.py ocr_dispatch.py ./

# uploads/output are ephemeral on free hosting — processed per session
RUN mkdir -p uploads output

EXPOSE 10000
CMD ["gunicorn", "server:app", "--bind", "0.0.0.0:10000", "--timeout", "300", "--workers", "2"]
