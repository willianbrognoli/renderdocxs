FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer poppler-utils fontconfig fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /srv
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY builder.py app.py template.docx ./
# Coloque aqui os arquivos de fonte reais (WinnerSans-WideBold.ttf, StageGrotesk-*.otf)
COPY fonts/ /usr/share/fonts/truetype/aguia/
RUN fc-cache -f || true

EXPOSE 8000
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
