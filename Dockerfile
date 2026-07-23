FROM python:3.11-slim

WORKDIR /app

# Системные зависимости: gcc+libssl-dev для cryptg, curl для HEALTHCHECK
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libssl-dev \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Python-зависимости (кеширование слоя)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Код приложения
COPY server.py .
COPY сайт.html .

# Временная директория для загружаемых файлов
RUN mkdir -p temp_files && chmod 777 temp_files

# Railway задаёт $PORT — server.py читает: os.environ.get("PORT", 4545)
# Railway сам маршрутизирует трафик и делает health-check на платформенном уровне
EXPOSE 4545

# python -u = unbuffered output — логи сразу видны в Railway
CMD ["python", "-u", "server.py"]
