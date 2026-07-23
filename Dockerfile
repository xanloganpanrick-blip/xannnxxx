FROM python:3.11-slim

WORKDIR /app

# Системные зависимости
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Python-зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Код
COPY server.py .

# Временные файлы
RUN mkdir -p temp_files

# Railway передаст свой PORT — не переопределяем!
# В server.py fallback = 4545, если PORT не задан

CMD ["python", "server.py"]
