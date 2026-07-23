FROM python:3.11-slim

WORKDIR /app

# Устанавливаем системные зависимости (если нужны)
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Копируем и устанавливаем Python-зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем код приложения
COPY server.py .

# Создаём папку для временных файлов
RUN mkdir -p temp_files

# Порт
EXPOSE 8765

# Запуск
CMD ["python", "server.py"]
