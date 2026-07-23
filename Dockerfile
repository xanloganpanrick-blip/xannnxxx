FROM python:3.11-slim

WORKDIR /app

# Системные зависимости (gcc + libssl-dev нужны для сборки cryptg)
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Python-зависимости (кеширование слоя)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Код приложения
COPY server.py .
COPY сайт.html .

# Временная директория для загружаемых файлов
RUN mkdir -p temp_files && chmod 777 temp_files

# Railway автоматически задаёт переменную PORT и сам пробрасывает трафик.
# EXPOSE носит информационный характер (Railway его игнорирует).
EXPOSE 4545

# Не-root пользователь (рекомендация Railway для безопасности)
RUN useradd --create-home --shell /bin/bash appuser
USER appuser

# python -u = unbuffered output — логи видны в Railway в реальном времени
CMD ["python", "-u", "server.py"]
