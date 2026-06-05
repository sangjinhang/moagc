FROM python:3.11-slim

WORKDIR /app

# 系统依赖（字体、中文字体）
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-wqy-zenhei fonts-wqy-microhei \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 暴露端口
EXPOSE 8080

# 启动
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "2", "--threads", "4", "--timeout", "600", "app:app"]
