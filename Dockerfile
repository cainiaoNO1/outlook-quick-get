FROM python:3.11-slim

# 安装 Node.js 和基础运行依赖
RUN apt-get update && apt-get install -y --no-install-recommends \
    nodejs \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 安装 Python 运行依赖
RUN pip install --no-cache-dir cryptography requests

# 拷贝代码
COPY . /app

# 确保无残留敏感/缓存文件并设置权限
RUN rm -f tool/account.txt tool/.token_cache.json config.json && \
    cp config.example.json config.json

ENV PYTHONUNBUFFERED=1 \
    DSH_MAIL_HOST=0.0.0.0 \
    PORT=4000

EXPOSE 4000

CMD ["python", "server.py"]
