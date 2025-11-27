# Excel Diff Tool Dockerfile
# 基于 Python 3.13 slim 镜像

FROM python:3.13-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# 安装系统依赖（如果需要的话）
# RUN apt-get update && apt-get install -y --no-install-recommends \
#     && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制项目文件
COPY app.py .
COPY static/ ./static/

# 暴露端口
EXPOSE 0731

# 启动命令
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "0731"]