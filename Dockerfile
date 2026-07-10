# ============================================================
# OnlyOffice 离线编辑器 — Docker 镜像
# 构建: docker build -t onlyoffice-offline .
# 运行: docker run -p 5000:5000 -v ./uploads:/app/uploads onlyoffice-offline
# ============================================================
FROM python:3.12-slim

WORKDIR /app

# 安装 Python 依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 复制应用代码
COPY app.py .
COPY templates/ templates/

# 复制 OnlyOffice SDK 静态资源（1GB，确保已从项目复制）
COPY static/packages/ static/packages/

# 上传目录
RUN mkdir -p uploads

EXPOSE 5000

CMD ["gunicorn", "-b", "0.0.0.0:5000", "-w", "2", "--timeout", "120", "app:app"]
