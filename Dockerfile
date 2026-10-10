FROM python:3.11-slim

# LibreOffice(Word→PDF) + 中文字体(保证中文转换不乱码)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY app.py .
COPY converter/ converter/
COPY templates/ templates/

# 非 root 运行，降低解析恶意文件被攻破后的影响
RUN useradd --create-home --uid 1000 appuser \
    && mkdir -p /app/storage && chown -R appuser:appuser /app
USER appuser

# 服务器模式：监听 0.0.0.0:7860，不弹浏览器
ENV HOST=0.0.0.0 PORT=7860 PDFWORD_NO_BROWSER=1
EXPOSE 7860

CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "1", "--threads", "8", "--timeout", "300", "app:app"]
