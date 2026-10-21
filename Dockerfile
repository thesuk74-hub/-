# 공사기간 적정성 검토 웹 — 사내 서버 배포용
FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer fonts-nanum && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV CPA_WEB_DATA=/data
VOLUME /data
EXPOSE 8000
CMD ["python", "-m", "cpa", "web", "--host", "0.0.0.0", "--port", "8000"]
