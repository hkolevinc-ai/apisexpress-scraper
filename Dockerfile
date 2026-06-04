FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY scraper.py .
RUN mkdir -p /data
CMD ["python", "scraper.py", "--output", "/data/apisexpress_products.xlsx"]
