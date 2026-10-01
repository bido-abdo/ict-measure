FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY measure_pdf.py measure_service.py ./
ENV PORT=8001
EXPOSE 8001
CMD gunicorn measure_service:app --workers 2 --threads 2 --timeout 180 --bind 0.0.0.0:$PORT
