FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY qumulo_rack/ ./qumulo_rack/
COPY web/ ./web/
COPY qrack.py .
COPY docker-entrypoint.sh .
RUN chmod +x docker-entrypoint.sh

RUN useradd --create-home --uid 1000 qrack && chown -R qrack:qrack /app
USER qrack

EXPOSE 8000 8443

CMD ["./docker-entrypoint.sh"]
