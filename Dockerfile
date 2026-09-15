FROM python:3.12-slim

WORKDIR /app

# libreoffice-impress (not the full libreoffice suite) renders the .pptx
# preview image server-side, from the exact same file a user would
# download -- no separate layout implementation to keep in sync. Our slide
# can be appended after a user-supplied template's own slides, and soffice's
# PNG export only ever rasterizes page 1 of a multi-page conversion, so the
# preview goes via PDF (every page) + poppler-utils' pdftoppm to pull out
# the one page we actually want.
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-impress poppler-utils \
    && rm -rf /var/lib/apt/lists/*

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
