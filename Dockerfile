# Hosted PPTX Splitter: LibreOffice renders the slides (there is no PowerPoint
# on a Linux server), ffmpeg turns animated GIFs into MP4s.
FROM python:3.12-slim

# Carlito and Caladea are metric-compatible stand-ins for Calibri and Cambria,
# the Office defaults; without them text reflows and slides look different.
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-impress \
      ffmpeg \
      fonts-crosextra-carlito fonts-crosextra-caladea \
      fonts-liberation2 fonts-dejavu-core fonts-noto-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn==23.0.0

COPY . .

ENV PPTX_SPLITTER_HOSTED=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8000

# One worker: jobs are tracked in that process's memory, so a second worker
# would not know about them. Threads serve pages while a deck converts, and no
# timeout, since a large upload or conversion can take many minutes.
CMD gunicorn --workers 1 --threads 8 --timeout 0 --bind 0.0.0.0:$PORT app:app
