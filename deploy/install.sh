#!/usr/bin/env bash
# Install or update the PPTX Splitter on the server. Run as root:
#   bash install.sh
# Touches only: apt packages (LibreOffice, ffmpeg, fonts, python3-venv), the
# pptx-splitter user, /opt/pptx-splitter, /var/lib/pptx-splitter and the
# pptx-splitter systemd service. It never touches PM2, nginx or other apps;
# the nginx block is added by hand (see nginx-splitter.conf).
set -euo pipefail

REPO=https://github.com/SherlyGuides/pptx-splitter.git
APP=/opt/pptx-splitter
DATA=/var/lib/pptx-splitter

if ss -ltn 2>/dev/null | grep -q '127.0.0.1:5055 ' && ! systemctl is-active --quiet pptx-splitter; then
  echo "Port 5055 is already used by something else; pick another port." >&2
  exit 1
fi

export DEBIAN_FRONTEND=noninteractive
apt-get update -q
apt-get install -y -q --no-install-recommends \
  git python3-venv libreoffice-impress ffmpeg \
  fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation2 fonts-dejavu-core

id pptx-splitter >/dev/null 2>&1 || useradd --system --home-dir "$DATA" --shell /usr/sbin/nologin pptx-splitter
mkdir -p "$DATA/tmp"
chown -R pptx-splitter:pptx-splitter "$DATA"

if [ -d "$APP/.git" ]; then
  git -C "$APP" pull --ff-only
else
  git clone "$REPO" "$APP"
fi
python3 -m venv "$APP/.venv"
"$APP/.venv/bin/pip" install -q -r "$APP/requirements.txt" gunicorn==23.0.0

install -m 644 "$APP/deploy/pptx-splitter.service" /etc/systemd/system/pptx-splitter.service
systemctl daemon-reload
systemctl enable --now pptx-splitter
systemctl restart pptx-splitter
sleep 3
curl -fsS http://127.0.0.1:5055/healthz && echo "  <- splitter is up on 127.0.0.1:5055"
