#!/usr/bin/env bash
# Сборка PDF из HTML. Слайд 1280×720.
set -uo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT="${1:-$DIR/Wagon-Park-Kazakhstan.pdf}"
CHROME="$(command -v google-chrome || command -v google-chrome-stable || command -v chromium)"
timeout 300 "$CHROME" \
  --headless=new \
  --disable-gpu \
  --no-sandbox \
  --disable-dev-shm-usage \
  --user-data-dir="$(mktemp -d)" \
  --no-pdf-header-footer \
  --virtual-time-budget=25000 \
  --print-to-pdf="$OUT" \
  "file://$DIR/index.html" >/dev/null 2>&1
if [[ -s "$OUT" ]]; then
  echo "PDF готов: $OUT"
else
  echo "Ошибка: PDF не создан" >&2
  exit 1
fi
