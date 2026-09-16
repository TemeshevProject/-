#!/usr/bin/env bash
# Рендер презентации из HTML в PDF (слайды 1280x720, формат 16:9).
set -uo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT="${1:-$DIR/Rail-Logistics-Hub-Astana.pdf}"

CHROME="$(command -v google-chrome || command -v google-chrome-stable || command -v chromium)"

# Chrome в headless-режиме не всегда завершает процесс после записи файла,
# поэтому выполнение ограничено таймаутом, а успех проверяется по наличию PDF.
timeout 90 "$CHROME" \
  --headless=new \
  --disable-gpu \
  --no-sandbox \
  --disable-dev-shm-usage \
  --user-data-dir="$(mktemp -d)" \
  --no-pdf-header-footer \
  --virtual-time-budget=8000 \
  --print-to-pdf="$OUT" \
  "file://$DIR/index.html" >/dev/null 2>&1

if [[ -s "$OUT" ]]; then
  echo "PDF готов: $OUT"
else
  echo "Ошибка: PDF не создан" >&2
  exit 1
fi
