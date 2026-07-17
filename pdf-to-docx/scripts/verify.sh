#!/usr/bin/env bash
# Round-trip: docx -> pdf -> jpg, so the agent can visually diff the rendered
# result against the original page images. If the original PDF is passed,
# also validates that the rendered page count matches it.
#
# Usage: bash verify.sh output.docx [original.pdf] [checkdir]
set -euo pipefail

if [ $# -lt 1 ]; then
  echo "usage: bash verify.sh output.docx [original.pdf] [checkdir]" >&2
  exit 2
fi

DOCX="$1"
ORIGINAL="${2:-}"
CHECKDIR="${3:-$(dirname "$DOCX")/verify}"
mkdir -p "$CHECKDIR"

RUN_TIMEOUT=""
command -v timeout >/dev/null 2>&1 && RUN_TIMEOUT="timeout 120"

# Isolated profile avoids the LibreOffice user-profile lock (headless
# conversion hangs silently if another instance ran/crashed) and permits
# parallel conversions.
$RUN_TIMEOUT soffice --headless -env:UserInstallation=file:///tmp/lo_verify_profile \
  --convert-to pdf --outdir "$CHECKDIR" "$DOCX" >/dev/null

BASE="$(basename "${DOCX%.docx}")"
RENDERED_PDF="$CHECKDIR/$BASE.pdf"
if [ ! -f "$RENDERED_PDF" ]; then
  echo "verify.sh: soffice produced no PDF for $DOCX" >&2
  exit 1
fi

page_count() {
  pdfinfo "$1" | awk '/^Pages:/ {print $2}'
}

if [ -n "$ORIGINAL" ]; then
  ORIG_PAGES="$(page_count "$ORIGINAL")"
  RENDER_PAGES="$(page_count "$RENDERED_PDF")"
  if [ "$ORIG_PAGES" != "$RENDER_PAGES" ]; then
    echo "verify.sh: PAGE COUNT MISMATCH — original has $ORIG_PAGES page(s), rendered docx has $RENDER_PAGES" >&2
    echo "verify.sh: content is overflowing or underfilling pages; fix the renderer" >&2
    exit 1
  fi
  echo "Page count OK: $RENDER_PAGES page(s), matches original"
fi

pdftoppm -jpeg -r 100 "$RENDERED_PDF" "$CHECKDIR/$BASE-render"

echo "Rendered pages:"
ls -1 "$CHECKDIR/$BASE-render"*.jpg
