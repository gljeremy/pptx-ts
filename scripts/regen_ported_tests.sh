#!/bin/sh
# Regenerate the machine-translated test files from a python-pptx checkout.
# Usage: scripts/regen_ported_tests.sh <python-pptx checkout> [python]
# Files listed in test/HAND_EDITED are skipped (they were edited after translation).
set -e
SRC="$1/tests"
PY="${2:-python}"
HERE="$(cd "$(dirname "$0")/.." && pwd)"
for f in $(cat "$HERE/test/PORTED"); do
  rel="${f#tests/}"
  dir="$(dirname "$rel")"
  base="$(basename "$rel" .py)"
  base="${base#test_}"
  case "$(basename "$rel")" in test_*) ext=".test.ts";; *) ext=".ts";; esac
  if [ "$dir" = "." ]; then out="test/$base$ext"; else out="test/$dir/$base$ext"; fi
  if grep -qx "$out" "$HERE/test/HAND_EDITED" 2>/dev/null; then echo "skip $out (hand edited)"; continue; fi
  mkdir -p "$HERE/$(dirname "$out")"
  "$PY" -I "$HERE/scripts/py2ts_tests.py" "$1/$f" "$HERE/$out"
  echo "$out"
done
