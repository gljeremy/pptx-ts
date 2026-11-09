#!/bin/sh
# Regenerate test/features/steps/*.ts from a python-pptx checkout's features/steps/*.py
# (helpers.py is ported by hand: test/features/steps/helpers.ts).
# Usage: scripts/regen_features.sh <python-pptx checkout> [python]
set -e
PY="${2:-python}"
HERE="$(cd "$(dirname "$0")/.." && pwd)"
for f in "$1"/features/steps/*.py; do
  base="$(basename "$f" .py)"
  [ "$base" = helpers ] && continue
  "$PY" -I "$HERE/scripts/py2ts_steps.py" "$f" "$HERE/test/features/steps/$base.ts"
done
