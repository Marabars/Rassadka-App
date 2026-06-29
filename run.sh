#!/bin/bash
set -e
if [ -f .env ]; then
  export $(grep -v '^#' .env | xargs)
fi
python -m uvicorn main:app \
  --host "${HOST:-0.0.0.0}" \
  --port "${PORT:-8002}" \
  --workers 1
