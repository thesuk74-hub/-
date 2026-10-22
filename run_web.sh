#!/usr/bin/env bash
# 공사기간 적정성 검토 웹 실행 (macOS/Linux)
cd "$(dirname "$0")"
python3 -m pip install -q -r requirements.txt
echo "브라우저에서 http://localhost:${PORT:-8000} 접속 (종료: Ctrl+C)"
python3 -m cpa web --host "${HOST:-127.0.0.1}" --port "${PORT:-8000}" --open
