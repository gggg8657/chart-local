#!/usr/bin/env bash
# chart local — 원샷 설치·실행 (macOS / Linux)
#   bash setup.sh          # Python + openpyxl·python-pptx·python-docx(없으면 venv) + LLM 서버 탐색(선택) + selftest + 웹 서버 + 브라우저
#   bash setup.sh stop
# 환경변수: LLM_BASE_URL / LLM_API / LLM_MODEL (제목·설명 제안용, 없어도 동작), PORT (8786), WORKSPACE (이력·내보낸 파일),
#           KORDOC_CLI (HWPX 생성기 kordoc 의 dist/cli.js — 없으면 ../kordoc-local 에서 찾음, 그래도 없으면 HWPX 만 꺼짐)
# selftest 는 스스로 WORKSPACE 를 임시 폴더로 바꿔 돌고 지운다 (실데이터 폴더에 흔적 없음).
# 폐쇄망: wheels/ 폴더에 openpyxl·python-pptx·python-docx 와 의존 휠(lxml, Pillow, XlsxWriter, et-xmlfile, typing-extensions)을
#         넣어 두면 인터넷 없이 설치 (인터넷 되는 PC 에서: pip download -r requirements.txt -d wheels).
set -euo pipefail
PORT="${PORT:-8786}"
if [ -t 1 ]; then B=$'\033[1m'; D=$'\033[2m'; C=$'\033[36m'; G=$'\033[32m'; R=$'\033[31m'; Y=$'\033[33m'; N=$'\033[0m'; else B= D= C= G= R= Y= N=; fi
STEP=0; step() { STEP=$((STEP+1)); printf '  %s[%d/7]%s %s%-14s%s ' "$D" "$STEP" "$N" "$B" "$1" "$N"; }
ok() { printf '%s✔%s %s\n' "$G" "$N" "${1:-}"; }; warn() { printf '%s!%s %s\n' "$Y" "$N" "${1:-}"; }; skip() { printf '%s–%s %s\n' "$D" "$N" "${1:-}"; }
die() { printf '%s✘ %s%s\n\n' "$R" "$*" "$N" >&2; exit 1; }
has() { command -v "$1" >/dev/null 2>&1; }; probe() { curl -fsS -m 2 "$1" >/dev/null 2>&1; }
wait_for() { for _ in $(seq 1 "${2:-30}"); do probe "$1" && return 0; sleep 1; done; return 1; }
printf '\n%s  chart local%s  표 → 그래프 후보 여러 개 → 고칠 수 있는 문서(XLSX·PPTX·DOCX·HWPX)\n\n' "$B" "$N"

step "OS 감지"; case "$(uname -s)" in Darwin*) OS=mac ;; Linux*) OS=linux ;; MINGW*|MSYS*|CYGWIN*) OS=windows ;; *) die "지원하지 않는 OS: $(uname -s)" ;; esac; ok "$OS ($(uname -m))"
step "패키지 확보"; cd "$(dirname "${BASH_SOURCE[0]}")"; [ -f app.py ] || die "app.py 가 없습니다"; ok "$(pwd)"
if [ "${1:-}" = "stop" ]; then [ -f .server.pid ] && kill "$(cat .server.pid)" 2>/dev/null && rm -f .server.pid && ok "웹 서버 종료" || skip "실행 중인 서버 없음"; exit 0; fi

step "Python+문서 라이브러리"
PY=""; for c in python3 python py; do has "$c" && "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' 2>/dev/null && { PY=$c; break; }; done
[ -n "$PY" ] || die "Python 3.9+ 가 없습니다 (sudo apt install python3 / brew install python)"
NEED='import openpyxl, pptx, docx, lxml'
if [ -x venv/bin/python ] && venv/bin/python -c "$NEED" 2>/dev/null; then PY=venv/bin/python
elif ! "$PY" -c "$NEED" 2>/dev/null; then
  [ -x venv/bin/python ] || { if has uv; then uv venv -q --python "$(command -v $PY)" venv; else "$PY" -m venv venv; fi; }
  if [ -d wheels ]; then venv/bin/python -m pip install -q --no-index --find-links wheels -r requirements.txt
  elif has uv; then VIRTUAL_ENV="$PWD/venv" uv pip install -q -r requirements.txt; else venv/bin/python -m pip install -q -r requirements.txt; fi \
    || die "설치 실패 (폐쇄망이면 wheels/ 에 휠을 넣고 재실행 — README 참고)"
  PY=venv/bin/python
fi
ok "$($PY --version 2>&1) + $($PY -c 'import openpyxl, pptx, docx; print(f"openpyxl {openpyxl.__version__} · python-pptx {pptx.__version__} · python-docx {docx.__version__}")')"

step "HWPX (선택)"
if $PY -c 'import sys; sys.path.insert(0,"."); import export; sys.exit(0 if export.hwpx_available() else 1)' 2>/dev/null; then
  ok "kordoc $($PY -c 'import sys; sys.path.insert(0,"."); import export; print(export.kordoc_cli())')"
else warn "Node+kordoc 없음 → HWPX 내보내기만 꺼짐 (kordoc-local 을 설치하거나 KORDOC_CLI=…/dist/cli.js)"; fi
has rsvg-convert || skip "rsvg-convert 없음 → HWPX 의 상자그림·히트맵은 그림 없이 표만"

step "LLM (선택)"
export LLM_API="${LLM_API:-}" LLM_BASE_URL="${LLM_BASE_URL:-}" LLM_MODEL="${LLM_MODEL:-}"
if [ -n "$LLM_BASE_URL" ]; then [ -n "$LLM_API" ] || { case "$LLM_BASE_URL" in *1143*) LLM_API=ollama ;; *) LLM_API=openai ;; esac; }
elif probe http://localhost:11434/api/tags; then LLM_API=ollama LLM_BASE_URL=http://localhost:11434
else for p in 8000 1234 8080; do probe "http://localhost:$p/v1/models" && { LLM_API=openai LLM_BASE_URL="http://localhost:$p/v1"; break; }; done; fi
if [ -n "$LLM_BASE_URL" ] && [ -z "$LLM_MODEL" ]; then
  if [ "$LLM_API" = ollama ]; then LLM_MODEL=$(curl -fsS -m 3 "$LLM_BASE_URL/api/tags" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["models"][0]["name"])' 2>/dev/null || true)
  else LLM_MODEL=$(curl -fsS -m 3 "$LLM_BASE_URL/models" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["data"][0]["id"])' 2>/dev/null || true); fi
fi
if [ -n "$LLM_BASE_URL" ]; then ok "$LLM_API $LLM_BASE_URL ${LLM_MODEL:-}"; else skip "없음 → 그래프 제목·설명 제안만 꺼짐 (나머지는 그대로)"; LLM_API=ollama; fi

step "자가검증"; env -u WORKSPACE "$PY" selftest.py >/dev/null 2>&1 || die "selftest 실패 ($PY selftest.py 로 확인)"; ok "타입 추론·후보 추천·SVG·4개 형식 내보내기+구조 검사 (임시 폴더)"

step "웹 서버"
[ -f .server.pid ] && kill "$(cat .server.pid)" 2>/dev/null || true
LLM_API=$LLM_API LLM_BASE_URL=$LLM_BASE_URL LLM_MODEL=$LLM_MODEL PORT=$PORT nohup "$PY" app.py > server.log 2>&1 & echo $! > .server.pid
wait_for "http://localhost:$PORT/api/health" 20 || { cat server.log; die "웹 서버 기동 실패 (server.log 확인)"; }
URL="http://localhost:$PORT"; ok "$URL"
case "$OS" in mac) open "$URL" ;; linux) [ -n "${WORKSPACE:-}" ] || { has xdg-open && xdg-open "$URL" >/dev/null 2>&1 || true; } ;; esac
printf '\n  %s준비 완료%s  %s%s%s   종료: bash setup.sh stop   로그: server.log\n\n' "$B" "$N" "$C" "$URL" "$N"
