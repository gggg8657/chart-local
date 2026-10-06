#!/usr/bin/env python3
"""chart local — 표 데이터 → 그래프 후보 여러 개 → 골라서 편집 → 고칠 수 있는 문서(XLSX·PPTX·DOCX·HWPX)로 내보내기.

  python3 app.py                                  # http://localhost:8786
  LLM_API=openai LLM_BASE_URL=http://gpu:8000/v1 LLM_MODEL=Qwen3-32B python3 app.py
  python3 app.py --cli 표.csv pptx 결과.pptx         # 추천 상위 4개를 문서로 (LLM 없이)

흐름: 붙여넣기/CSV/XLSX → 머리글·단위·열 타입 추론(core.py) → 규칙 기반 후보 추천 → SVG 미리보기(svg.py)
      → (선택) 로컬 LLM 이 후보마다 제목·한 줄 설명을 붙임 — 숫자는 만들지 않으며, 표에 없는 숫자가 섞이면 버린다
      → 고른 그래프 편집 → export.py 가 네이티브 차트 문서로 저장 + 다시 열어 구조 검증(verify)
"""
import base64
import datetime
import json
import os
import re
import secrets
import shutil
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
import core  # noqa: E402
import export  # noqa: E402
import samples  # noqa: E402
import svg  # noqa: E402

WS = os.environ.get("WORKSPACE") or os.path.join(ROOT, "_workspace")  # 포털이 AGENT_DATA/<도구> 로 모아 줌
LLM_API = os.environ.get("LLM_API", "ollama")
LLM_BASE = os.environ.get("LLM_BASE_URL", "http://localhost:8000/v1" if LLM_API == "openai" else "http://localhost:11434").rstrip("/")
MODEL = os.environ.get("LLM_MODEL", "qwen3:8b")
LLM_KEY = os.environ.get("LLM_API_KEY", "")
NUM_CTX = int(os.environ.get("NUM_CTX", "8192"))
PORT = int(os.environ.get("PORT", "8786"))
MAX_UPLOAD = 20 * 1024 * 1024
LOCK = threading.Lock()
FORMATS = {"xlsx": "엑셀 (XLSX)", "pptx": "파워포인트 (PPTX)", "docx": "워드 (DOCX)", "hwpx": "한글 (HWPX)"}


def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


# ── LLM ─────────────────────────────────────────────────────────────────
def llm(system, user, model=None, temperature=0.3):
    model = model or MODEL
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]
    try:
        if LLM_API == "openai":
            body = {"model": model, "messages": msgs, "temperature": temperature, "stream": False}
            hdr = {"Content-Type": "application/json", **({"Authorization": f"Bearer {LLM_KEY}"} if LLM_KEY else {})}
            req = urllib.request.Request(LLM_BASE + "/chat/completions", json.dumps(body).encode(), hdr)
            with urllib.request.urlopen(req, timeout=600) as r:
                return json.load(r)["choices"][0]["message"]["content"] or ""
        body = {"model": model, "stream": False, "think": False, "messages": msgs, "format": "json",
                "options": {"temperature": temperature, "num_ctx": NUM_CTX}}
        for attempt in (0, 1):
            try:
                req = urllib.request.Request(LLM_BASE + "/api/chat", json.dumps(body).encode(), {"Content-Type": "application/json"})
                with urllib.request.urlopen(req, timeout=600) as r:
                    j = json.load(r)
                if "error" in j:
                    raise RuntimeError(j["error"])
                return j.get("message", {}).get("content", "")
            except urllib.error.HTTPError as e:
                msg = e.read().decode(errors="replace")
                if attempt == 0 and "think" in msg:
                    body.pop("think")
                    continue
                raise RuntimeError(f"LLM HTTP {e.code}: {msg[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"LLM 서버 연결 실패 ({LLM_BASE}): {e.reason}")
    return ""


def models():
    try:
        if LLM_API == "openai":
            req = urllib.request.Request(LLM_BASE + "/models", headers={"Authorization": f"Bearer {LLM_KEY}"} if LLM_KEY else {})
            with urllib.request.urlopen(req, timeout=3) as r:
                return [m["id"] for m in json.load(r)["data"]]
        with urllib.request.urlopen(LLM_BASE + "/api/tags", timeout=3) as r:
            return [m["name"] for m in json.load(r)["models"]]
    except Exception:
        return []


DESCRIBE_SYS = """너는 연구기관 보고서의 그래프 편집자다. 표 요약과 그래프 후보 목록을 받으면, 후보마다
(1) 보고서 그림 제목(title, 30자 이내, 명사형, 마침표 없음)과
(2) 이 그래프로 무엇을 보여 줄 수 있는지 한 줄 설명(insight, 60자 이내, '~를 보여 줌' '~를 비교' 같은 끝맺음)을 쓴다.
규칙:
- 숫자를 새로 만들거나 계산하지 않는다. 증가율·합계·평균·배수·순위 숫자를 쓰지 않는다. 숫자는 꼭 필요할 때만 요약에 그대로 있는 연도·값을 쓴다.
- 데이터에 없는 원인·해석(예: '정책 효과로')을 지어내지 않는다. 그래프 모양으로 확인 가능한 것만 말한다.
- 열 이름과 단위는 표에 적힌 그대로 쓴다.
- 출력은 JSON 하나: {"c1": {"title": "...", "insight": "..."}, "c2": {...}} — 후보 id 를 그대로 키로."""


_NUMTOK = re.compile(r"\d[\d,]*(?:\.\d+)?")


def allowed_numbers(table):
    ok = set()

    def put(s):
        for t in _NUMTOK.findall(str(s)):
            t = t.replace(",", "")
            ok.add(t)
            try:
                f = float(t)
                ok.add(repr(f).rstrip("0").rstrip("."))
                ok.add(str(int(f)) if f.is_integer() else t)
            except ValueError:
                pass
    for c in table["columns"]:
        put(c["name"])
        put(c.get("unit", ""))
    put(table.get("title", ""))
    for r in table["rows"]:
        for v in r:
            put(v)
    put(len(table["rows"]))
    put(len(table["columns"]))
    return ok


def number_safe(text, ok):
    """LLM 문장 속 숫자가 모두 표에 있는 숫자인지"""
    for t in _NUMTOK.findall(text or ""):
        t = t.replace(",", "")
        if t not in ok and t.rstrip("0").rstrip(".") not in ok:
            return False
    return True


def describe(table, cands, model=None):
    lines = []
    for c in cands:
        ys = ", ".join(c.get("ys") or [])
        lines.append(f"{c['id']}: {core.KINDS.get(c['kind'], c['kind'])} · x={c.get('x') or '-'} · y={ys}"
                     + (f" · 그룹={c['group']}" if c.get("group") else "") + f" — 추천 이유: {c.get('why', '')}")
    user = f"[표 요약]\n{core.summarize(table)}\n\n[그래프 후보]\n" + "\n".join(lines)
    raw = llm(DESCRIBE_SYS, user, model)
    raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.S).strip()
    m = re.search(r"\{.*\}", raw, re.S)
    try:
        data = json.loads(m.group(0)) if m else {}
    except json.JSONDecodeError:
        data = {}
    ok = allowed_numbers(table)
    out, dropped = {}, 0
    ids = {c["id"] for c in cands}
    for k, v in (data.items() if isinstance(data, dict) else []):
        if k not in ids or not isinstance(v, dict):
            continue
        item = {}
        for f, lim in (("title", 40), ("insight", 90)):
            s = re.sub(r"\s+", " ", str(v.get(f) or "")).strip().strip('"')[:lim]
            if s and number_safe(s, ok):
                item[f] = s
            elif s:
                dropped += 1
        if item:
            out[k] = item
    return {"items": out, "dropped": dropped, "model": model or MODEL}


# ── 이력 (WORKSPACE/history/<id>/rec.json + 내보낸 파일) ─────────────────────
def hid_ok(hid):
    if not re.fullmatch(r"\d{8}-\d{6}-[0-9a-f]{4}", hid or ""):
        raise ValueError("잘못된 이력 id")
    return hid


def hdir(hid):
    return os.path.join(WS, "history", hid_ok(hid))


def hist_load(hid):
    return json.loads(read(os.path.join(hdir(hid), "rec.json")))


def hist_save(rec):
    d = hdir(rec["id"])
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, "rec.json")
    with LOCK:
        with open(p + ".tmp", "w", encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False, indent=1)
        os.replace(p + ".tmp", p)


def hist_list(limit=200):
    d = os.path.join(WS, "history")
    if not os.path.isdir(d):
        return []
    out = []
    for name in sorted(os.listdir(d), reverse=True)[:limit]:
        try:
            r = hist_load(name)
            out.append({"id": r["id"], "title": r.get("title", ""), "ts": r.get("ts", ""), "rows": len(r["table"]["rows"]),
                        "cols": len(r["table"]["columns"]), "exports": r.get("exports", [])[-4:], "source": r.get("source", "")})
        except Exception:
            pass
    return out


def hist_delete(hid):
    d = hdir(hid)
    if os.path.isdir(d):
        shutil.rmtree(d)


def new_id():
    return f"{datetime.datetime.now():%Y%m%d-%H%M%S}-{secrets.token_hex(2)}"


# ── 작업 ────────────────────────────────────────────────────────────────
def check_table(t):
    if not isinstance(t, dict) or not isinstance(t.get("columns"), list) or not isinstance(t.get("rows"), list):
        raise ValueError("표 형식이 아닙니다")
    if not t["columns"] or not t["rows"]:
        raise ValueError("표가 비었습니다")
    if len(t["rows"]) > core.MAX_ROWS or len(t["columns"]) > core.MAX_COLS:
        raise ValueError(f"표가 너무 큽니다 (최대 {core.MAX_ROWS}행 × {core.MAX_COLS}열)")
    for c in t["columns"]:
        c["name"] = str(c.get("name") or "").strip() or "열"
        c["type"] = c.get("type") if c.get("type") in ("num", "date", "cat") else "cat"
        c["unit"] = str(c.get("unit") or "").strip()[:12]
    t["rows"] = [[("" if v is None else str(v)) for v in r] for r in t["rows"]]
    t["title"] = str(t.get("title") or "")
    return t


def do_parse(req):
    sheets, sheet = None, None
    if req.get("file_b64"):
        blob = base64.b64decode(req["file_b64"])
        if len(blob) > MAX_UPLOAD:
            raise ValueError("파일이 너무 큽니다 (20MB 이하)")
        name = (req.get("filename") or "").lower()
        if name.endswith((".xlsx", ".xlsm")):
            sheets = core.xlsx_sheets(blob)
            rows, sheet = core.read_xlsx(blob, req.get("sheet"))
        elif name.endswith(".xls"):
            raise ValueError("옛 .xls 형식은 엑셀에서 .xlsx 로 저장한 뒤 올리거나, 셀을 복사해 붙여 넣어 주세요")
        else:
            rows = core.read_csv_bytes(blob)
    else:
        rows = core.read_text(req.get("text") or "")
    if not rows:
        raise ValueError("표 데이터를 찾지 못했습니다")
    table, notes = core.normalize(rows, transpose=req.get("transpose"))
    return {"table": table, "notes": notes, "sheets": sheets, "sheet": sheet}


def do_suggest(req):
    table = check_table(req.get("table"))
    cands = core.recommend(table)
    if not cands:
        raise ValueError("숫자 열이 없어 그래프를 만들 수 없습니다 — 열 타입을 '숫자'로 바꿔 보세요")
    for c in cands:
        c["svg"] = svg.render(table, c, 720, 440)
    hid = req.get("id")
    rec = None
    if hid:
        try:
            rec = hist_load(hid)
        except (FileNotFoundError, ValueError):
            rec = None
    if not rec:
        rec = {"id": new_id(), "exports": []}
    rec.update(ts=datetime.datetime.now().isoformat(timespec="seconds"), table=table,
               title=table.get("title") or (cands[0].get("title") if cands else "") or "표",
               source=req.get("source") or rec.get("source", ""))
    rec["candidates"] = [{k: v for k, v in c.items() if k != "svg"} for c in cands]
    hist_save(rec)
    return {"id": rec["id"], "candidates": cands}


def do_render(req):
    table = check_table(req.get("table"))
    spec = req.get("spec") or {}
    if spec.get("kind") not in core.KINDS:
        raise ValueError("그래프 종류가 없습니다")
    note = ""
    if spec["kind"] in core.CAT_KINDS:
        try:
            note = core.category_data(table, spec)["note"]
        except Exception:
            pass
    return {"svg": svg.render(table, spec, int(req.get("w") or 720), int(req.get("h") or 440)), "note": note}


def safe_name(s):
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", s or "").strip("._")
    return (s or "그래프")[:40]


def do_export(req):
    table = check_table(req.get("table"))
    specs = req.get("specs") or []
    if not specs:
        raise ValueError("내보낼 그래프를 하나 이상 고르세요")
    if len(specs) > 30:
        raise ValueError("한 문서에 그래프는 30개까지")
    fmt_ = req.get("format")
    if fmt_ not in FORMATS:
        raise ValueError("형식은 xlsx|pptx|docx|hwpx")
    if fmt_ == "hwpx" and not export.hwpx_available():
        raise ValueError("이 서버에는 HWPX 생성기(Node + kordoc)가 없습니다 — XLSX·PPTX·DOCX 를 쓰거나 kordoc-local 을 설치하세요")
    for s in specs:
        if s.get("kind") not in core.KINDS:
            raise ValueError("알 수 없는 그래프 종류")
    hid = req.get("id")
    try:
        rec = hist_load(hid) if hid else None
    except (FileNotFoundError, ValueError):
        rec = None
    if not rec:
        rec = {"id": new_id(), "exports": [], "title": req.get("title") or "그래프", "table": table,
               "ts": datetime.datetime.now().isoformat(timespec="seconds")}
    doc_title = (req.get("title") or rec.get("title") or "그래프").strip()[:80]
    d = hdir(rec["id"])
    os.makedirs(d, exist_ok=True)
    fname = f"{safe_name(doc_title)}_{datetime.datetime.now():%m%d_%H%M%S}.{fmt_}"
    path = os.path.join(d, fname)
    report = export.BUILDERS[fmt_](table, specs, path, doc_title=doc_title, include_table=bool(req.get("include_table", True)))
    ver = export.verify(path)
    rec["table"] = table
    rec["selected"] = specs
    rec.setdefault("exports", []).append({"file": fname, "format": fmt_, "ts": datetime.datetime.now().isoformat(timespec="seconds"),
                                          "n": len(specs), "ok": ver["ok"]})
    hist_save(rec)
    return {"id": rec["id"], "file": fname, "url": f"api/file/{rec['id']}/{urllib.parse.quote(fname)}", "report": report,
            "verify": ver, "size": os.path.getsize(path)}


def meta():
    return {"samples": [{"id": s["id"], "label": s["label"]} for s in samples.SAMPLES], "kinds": core.KINDS,
            "formats": [{"id": k, "label": v, "ok": k != "hwpx" or export.hwpx_available()} for k, v in FORMATS.items()],
            "model": MODEL, "palette": svg.PAL}


# ── HTTP ────────────────────────────────────────────────────────────────
HTML = read(os.path.join(ROOT, "ui.html")) if os.path.exists(os.path.join(ROOT, "ui.html")) else "ui.html 없음"

# ── 저작권 표기 (LICENSE·NOTICE 참고) ─────────────────────────────────────
_SIG = __import__("base64").b64decode("wqkgMjAyNiBnZ2dnODY1NyDCtyBkb25nanVraW0uZGV2QGdtYWlsLmNvbQ==").decode()
_SIG_A = __import__("base64").b64decode("Z2dnZzg2NTcgPGRvbmdqdWtpbS5kZXZAZ21haWwuY29tPg==").decode()


def signed(html):
    """화면에 저작권 표기를 붙인다. ui.html 에서 지워져도 서버가 내보낼 때 다시 붙는다."""
    name, mail = _SIG.split(" · ")
    if 'name="author"' not in html:
        meta_ = f'<meta name="author" content="{name[7:]} <{mail}>">'
        html = html.replace("<head>", "<head>" + meta_, 1) if "<head>" in html else meta_ + html
    if "data-sig" not in html:
        tag = (f'<!-- {_SIG} --><div data-sig title="{mail}" style="text-align:center;font-size:11px;color:#9aa0a6;'
               f'opacity:.55;margin:28px 0 8px">{name}</div>')
        html = html.replace("</body>", tag + "</body>", 1) if "</body>" in html else html + tag
    return html


MIME = {"xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "hwpx": "application/hwp+zip"}


class H(BaseHTTPRequestHandler):
    def log_message(self, fmt, *a):
        if "/api/export" in (a[0] if a else ""):
            super().log_message(fmt, *a)

    def _send(self, body, ctype="application/json", code=200, extra=None):
        b = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("X-Author", _SIG_A)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        path = urllib.parse.unquote(self.path.split("?")[0])
        try:
            if path == "/api/health":
                return self._send({"ok": True})
            if path == "/api/meta":
                return self._send(meta())
            if path == "/api/models":
                return self._send(models())
            if path == "/api/history":
                return self._send(hist_list())
            m = re.fullmatch(r"/api/sample/(\w+)", path)
            if m:
                s = next((x for x in samples.SAMPLES if x["id"] == m.group(1)), None)
                if not s:
                    raise FileNotFoundError
                return self._send({"text": s["text"], "label": s["label"]})
            m = re.fullmatch(r"/api/history/([\w-]+)", path)
            if m:
                return self._send(hist_load(m.group(1)))
            m = re.fullmatch(r"/api/file/([\w-]+)/([^/]+)", path)
            if m:
                name = m.group(2)
                if "/" in name or name.startswith(".") or "\\" in name:
                    raise ValueError("잘못된 파일 이름")
                p = os.path.join(hdir(m.group(1)), name)
                ext = name.rsplit(".", 1)[-1].lower()
                with open(p, "rb") as f:
                    blob = f.read()
                return self._send(blob, MIME.get(ext, "application/octet-stream"),
                                  extra={"Content-Disposition": f"attachment; filename*=UTF-8''{urllib.parse.quote(name)}"})
            self._send(signed(HTML).encode(), "text/html; charset=utf-8")
        except (FileNotFoundError, ValueError):
            self._send({"error": "없음"}, code=404)
        except Exception as e:
            self._send({"error": f"{type(e).__name__}: {e}"}, code=500)

    def do_POST(self):
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n > MAX_UPLOAD * 1.4:
                return self._send({"error": "요청이 너무 큽니다"}, code=413)
            req = json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            return self._send({"error": "잘못된 요청"}, code=400)
        routes = {"/api/parse": do_parse, "/api/suggest": do_suggest, "/api/render": do_render, "/api/export": do_export,
                  "/api/describe": lambda r: describe(check_table(r.get("table")), r.get("candidates") or [], r.get("model")),
                  "/api/history/delete": lambda r: (hist_delete(r.get("id")), {"ok": True})[1]}
        fn = routes.get(self.path)
        if not fn:
            return self._send({"error": "없는 경로"}, code=404)
        try:
            self._send(fn(req))
        except (ValueError, RuntimeError) as e:
            self._send({"error": str(e)}, code=400)
        except Exception as e:
            self._send({"error": f"{type(e).__name__}: {e}"}, code=500)


def cli(argv):
    src, fmt_, out = argv[0], argv[1], argv[2]
    with open(src, "rb") as f:
        blob = f.read()
    rows = core.read_xlsx(blob)[0] if src.lower().endswith(".xlsx") else core.read_csv_bytes(blob)
    table, notes = core.normalize(rows)
    specs = core.recommend(table)[:4]
    rep = export.BUILDERS[fmt_](table, specs, out, doc_title=table.get("title") or os.path.basename(src))
    for r in rep:
        print(f"그림 {r['n']}. {r['title']} — {r['label']}")
    v = export.verify(out)
    print("검증:", "통과" if v["ok"] else "; ".join(v["errors"]))


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        cli(sys.argv[2:])
        sys.exit(0)
    print(f"chart local → http://localhost:{PORT}  (llm={LLM_API} {LLM_BASE} {MODEL}, workspace={WS}, "
          f"hwpx={'kordoc' if export.hwpx_available() else '없음'})  {_SIG}")
    ThreadingHTTPServer(("0.0.0.0", PORT), H).serve_forever()
