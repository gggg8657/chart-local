"""chart local — 표 읽기·열 타입 추론·그래프 후보 추천·차트용 데이터 준비 (stdlib 만)

표(table) = {"columns": [{"name", "type": num|date|cat, "unit"}], "rows": [[문자열…]…], "title": 제목 후보}
그래프 명세(spec) = {"kind", "x", "ys": [...], "group", "agg", 스타일…} — 렌더러(svg.py)·내보내기(export.py) 공통 입력.
"""
import csv
import io
import math
import re

MAX_ROWS = 5000
MAX_COLS = 60

# ── 숫자·단위 ──────────────────────────────────────────────────────────────
_NUM = re.compile(r"^([+\-−–]?)\s*([₩$€¥£]?)\s*(\d{1,3}(?:,\d{3})+|\d+)?(\.\d+)?(?:[eE]([+\-]?\d+))?\s*(.*?)$")
_UNIT_OK = re.compile(r"^(%|‰|[A-Za-zµμ°Ω/²³·.\-]{1,10}|[가-힣]{1,4}|[A-Za-z]+/[A-Za-z가-힣]+|천?[만억조]?원|[천만억조]+[가-힣]{0,2})$")
_EMPTY = {"", "-", "–", "—", "n/a", "N/A", "NA", "na", "null", "None", "nan", "NaN", "#N/A", "#DIV/0!", "."}


def split_num(s):
    """'1,234.5억원' → (1234.5, '억원'); 숫자가 아니면 (None, None). 괄호 음수 (123) 지원."""
    if s is None:
        return None, None
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return (float(s), "") if math.isfinite(s) else (None, None)
    t = str(s).strip().replace(" ", " ")
    if t in _EMPTY:
        return None, None
    neg = False
    if len(t) > 2 and t[0] == "(" and t[-1] == ")":
        neg, t = True, t[1:-1].strip()
    if t.startswith("△") or t.startswith("▽"):  # 공문서 음수 표기
        neg, t = True, t[1:].strip()
    m = _NUM.match(t)
    if not m or not (m.group(3) or m.group(4)):
        return None, None
    sign, _cur, ip, fp, ex, rest = m.groups()
    rest = rest.strip()
    if rest and not _UNIT_OK.match(rest):
        return None, None
    v = float((ip or "0").replace(",", "") + (fp or "") + (("e" + ex) if ex else ""))
    if sign in "-−–" and sign:
        v = -v
    if neg:
        v = -v
    return v, rest


def num(s):
    return split_num(s)[0]


# ── 날짜·기간 ──────────────────────────────────────────────────────────────
_DATE_PATS = [
    re.compile(r"^(19|20)\d{2}$"),
    re.compile(r"^(19|20)\d{2}\s*년(도)?$"),
    re.compile(r"^(19|20)\d{2}[-./]\s?\d{1,2}([-./]\s?\d{1,2})?\.?$"),
    re.compile(r"^(19|20)\d{2}\s*년\s*\d{1,2}\s*월(\s*\d{1,2}\s*일)?$"),
    re.compile(r"^\d{1,2}\s*월$"),
    re.compile(r"^(19|20)?\d{2}\s*[-./]?\s*[1-4]\s*(Q|분기)$", re.I),
    re.compile(r"^((19|20)\d{2}\s*)?[1-4]\s*분기$"),
    re.compile(r"^((19|20)\d{2}\s*)?Q[1-4]$", re.I),
    re.compile(r"^((19|20)\d{2}\s*)?(상|하)반기$"),
    re.compile(r"^(19|20)\d{2}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2})?$"),
]
_TIME_NAME = re.compile(r"(연도|년도|년|월|일자|날짜|기간|분기|반기|시점|date|year|month|quarter|time)", re.I)


def is_date(s):
    t = str(s).strip()
    return any(p.match(t) for p in _DATE_PATS)


def date_key(s):
    """정렬용 키 — 숫자만 뽑아 튜플로. 분기·반기는 그 순서대로."""
    t = str(s)
    nums = [int(x) for x in re.findall(r"\d+", t)]
    if "하반기" in t:
        nums.append(2)
    elif "상반기" in t:
        nums.append(1)
    return tuple(nums)


# ── 표 읽기 ──────────────────────────────────────────────────────────────
def read_text(text):
    """붙여 넣은 텍스트(엑셀 탭 구분·CSV·세미콜론) → 2차원 문자열 배열"""
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    if not text.strip():
        return []
    lines = [l for l in text.split("\n")]
    if any("\t" in l for l in lines[:20]):
        delim = "\t"
    else:
        sample = "\n".join(lines[:20])
        counts = {d: min((l.count(d) for l in lines[:20] if l.strip()), default=0) for d in (",", ";", "|")}
        delim = max(counts, key=counts.get) if max(counts.values()) > 0 else None
        if delim == "," and re.search(r"\d,\d{3}", sample) and "\"" not in sample and counts.get(";", 0) == 0:
            # 천단위 쉼표 숫자가 섞인 '쉼표 구분' — csv 모듈이 따옴표 없는 1,234 를 쪼갠다: 공백 2개 이상 구분 시도
            if all(re.search(r"\S\s{2,}\S", l) for l in lines[:5] if l.strip()):
                delim = "  "
    if delim is None:  # 공백 2칸 이상 정렬된 표 또는 한 열
        return [re.split(r"\s{2,}", l.strip()) for l in lines if l.strip()]
    if delim == "  ":
        return [re.split(r"\s{2,}", l.strip()) for l in lines if l.strip()]
    return [row for row in csv.reader(io.StringIO(text), delimiter=delim)]


def xlsx_sheets(blob):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def read_xlsx(blob, sheet=None):
    import datetime
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        out = []
        for row in ws.iter_rows(values_only=True):
            r = []
            for v in row[:MAX_COLS]:
                if v is None:
                    r.append("")
                elif isinstance(v, datetime.datetime):
                    r.append(v.strftime("%Y-%m-%d") if not (v.hour or v.minute) else v.strftime("%Y-%m-%d %H:%M"))
                elif isinstance(v, datetime.date):
                    r.append(v.isoformat())
                elif isinstance(v, float):
                    r.append(repr(round(v, 10)).rstrip("0").rstrip(".") if v != int(v) or abs(v) > 1e15 else str(int(v)))
                else:
                    r.append(str(v))
            out.append(r)
            if len(out) > MAX_ROWS + 10:
                break
        return out, ws.title
    finally:
        wb.close()


def read_csv_bytes(blob):
    for enc in ("utf-8-sig", "cp949", "euc-kr", "latin-1"):
        try:
            return read_text(blob.decode(enc))
        except UnicodeDecodeError:
            continue
    return []


# ── 정규화: 머리글·단위 행·제목 행·전치 ────────────────────────────────────
_UNIT_CELL = re.compile(r"^[\(\[〔<]?\s*(단위\s*[:：]?\s*)?([^\d\s\(\)\[\]]{1,10})\s*[\)\]〕>]?$")
_HEAD_UNIT = re.compile(r"^(.*?)\s*[\(\[〔]\s*(?:단위\s*[:：]?\s*)?([^\(\)\[\]〔〕]{1,12})\s*[\)\]〕]\s*$")
_GLOBAL_UNIT = re.compile(r"[\(\[]?\s*단위\s*[:：]\s*([^\)\]\s,]{1,10})\s*[\)\]]?")


def _numlike(c):
    return num(c) is not None


def normalize(rows, transpose=None):
    """원시 2차원 배열 → table + notes(무엇을 자동 처리했는지 한국어 설명)"""
    notes = []
    rows = [[("" if c is None else str(c)).strip() for c in r] for r in rows]
    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        raise ValueError("표 데이터가 없습니다")
    width = min(MAX_COLS, max(len(r) for r in rows))
    rows = [(r + [""] * width)[:width] for r in rows]
    keep = [j for j in range(width) if any(r[j] for r in rows)]
    rows = [[r[j] for j in keep] for r in rows]
    width = len(keep)
    title, gunit = "", ""
    # 제목·단위 줄: 칸이 하나뿐인 앞쪽 줄
    while len(rows) > 2 and sum(1 for c in rows[0] if c) == 1 and sum(1 for c in rows[1] if c) >= 2 and width >= 2:
        only = next(c for c in rows[0] if c)
        gm = _GLOBAL_UNIT.search(only)
        if gm:
            gunit = gm.group(1)
            rest = _GLOBAL_UNIT.sub("", only).strip(" -·:")
            if rest and not title:
                title = rest
        elif not title:
            title = only
        rows = rows[1:]
        notes.append(f"첫 줄 '{only[:30]}' 을(를) {'단위' if gm else '제목'}로 사용")
    # 머리글 판단
    head = rows[0]
    body = rows[1:]
    header = True
    if body:
        num_cols = [j for j in range(width) if sum(_numlike(r[j]) for r in body) >= max(1, 0.7 * sum(1 for r in body if r[j]))]
        head_num = sum(1 for j in num_cols if _numlike(head[j]) and not is_date(head[j]))
        head_txt = sum(1 for c in head if c and not _numlike(c))
        if num_cols and head_num == len(num_cols) and head_txt == 0:
            header = False
    if not body:
        header = False
    if header:
        names = [c or f"열{j + 1}" for j, c in enumerate(head)]
        rows = body
    else:
        names = [f"열{j + 1}" for j in range(width)]
        notes.append("머리글이 없어 보여 열1·열2… 로 이름 붙임")
    # 머리글 바로 아래 단위 행
    units = [""] * width
    if len(rows) > 2 and all((not c) or (_UNIT_CELL.match(c) and not _numlike(c)) for c in rows[0]) \
            and any(rows[0]) and any(_numlike(c) for c in rows[1]):
        for j, c in enumerate(rows[0]):
            m = _UNIT_CELL.match(c) if c else None
            units[j] = m.group(2).strip() if m else ""
        rows = rows[1:]
        notes.append("머리글 아래 단위 행을 열 단위로 사용")
    # 머리글 안 단위: '예산(억원)'
    for j, n in enumerate(names):
        m = _HEAD_UNIT.match(n)
        if m and m.group(1).strip():
            names[j], units[j] = m.group(1).strip(), units[j] or m.group(2).strip()
    names = _dedupe(names)
    if len(rows) > MAX_ROWS:
        notes.append(f"행이 많아 앞 {MAX_ROWS}행만 사용")
        rows = rows[:MAX_ROWS]
    table = {"columns": [{"name": n, "unit": u, "type": "cat"} for n, u in zip(names, units)], "rows": rows, "title": title}
    infer_types(table, gunit)
    # 연도가 머리글에 가로로 늘어선 표 → 전치 (첫 열이 항목 이름이고 나머지 머리글이 연도·기간)
    cols = table["columns"]
    wide = (header and len(cols) >= 4 and cols[0]["type"] == "cat" and all(is_date(c["name"]) for c in cols[1:])
            and all(c["type"] == "num" for c in cols[1:]) and len(rows) <= 12)
    if transpose is True or (transpose is None and wide):
        table = transpose_table(table)
        notes.append("기간이 열 머리글에 가로로 놓여 있어 행·열을 바꿈" if transpose is None else "행·열을 바꿈")
    return table, notes


def transpose_table(table):
    cols, rows = table["columns"], table["rows"]
    first = cols[0]
    names = [first["name"] or "구분"] + [r[0] or f"행{i + 1}" for i, r in enumerate(rows)]
    unit = next((c["unit"] for c in cols[1:] if c["unit"]), "")
    new_rows = [[c["name"]] + [r[j + 1] if j + 1 < len(r) else "" for r in rows] for j, c in enumerate(cols[1:])]
    t = {"columns": [{"name": n, "unit": "" if i == 0 else unit, "type": "cat"} for i, n in enumerate(_dedupe(names))],
         "rows": new_rows, "title": table.get("title", "")}
    infer_types(t)
    return t


def _dedupe(names):
    seen, out = {}, []
    for n in names:
        k = n
        while k in seen:
            seen[n] = seen.get(n, 1) + 1
            k = f"{n}_{seen[n]}"
        seen[k] = 1
        out.append(k)
    return out


def infer_types(table, gunit=""):
    """열 타입(num|date|cat)과 단위 추론. 값 뒤의 공통 접미사(%, kg, 억원 …)는 단위로 옮긴다."""
    rows = table["rows"]
    for j, c in enumerate(table["columns"]):
        vals = [r[j] for r in rows if j < len(r) and str(r[j]).strip() not in _EMPTY]
        if not vals:
            c["type"] = "cat"
            continue
        parsed = [split_num(v) for v in vals]
        nnum = sum(1 for v, _ in parsed if v is not None)
        dates = sum(1 for v in vals if is_date(v))
        sufs = {u for v, u in parsed if v is not None and u}
        if dates >= 0.8 * len(vals) and not (nnum == len(vals) and not _TIME_NAME.search(c["name"]) and not all(re.fullmatch(r"(19|20)\d{2}", v.strip()) for v in vals)):
            c["type"] = "date"
        elif nnum >= 0.8 * len(vals):
            ints = [v for v, _ in parsed if v is not None]
            yearish = all(float(v).is_integer() and 1900 <= v <= 2100 for v in ints)
            if yearish and (_TIME_NAME.search(c["name"]) or (j == 0 and len(set(ints)) == len(ints))):
                c["type"] = "date"
            else:
                c["type"] = "num"
                if len(sufs) == 1 and not c.get("unit"):
                    c["unit"] = sufs.pop()
                if not c.get("unit") and gunit:
                    c["unit"] = gunit
        else:
            c["type"] = "cat"
    return table


def col_index(table, name):
    for j, c in enumerate(table["columns"]):
        if c["name"] == name:
            return j
    raise ValueError(f"열 '{name}' 이(가) 없습니다")


def col_values(table, name):
    j = col_index(table, name)
    return [r[j] if j < len(r) else "" for r in table["rows"]]


def col_nums(table, name):
    return [num(v) for v in col_values(table, name)]


def unit_of(table, name):
    try:
        return table["columns"][col_index(table, name)].get("unit", "")
    except ValueError:
        return ""


# ── 통계 ──────────────────────────────────────────────────────────────────
def quantile(sorted_vals, q):
    """엑셀 QUARTILE.INC 와 같은 선형 보간"""
    n = len(sorted_vals)
    if n == 1:
        return sorted_vals[0]
    pos = (n - 1) * q
    lo = int(math.floor(pos))
    hi = min(lo + 1, n - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (pos - lo)


def box_stats(vals):
    v = sorted(x for x in vals if x is not None)
    if not v:
        return None
    q1, med, q3 = quantile(v, .25), quantile(v, .5), quantile(v, .75)
    iqr = q3 - q1
    lo_f, hi_f = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    inside = [x for x in v if lo_f <= x <= hi_f] or v
    return {"min": v[0], "q1": q1, "median": med, "q3": q3, "max": v[-1], "wlo": inside[0], "whi": inside[-1],
            "out": [x for x in v if x < lo_f or x > hi_f], "n": len(v), "mean": sum(v) / len(v)}


def linreg(xs, ys):
    pts = [(x, y) for x, y in zip(xs, ys) if x is not None and y is not None]
    n = len(pts)
    if n < 3:
        return None
    mx = sum(p[0] for p in pts) / n
    my = sum(p[1] for p in pts) / n
    sxx = sum((p[0] - mx) ** 2 for p in pts)
    syy = sum((p[1] - my) ** 2 for p in pts)
    sxy = sum((p[0] - mx) * (p[1] - my) for p in pts)
    if sxx == 0:
        return None
    a = sxy / sxx
    b = my - a * mx
    r2 = (sxy * sxy) / (sxx * syy) if syy else 1.0
    return {"slope": a, "intercept": b, "r2": r2, "r": (sxy / math.sqrt(sxx * syy)) if syy else 1.0}


def nice_step(span, target=5):
    if span <= 0:
        return 1
    raw = span / target
    p = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * p:
            return m * p
    return 10 * p


def hist_bins(vals, nbins=None):
    """깔끔한 구간 경계(1·2·2.5·5 배수) — [(lo, hi, count)]. 마지막 구간은 오른쪽 끝 포함."""
    v = sorted(x for x in vals if x is not None)
    if not v:
        return []
    lo, hi = v[0], v[-1]
    if lo == hi:
        return [(lo, hi, len(v))]
    k = nbins or max(5, min(20, int(math.ceil(math.log2(len(v)) + 1))))
    step = nice_step(hi - lo, k)
    start = math.floor(lo / step) * step
    edges = [start]
    while edges[-1] <= hi - 1e-12 * max(1, abs(hi)):
        edges.append(round(edges[-1] + step, 10))
    if len(edges) < 2:
        edges.append(start + step)
    out = []
    for i in range(len(edges) - 1):
        a, b = edges[i], edges[i + 1]
        last = i == len(edges) - 2
        out.append((a, b, sum(1 for x in v if a <= x < b or (last and x == b))))
    return out


# ── 서식 ──────────────────────────────────────────────────────────────────
def decimals_of(vals, cap=2):
    d = 0
    for v in vals:
        if v is None or not math.isfinite(v):
            continue
        s = f"{abs(v):.6f}".rstrip("0")
        frac = s.split(".")[1] if "." in s else ""
        d = max(d, len(frac))
    return min(d, cap)


def fmt(v, dec=None, unit=""):
    if v is None:
        return ""
    if dec is None:
        dec = decimals_of([v])
    a = abs(v)
    if a >= 1e12 and unit not in ("%",):
        s = f"{v / 1e12:,.{1 if a < 1e13 else 0}f}조"
    elif a >= 1e8 and unit not in ("%",):
        s = f"{v / 1e8:,.{1 if a < 1e9 else 0}f}억"
    else:
        s = f"{v:,.{dec}f}"
    if s.startswith("-"):
        s = "−" + s[1:]
    return s + ("%" if unit == "%" else "")


# ── 차트 데이터 준비 ───────────────────────────────────────────────────────
KINDS = {
    "col": "세로 막대", "bar": "가로 막대", "col_group": "묶은 세로 막대", "bar_group": "묶은 가로 막대",
    "col_stack": "누적 세로 막대", "col_stack100": "100% 누적 막대", "bar_stack": "누적 가로 막대",
    "line": "꺾은선", "area": "영역", "area_stack": "누적 영역", "scatter": "산점도", "pie": "원그래프",
    "doughnut": "도넛", "hist": "히스토그램", "box": "상자그림", "heatmap": "히트맵", "panels": "단위별 2단 비교",
    "index": "지수 비교(첫 값=100)", "dual": "이중축 (보조 y축)",
}
CAT_KINDS = {"dual", "col", "bar", "col_group", "bar_group", "col_stack", "col_stack100", "bar_stack", "line", "area",
             "area_stack", "pie", "doughnut", "heatmap", "panels", "index"}


def visible_ys(spec):
    hidden = set(spec.get("hidden") or [])
    return [y for y in (spec.get("ys") or []) if y not in hidden]


def category_data(table, spec):
    """범주형 차트 공통: x 순서대로 cats + 계열별 값. group 이 있으면 긴 형식 → 넓은 형식 피벗.
    중복된 x 는 spec.agg(sum|mean) 로 묶는다. 반환: {cats, series:[{name, values, unit, idx}], agg_note}"""
    x = spec.get("x")
    ys = list(spec.get("ys") or [])
    if spec.get("kind") == "dual" and spec.get("y2") and spec["y2"] not in ys:
        ys.append(spec["y2"])  # 이중축: 왼쪽 축 계열(ys) + 오른쪽 축 계열(y2)
    xs = col_values(table, x) if x else [str(i + 1) for i in range(len(table["rows"]))]
    agg = spec.get("agg") or "sum"
    order, seen = [], set()
    for v in xs:
        if v not in seen:
            seen.add(v)
            order.append(v)
    xtype = table["columns"][col_index(table, x)]["type"] if x else "cat"
    if xtype == "date":
        try:
            order.sort(key=date_key)
        except TypeError:
            pass
    note = ""
    series = []
    if spec.get("group"):
        y = ys[0]
        gs = col_values(table, spec["group"])
        vals = col_nums(table, y)
        levels = []
        for g in gs:
            if g not in levels:
                levels.append(g)
        acc = {}
        for xv, g, v in zip(xs, gs, vals):
            if v is None:
                continue
            acc.setdefault((xv, g), []).append(v)
        if any(len(a) > 1 for a in acc.values()):
            note = f"같은 ({x}, {spec['group']}) 조합이 여러 행이라 {'평균' if agg == 'mean' else '합계'}로 묶음"
        hidden = set(spec.get("hidden") or [])
        for i, g in enumerate(levels):
            vs = []
            for c in order:
                a = acc.get((c, g))
                vs.append(None if not a else (sum(a) / len(a) if agg == "mean" else sum(a)))
            series.append({"name": g or "(빈 값)", "values": vs, "unit": unit_of(table, y), "idx": i, "hidden": g in hidden})
    else:
        pos = {c: i for i, c in enumerate(order)}
        for i, y in enumerate(ys):
            vals = col_nums(table, y)
            acc = [[] for _ in order]
            for xv, v in zip(xs, vals):
                if v is not None:
                    acc[pos[xv]].append(v)
            if any(len(a) > 1 for a in acc):
                note = f"같은 '{x}' 값이 여러 행이라 {'평균' if agg == 'mean' else '합계'}로 묶음"
            vs = [None if not a else (sum(a) / len(a) if agg == "mean" else sum(a)) for a in acc]
            series.append({"name": y, "values": vs, "unit": unit_of(table, y), "idx": i, "right": y == spec.get("y2") and spec.get("kind") == "dual",
                           "hidden": y in set(spec.get("hidden") or [])})
    cats = [c if c != "" else "(빈 값)" for c in order]
    # 정렬 (시간축은 정렬하지 않음)
    srt = spec.get("sort") or "none"
    vis = [s for s in series if not s["hidden"]]
    if srt in ("asc", "desc") and xtype != "date" and vis:
        tot = [sum((s["values"][i] or 0) for s in vis) for i in range(len(cats))]
        idx = sorted(range(len(cats)), key=lambda i: tot[i], reverse=srt == "desc")
        cats = [cats[i] for i in idx]
        for s in series:
            s["values"] = [s["values"][i] for i in idx]
    top = int(spec.get("top") or 0)
    if top and len(cats) > top:
        cats = cats[:top]
        for s in series:
            s["values"] = s["values"][:top]
    return {"cats": cats, "series": series, "note": note, "xtype": xtype}


def scatter_data(table, spec):
    x = spec["x"]
    xs = col_nums(table, x)
    out = []
    if spec.get("group"):
        gs = col_values(table, spec["group"])
        y = spec["ys"][0]
        ys = col_nums(table, y)
        levels = []
        for g in gs:
            if g not in levels:
                levels.append(g)
        for i, g in enumerate(levels):
            pts = [(a, b) for a, b, gg in zip(xs, ys, gs) if gg == g and a is not None and b is not None]
            out.append({"name": g, "points": pts, "idx": i, "hidden": g in set(spec.get("hidden") or [])})
    else:
        for i, y in enumerate(spec["ys"]):
            ys = col_nums(table, y)
            pts = [(a, b) for a, b in zip(xs, ys) if a is not None and b is not None]
            out.append({"name": y, "points": pts, "idx": i, "hidden": y in set(spec.get("hidden") or [])})
    return out


def dist_data(table, spec):
    """히스토그램·상자그림: 계열(열 또는 그룹)별 숫자 목록"""
    if spec.get("group"):
        y = spec["ys"][0]
        gs = col_values(table, spec["group"])
        vs = col_nums(table, y)
        levels = []
        for g in gs:
            if g not in levels:
                levels.append(g)
        return [{"name": g, "values": [v for v, gg in zip(vs, gs) if gg == g and v is not None], "idx": i,
                 "hidden": g in set(spec.get("hidden") or [])} for i, g in enumerate(levels)]
    return [{"name": y, "values": [v for v in col_nums(table, y) if v is not None], "idx": i,
             "hidden": y in set(spec.get("hidden") or [])} for i, y in enumerate(spec["ys"])]


def wa(word):
    """받침 따라 '와/과' — 마지막 글자가 한글이 아니면 '와'"""
    ch = (word or " ").rstrip()[-1:]
    return word + ("과" if "가" <= ch <= "힣" and (ord(ch) - 0xAC00) % 28 else "와")


# ── 추천 규칙 ─────────────────────────────────────────────────────────────
def _levels(vals):
    out = []
    for v in vals:
        if v not in out:
            out.append(v)
    return out


def recommend(table, limit=12):
    """데이터 모양을 보고 적합한 그래프 후보 명세 목록. 부적합한 차트는 만들지 않는다."""
    cols = table["columns"]
    rows = table["rows"]
    n = len(rows)
    dates = [c["name"] for c in cols if c["type"] == "date"]
    cats = [c["name"] for c in cols if c["type"] == "cat"]
    nums = [c["name"] for c in cols if c["type"] == "num" and sum(v is not None for v in col_nums(table, c["name"])) >= 2]
    out = []

    def add(kind, why, score, **kw):
        sp = {"kind": kind, "why": why, "score": score, **kw}
        key = (kind, sp.get("x"), tuple(sp.get("ys") or []), sp.get("group"), sp.get("y2"))
        if any((o["kind"], o.get("x"), tuple(o.get("ys") or []), o.get("group"), o.get("y2")) == key for o in out):
            return
        out.append(sp)

    if not nums:
        return []
    by_unit = {}
    for y in nums:
        by_unit.setdefault(unit_of(table, y), []).append(y)
    main_unit, main_ys = max(by_unit.items(), key=lambda kv: len(kv[1]))
    main_ys = main_ys[:8]
    x_time = dates[0] if dates else None
    x_cat = None
    for c in cats:
        lv = _levels(col_values(table, c))
        if 2 <= len(lv) <= 60:
            x_cat = c
            break
    x = x_time or x_cat
    # 긴 형식(x 반복 + 다른 범주 열) → 피벗
    group = None
    if x:
        xs = col_values(table, x)
        nx = len(_levels(xs))
        if nx < n:
            for g in [c for c in cats + dates if c != x]:
                gl = _levels(col_values(table, g))
                pairs = len(set(zip(xs, col_values(table, g))))
                if 2 <= len(gl) <= 8 and pairs >= 0.9 * n:
                    group = g
                    break
    nx = len(_levels(col_values(table, x))) if x else n
    nonneg = lambda ys: all(v is None or v >= 0 for y in ys for v in col_nums(table, y))

    def add_dual(xc, units):
        """단위가 다른 두 묶음: 왼쪽 축 1~2 계열(막대) + 오른쪽 축 1 계열(꺾은선)"""
        if nx > 24:
            return
        lu = main_unit if main_unit in units else units[0]
        ru = next(u for u in units if u != lu)
        left, right = by_unit[lu][:2], by_unit[ru][0]
        add("dual", f"{'·'.join(left)}({lu or '단위 없음'}) 막대 + {right}({ru or '단위 없음'}) 오른쪽 축 꺾은선 — 두 축 눈금이 독립, 추세 비교 시 주의",
            83, x=xc, ys=left, y2=right, dual_shape="bar_line")

    if x and group:
        y0 = main_ys[0]
        k = len(_levels(col_values(table, group)))
        if x == x_time:
            add("line", f"{x} 흐름에 따라 {group}별 {y0} 추이를 비교", 96, x=x, ys=[y0], group=group)
        if nx <= 16:
            add("col_group", f"{x}마다 {group} {k}개를 나란히 비교", 92, x=x, ys=[y0], group=group)
        if nonneg([y0]) and nx <= 24:
            add("col_stack", f"{group}별 구성과 {x}별 합계를 함께", 85, x=x, ys=[y0], group=group)
            add("col_stack100", f"{x}마다 {group} 비중(%)이 어떻게 달라지는지", 72, x=x, ys=[y0], group=group)
        if nx >= 3 and k >= 3:
            add("heatmap", f"{x} × {group} 값을 색 농도로 한눈에", 66, x=x, ys=[y0], group=group)
        add("box", f"{group}별 {y0} 분포(중앙값·사분위) 비교", 50, x=None, ys=[y0], group=group)

    if x_time and main_ys:
        k = len(main_ys)
        add("line", f"{x_time} 순서가 있어 추세를 보기 좋음" + (f" — {k}개 계열 비교" if k > 1 else ""), 95,
            x=x_time, ys=main_ys[:6])
        if nx <= 24:
            add("col", f"{x_time}별 {main_ys[0]} 크기를 막대로", 88 if k == 1 else 70, x=x_time, ys=main_ys[:1])
        if 2 <= k <= 4 and nx <= 12:
            add("col_group", f"{x_time}마다 {k}개 계열을 나란히 비교", 86, x=x_time, ys=main_ys[:4])
        if k >= 2 and nonneg(main_ys):
            add("col_stack", f"계열 합계와 구성을 함께 (같은 단위{('·' + main_unit) if main_unit else ''})", 78, x=x_time, ys=main_ys[:6])
            add("area_stack", "누적 흐름으로 전체 규모 변화와 구성", 64, x=x_time, ys=main_ys[:6])
        if k == 1 and nonneg(main_ys):
            add("area", f"{main_ys[0]} 추이를 면적으로 강조", 60, x=x_time, ys=main_ys[:1])
        if len(by_unit) >= 2:
            units = list(by_unit)[:2]
            a, b = by_unit[units[0]][0], by_unit[units[1]][0]
            add("panels", f"단위가 다른 {a}({units[0] or '단위 없음'})·{b}({units[1] or '단위 없음'}) — 축을 나눠 위아래 두 그래프로", 84,
                x=x_time, ys=[a, b])
            add_dual(x_time, units)
            if all(v and v > 0 for v in col_nums(table, a)[:1] + col_nums(table, b)[:1]):
                add("index", f"단위가 다른 {a}·{b} 증감률을 첫 값=100 으로 맞춰 비교", 74, x=x_time, ys=[a, b])
        elif k >= 2 and nx >= 3:
            firsts = [col_nums(table, y)[0] for y in main_ys]
            if all(f and f > 0 for f in firsts):
                mx = max(max(v for v in col_nums(table, y) if v is not None) for y in main_ys)
                mn = min(max(v for v in col_nums(table, y) if v is not None) for y in main_ys)
                if mn > 0 and mx / mn >= 20 and nx <= 24:  # 같은 단위라도 자릿수가 크게 다르면 큰 쪽 왼축 막대 + 작은 쪽 오른축 선
                    big = max(main_ys, key=lambda y: max(v for v in col_nums(table, y) if v is not None))
                    small = min(main_ys, key=lambda y: max(v for v in col_nums(table, y) if v is not None))
                    add("dual", f"{big}·{small} 크기 자릿수가 달라 {small}을(를) 오른쪽 축 꺾은선으로 — 두 축 눈금이 독립, 추세 비교 시 주의", 66,
                        x=x_time, ys=[big], y2=small, dual_shape="bar_line")
                if mn > 0 and mx / mn >= 5:
                    add("index", "크기 차이가 큰 계열들의 증감률을 첫 값=100 으로 맞춰 비교", 68, x=x_time, ys=main_ys[:4])
        if k >= 3 and nx >= 3:
            add("heatmap", f"{x_time} × 계열 값을 색 농도로", 55, x=x_time, ys=main_ys[:8])

    if x_cat and not group and main_ys and x_cat != x_time:
        k = len(main_ys)
        y0 = main_ys[0]
        long_lbl = max((len(str(v)) for v in col_values(table, x_cat)), default=0) > 6
        if nx <= 30:
            add("bar", f"항목 {nx}개를 크기순으로 — 이름이 길어도 읽기 쉬움", 93 if (long_lbl or nx > 8) else 84,
                x=x_cat, ys=[y0], sort="desc")
        if nx <= 12:
            add("col", f"{x_cat}별 {y0} 비교", 86 if not long_lbl else 70, x=x_cat, ys=[y0])
        if 2 <= k <= 4 and nx <= 12:
            add("col_group" if not long_lbl else "bar_group", f"{x_cat}마다 {k}개 계열을 나란히", 85, x=x_cat, ys=main_ys[:4])
        if k >= 2 and nonneg(main_ys) and nx <= 20:
            add("bar_stack" if long_lbl else "col_stack", "항목별 합계와 구성을 함께", 76, x=x_cat, ys=main_ys[:6], sort="desc")
            add("col_stack100", "항목마다 구성 비율(%) 비교", 70, x=x_cat, ys=main_ys[:6])
        vals = [v for v in col_nums(table, y0) if v is not None]
        if 2 <= nx <= 6 and vals and all(v > 0 for v in vals) and len(vals) == nx:
            add("doughnut", f"{nx}개 항목이 전체에서 차지하는 비중", 80, x=x_cat, ys=[y0])
            add("pie", f"{nx}개 항목의 구성비 — 범주가 적어 원그래프가 적합", 74, x=x_cat, ys=[y0])
        if k >= 3 and nx >= 3:
            add("heatmap", f"{x_cat} × 계열 값을 색 농도로 비교", 62, x=x_cat, ys=main_ys[:8])
        if len(by_unit) >= 2 and nx <= 20:
            units = list(by_unit)[:2]
            add_dual(x_cat, units)
            add("panels", "단위가 다른 두 계열 — 축을 나눠 위아래 두 그래프로", 72, x=x_cat,
                ys=[by_unit[units[0]][0], by_unit[units[1]][0]])

    # 숫자 열 쌍 → 산점도 (상관 강한 순 2개)
    pairs = []
    if len(nums) >= 2 and n >= (12 if x_time else 8):
        for i in range(len(nums)):
            for j in range(len(nums)):
                if i < j:
                    lr = linreg(col_nums(table, nums[i]), col_nums(table, nums[j]))
                    if lr:
                        pairs.append((abs(lr["r"]), nums[i], nums[j], lr))
        pairs.sort(reverse=True)
        sc_group = None
        for c in cats:
            lv = _levels(col_values(table, c))
            if 2 <= len(lv) <= 3 and len(lv) < n:
                sc_group = c
                break
        for r, a, b, lr in pairs[:2]:
            add("scatter", f"{wa(a)} {b}의 관계" + (f" — 상관계수 r={lr['r']:.2f}, 추세선 포함" if r >= .3 else " (상관 약함)"),
                (45 if x_time else 70) + 20 * r, x=a, ys=[b], trend=r >= .3)
            if sc_group:
                add("scatter", f"{sc_group}별로 색을 나눠 {a}–{b} 관계 비교", 70 + 5 * r, x=a, ys=[b], group=sc_group, trend=False)

    # 분포 — 산점도의 결과 변수(y)를 먼저
    resp = [p[2] for p in pairs[:2]] if len(nums) >= 2 and n >= (12 if x_time else 8) else []
    if n >= 15:
        dn = [y for y in nums if len(set(v for v in col_nums(table, y) if v is not None)) >= 6]
        dn.sort(key=lambda y: 0 if y in resp else 1)
        if dn:
            add("hist", f"{dn[0]} 값 {n}개가 어느 구간에 몰려 있는지", 75 if not x_time else 52, ys=[dn[0]])
        if len(dn) >= 2:
            add("hist", f"{dn[1]} 값의 분포", 48, ys=[dn[1]])
    bg = next((c for c in cats if 2 <= len(_levels(col_values(table, c))) <= 8 and len(_levels(col_values(table, c))) * 3 <= n), None)
    if bg and not group and n >= 8:
        yb = resp[0] if resp else main_ys[0]
        add("box", f"{bg}별 {yb} 분포(중앙값·사분위·이상값) 비교", 66, x=None, ys=[yb], group=bg)
    if n >= 8 and 2 <= len(main_ys) <= 8 and not group:
        add("box", f"{'·'.join(main_ys[:4])}{'…' if len(main_ys) > 4 else ''} 분포를 중앙값·사분위로 비교", 58 if not x_time else 40,
            ys=main_ys[:8])
    if not out and nums:  # 범주 열이 없고 행이 적음 → 행 번호 기준 막대
        add("col", "행 순서대로 값 비교", 50, x=None, ys=nums[:1])

    out.sort(key=lambda s: -s["score"])
    out = out[:limit]
    for i, s in enumerate(out):
        s["id"] = f"c{i + 1}"
        s["name"] = KINDS[s["kind"]]
        s.setdefault("title", default_title(table, s))
        s.setdefault("legend", "top")
        s.setdefault("labels", "auto")
    return out


def default_title(table, spec):
    ys = spec.get("ys") or []
    k = spec["kind"]
    if k == "scatter":
        return f"{wa(spec['x'])} {ys[0]}의 관계"
    if k in ("hist",):
        return f"{ys[0]} 분포"
    if k == "box":
        return f"{spec['group']}별 {ys[0]} 분포" if spec.get("group") else f"{'·'.join(ys[:3])} 분포"
    if k == "index":
        return f"{'·'.join(ys[:3])} 증감 비교 (첫 값=100)"
    if k == "dual":
        return f"{spec['x']}별 {'·'.join(ys[:2])}과 {spec.get('y2', '')}"
    if spec.get("group"):
        return f"{spec['x']}·{spec['group']}별 {ys[0]}"
    if k in ("pie", "doughnut"):
        return f"{spec['x']}별 {ys[0]} 구성비"
    if k == "col_stack100":
        return f"{spec.get('x') or '항목'}별 구성비"
    nm = "·".join(ys[:3]) + ("…" if len(ys) > 3 else "")
    return f"{spec['x']}별 {nm}" if spec.get("x") else nm


def summarize(table, max_cols=12):
    """LLM 에 넘길 요약(숫자는 실제 값만)"""
    out = []
    for c in table["columns"][:max_cols]:
        vals = col_values(table, c["name"])
        if c["type"] == "num":
            ns = [v for v in (num(x) for x in vals) if v is not None]
            if ns:
                out.append(f"- {c['name']} (숫자{', 단위 ' + c['unit'] if c['unit'] else ''}): 최소 {fmt(min(ns))}, 최대 {fmt(max(ns))}, 첫 값 {fmt(ns[0])}, 끝 값 {fmt(ns[-1])}")
        else:
            lv = _levels(vals)
            out.append(f"- {c['name']} ({'날짜·기간' if c['type'] == 'date' else '범주'}, {len(lv)}종): {', '.join(lv[:8])}{' …' if len(lv) > 8 else ''}")
    return f"행 {len(table['rows'])}개" + (f", 표 제목 '{table['title']}'" if table.get("title") else "") + "\n" + "\n".join(out)
