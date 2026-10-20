"""chart local — 색 프로필(팔레트)

명세의 palette: "<id>" 로 고른다. 계열별 colors 가 있으면 그것이 먼저(덮어쓰기).
palette: "custom" + palette_colors: ["#rrggbb", …] 는 사용자가 고친 색(화면 '내 색 프로필'·'원본 그림 색').

kind
  categorical — 계열마다 다른 색(순서 고정, 순환 금지 규칙은 svg 쪽과 같음)
  sequential  — 한 색의 명도 단계(진함→옅음). 누적·순서 있는 계열에 알맞음
  print       — 흑백 인쇄: 회색 단계 + 막대·조각은 무늬, 선은 점선 모양으로 구분 (SVG·오피스 네이티브 둘 다)
히트맵·트리맵은 팔레트의 주 색(colors[0], seq 가 있으면 그것) 명도 단계를 쓴다.
"""
import re

HEX = re.compile(r"#[0-9a-fA-F]{6}")

# 기본 — svg.PAL 과 같은 값(바꾸지 않음: 지난 결과와 같은 그림)
DEFAULT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
DEFAULT_SEQ = ["#cde2fb", "#b7d3f6", "#9ec5f4", "#86b6ef", "#6da7ec", "#5598e7", "#3987e5", "#2a78d6", "#256abf",
               "#1c5cab", "#184f95", "#104281", "#0d366b"]

# 흑백 인쇄: 면(막대·조각) 회색 + 무늬, 선은 진한 회색 + 선 모양. 순서마다 (면 회색, 선 회색, 무늬, 선 모양)
PRINT = [
    ("#3d3d3d", "#000000", None, None),
    ("#bdbdbd", "#3a3a3a", "diag", "6 3"),
    ("#ededed", "#1f1f1f", "dots", "2 2.5"),
    ("#8a8a8a", "#5a5a5a", "cross", "9 3 2 3"),
    ("#d4d4d4", "#2b2b2b", "horz", "12 4"),
    ("#5e5e5e", "#6e6e6e", "vert", "3 3 3 6"),
    ("#a3a3a3", "#141414", "back", "1.5 3"),
    ("#f7f7f7", "#4a4a4a", "grid", "10 3 2 3 2 3"),
]
# 오피스(DrawingML) 대응: 무늬 → a:pattFill prst, 선 모양 → a:prstDash
OFFICE_PATT = {"diag": "wdUpDiag", "dots": "pct20", "cross": "diagCross", "horz": "narHorz", "vert": "narVert",
               "back": "wdDnDiag", "grid": "smGrid"}
OFFICE_DASH = {"6 3": "dash", "2 2.5": "sysDot", "9 3 2 3": "dashDot", "12 4": "lgDash", "3 3 3 6": "sysDash",
               "1.5 3": "sysDot", "10 3 2 3 2 3": "lgDashDotDot"}

PALETTES = [
    {"id": "default", "name": "기본", "kind": "categorical", "colors": DEFAULT, "seq": DEFAULT_SEQ,
     "desc": "지금까지 쓰던 색 — 화면·보고서 어디에나 무난"},
    {"id": "pastel", "name": "파스텔", "kind": "categorical",
     "colors": ["#7fb0e3", "#f4a582", "#86cfae", "#f2c86b", "#c2a5de", "#f09cb5", "#8ccfdc", "#c4bd78"],
     "seq_base": "#5b93cf", "desc": "부드러운 색 — 발표 화면·안내문. 값 글자는 진한 색"},
    {"id": "vivid", "name": "선명(비비드)", "kind": "categorical",
     "colors": ["#0057e0", "#ff5a00", "#00a651", "#ffb400", "#e6007e", "#7a2fbf", "#00a3e0", "#c8102e"],
     "desc": "채도 높은 색 — 멀리서 보는 발표 화면"},
    {"id": "contrast", "name": "고대비", "kind": "categorical",
     "colors": ["#003d99", "#d14900", "#1a1a1a", "#00703c", "#a3001b", "#6b2fa0", "#b07d00", "#007a99"],
     "desc": "어두운 진한 색 — 흰 바탕에서 또렷, 흐린 프로젝터·저시력"},
    {"id": "okabe", "name": "색각이상 배려 (Okabe–Ito)", "kind": "categorical",
     "colors": ["#0072b2", "#e69f00", "#009e73", "#d55e00", "#cc79a7", "#56b4e9", "#000000", "#f0e442"],
     "desc": "적록 색각이상에서도 구분되는 Okabe–Ito 8색"},
    {"id": "print", "name": "흑백 인쇄", "kind": "print", "colors": [p[0] for p in PRINT],
     "seq": ["#f5f5f5", "#e6e6e6", "#d6d6d6", "#c4c4c4", "#b3b3b3", "#a1a1a1", "#8f8f8f", "#7d7d7d", "#6b6b6b",
             "#595959", "#474747", "#333333", "#1f1f1f"],
     "line_colors": [p[1] for p in PRINT], "patterns": [p[2] for p in PRINT], "dashes": [p[3] for p in PRINT],
     "desc": "회색 단계 + 막대 무늬·선 모양 — 흑백 프린터에서도 계열 구분"},
    {"id": "warm", "name": "따뜻한 색", "kind": "categorical",
     "colors": ["#c0392b", "#ef7d00", "#e8b100", "#8c2d04", "#e3667a", "#b8651b", "#f4a582", "#7a1f3d"],
     "desc": "빨강·주황·노랑 계열"},
    {"id": "cool", "name": "차가운 색", "kind": "categorical",
     "colors": ["#08519c", "#1497a8", "#6a51a3", "#2ca25f", "#6baed6", "#253494", "#74c476", "#9e9ac8"],
     "desc": "파랑·청록·보라·초록 계열"},
    {"id": "mono", "name": "단색 계열", "kind": "sequential",
     "colors": ["#0b2f5e", "#1f5fa8", "#4a8fd1", "#86b9e6", "#bcd8f2", "#163f78", "#3477bf", "#6aa6dc"],
     "seq_base": "#1f5fa8", "desc": "한 색의 명도 단계(앞 5개가 진함→옅음) — 누적·순서 있는 계열(1분기→4분기 등)"},
    {"id": "report", "name": "보고서용", "kind": "categorical",
     "colors": ["#1f3a5f", "#5b84b1", "#a9bcd6", "#b8863b", "#4f7472", "#7d6b91", "#8c8c8c", "#2e5e4e"],
     "seq_base": "#1f3a5f", "desc": "차분한 남색 중심 — 공문·연구 보고서"},
]
BY_ID = {p["id"]: p for p in PALETTES}
IDS = [p["id"] for p in PALETTES]
SPECIAL = ("custom",)   # palette_colors 를 함께 줌
MAX_CUSTOM = 12


def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _hex(r, g, b):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(v)))) for v in (r, g, b))


def mix(a, b, t):
    """a → b 로 t(0~1) 만큼"""
    ra, rb = _rgb(a), _rgb(b)
    return _hex(*(x + (y - x) * t for x, y in zip(ra, rb)))


def ramp(base, n=13):
    """주 색 하나 → 옅음→진함 n 단계 (주 색이 2/3 쯤, 기본 SEQ 와 같은 짜임)"""
    k = round((n - 1) * 7 / 12)
    out = []
    for i in range(n):
        if i <= k:
            out.append(mix(base, "#ffffff", 0.84 * (k - i) / k))
        else:
            out.append(mix(base, "#000000", 0.58 * (i - k) / (n - 1 - k)))
    return out


def clean_colors(v):
    """palette_colors → 검증된 '#rrggbb' 목록(최대 12). 형식이 틀린 항목은 버림"""
    if not isinstance(v, list):
        return []
    return [c.lower() for c in v[:MAX_CUSTOM] if isinstance(c, str) and HEX.fullmatch(c)]


def get(spec):
    """명세 → 팔레트 dict (모르는 id·빈 custom 이면 기본)"""
    pid = (spec or {}).get("palette")
    if pid == "custom":
        cs = clean_colors((spec or {}).get("palette_colors"))
        if cs:
            return {"id": "custom", "name": "내 색 프로필", "kind": "categorical", "colors": cs}
        return BY_ID["default"]
    return BY_ID.get(pid) or BY_ID["default"]


def colors(spec):
    return get(spec)["colors"]


def seq(spec):
    p = get(spec)
    if p.get("seq"):
        return p["seq"]
    return ramp(p.get("seq_base") or p["colors"][0])


def is_default(spec):
    return get(spec)["id"] == "default"


def is_print(spec):
    return get(spec)["kind"] == "print"


def key(spec):
    """SVG 안 id(그라데이션·무늬)용 짧은 꼬리 — 한 화면에 여러 팔레트 그림이 있어도 id 가 안 겹치게"""
    p = get(spec)
    if p["id"] != "custom":
        return p["id"]
    import hashlib
    return "c" + hashlib.sha1(",".join(p["colors"]).encode()).hexdigest()[:8]


def normalize_spec(spec):
    """서버가 받는 명세의 palette 칸 정리: 모르는 id 는 지움, custom 은 색 검증. 제자리 수정 + 반환"""
    if not isinstance(spec, dict):
        return spec
    pid = spec.get("palette")
    if pid is None:
        spec.pop("palette_colors", None)
        return spec
    if pid == "custom":
        cs = clean_colors(spec.get("palette_colors"))
        if cs:
            spec["palette_colors"] = cs
        else:
            spec.pop("palette", None)
            spec.pop("palette_colors", None)
    elif pid not in BY_ID:
        spec.pop("palette", None)
        spec.pop("palette_colors", None)
    else:
        spec.pop("palette_colors", None)
    return spec


def source_colors_for(spec, source, fallback=None):
    """'원본 그림 색': 열 이름 → 색 지도(source)를 이 명세의 값 열(ys) 순서 색 목록으로.
    열 이름이 같거나 '이름 y'·'이름 (상대값)' 처럼 원본 계열 이름으로 시작하면 짝. 못 찾으면 기본 팔레트 색"""
    fb = fallback or DEFAULT
    ys = list((spec or {}).get("ys") or [])
    out = []
    for i, y in enumerate(ys):
        c = match_source(y, source)
        out.append(c or fb[i % len(fb)])
    return out


def match_source(col, source):
    if not isinstance(source, dict):
        return None
    if col in source:
        return source[col]
    best = None
    for name, c in source.items():
        if name and (col.startswith(name + " ") or col.startswith(name + "(")) and (best is None or len(name) > len(best[0])):
            best = (name, c)
    return best[1] if best else None


def apply_choice(spec, choice, palette_colors=None, source=None):
    """화면 갤러리 수준 선택을 명세 하나에 적용(제자리 수정). choice: 팔레트 id | 'custom' | 'source'.
    'source' 는 값 열 이름으로 원본 그림 색을 찾아 custom 으로 바꾼다(그룹 계열처럼 열 이름이 아닌 계열은 기본 색)"""
    if not choice or choice == "default":
        spec.pop("palette", None)
        spec.pop("palette_colors", None)
        return spec
    if choice == "source":
        if spec.get("group") or not isinstance(source, dict) or not source:
            return spec
        spec["palette"] = "custom"
        spec["palette_colors"] = [c.lower() for c in source_colors_for(spec, source)]
        return normalize_spec(spec)
    spec["palette"] = choice
    if choice == "custom":
        spec["palette_colors"] = palette_colors
    return normalize_spec(spec)


def meta_list():
    out = []
    for p in PALETTES:
        d = {"id": p["id"], "name": p["name"], "kind": p["kind"], "colors": list(p["colors"]), "seq": seq({"palette": p["id"]}),
             "desc": p["desc"]}
        if p["kind"] == "print":
            d["line_colors"] = list(p["line_colors"])
            d["patterns"] = list(p["patterns"])
            d["dashes"] = list(p["dashes"])
        out.append(d)
    return out


# ── 대비 ──────────────────────────────────────────────────────────────────
def lum(hexc):
    r, g, b = (v / 255 for v in _rgb(hexc))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4  # noqa: E731
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a, b):
    la, lb = sorted((lum(a), lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


# ── 흑백 인쇄: SVG 무늬·선 모양 ───────────────────────────────────────────
def _pattern_svg(pid, kind, bg):
    fg = "#ffffff" if lum(bg) < 0.18 else "#1a1a1a"
    op = "0.85" if fg == "#ffffff" else "0.7"
    body = {
        "diag": f'<path d="M-1,1 l2,-2 M0,8 l8,-8 M7,9 l2,-2" stroke="{fg}" stroke-width="1.3" stroke-opacity="{op}"/>',
        "back": f'<path d="M-1,7 l2,2 M0,0 l8,8 M7,-1 l2,2" stroke="{fg}" stroke-width="1.3" stroke-opacity="{op}"/>',
        "dots": f'<circle cx="2" cy="2" r="1.1" fill="{fg}" fill-opacity="{op}"/><circle cx="6" cy="6" r="1.1" fill="{fg}" fill-opacity="{op}"/>',
        "cross": f'<path d="M0,8 l8,-8 M0,0 l8,8" stroke="{fg}" stroke-width="0.9" stroke-opacity="{op}"/>',
        "horz": f'<path d="M0,2 h8 M0,6 h8" stroke="{fg}" stroke-width="1.1" stroke-opacity="{op}"/>',
        "vert": f'<path d="M2,0 v8 M6,0 v8" stroke="{fg}" stroke-width="1.1" stroke-opacity="{op}"/>',
        "grid": f'<path d="M0,4 h8 M4,0 v8" stroke="{fg}" stroke-width="0.8" stroke-opacity="{op}"/>',
    }[kind]
    return (f'<pattern id="{pid}" patternUnits="userSpaceOnUse" width="8" height="8">'
            f'<rect width="8" height="8" fill="{bg}"/>{body}</pattern>')


_ATTR = re.compile(r'(<(rect|path|circle|polygon|line|polyline|ellipse)\b)([^>]*?)(/?>)')


def print_svg(svg_text, spec):
    """흑백 인쇄 팔레트 그림 후처리: 면 회색 → 무늬, 선 회색 → 진한 회색 + 선 모양(stroke-dasharray).
    무늬·선 모양은 계열 순서(팔레트 색)로 정해지므로 범례 견본도 같은 모양이 된다. 글자(<text>)는 건드리지 않는다"""
    p = get(spec)
    if p["kind"] != "print":
        return svg_text
    fills = {c: i for i, c in enumerate(p["colors"])}
    used = set()

    def one(m):
        head, tag, attrs, end = m.group(1), m.group(2), m.group(3), m.group(4)
        fm = re.search(r'\bfill="(#[0-9a-f]{6})"', attrs)
        sm = re.search(r'\bstroke="(#[0-9a-f]{6})"', attrs)
        if fm and fm.group(1) in fills:
            i = fills[fm.group(1)]
            if tag == "circle":   # 점(표식)은 무늬 대신 진한 선 회색
                attrs = attrs.replace(fm.group(0), f'fill="{p["line_colors"][i]}"')
            elif p["patterns"][i]:
                used.add(i)
                attrs = attrs.replace(fm.group(0), f'fill="url(#pp-{i})" data-fill="{fm.group(1)}"')
        if sm and sm.group(1) in fills:
            i = fills[sm.group(1)]
            translucent = re.search(r'\bfill-opacity="0?\.\d+"', attrs)
            filled = re.search(r'\bfill="(?!none)', attrs) and tag != "line" and not translucent
            if not filled:   # 선(경로·선분) — 진한 회색 + 선 모양
                new = f'stroke="{p["line_colors"][i]}"'
                if p["dashes"][i] and "stroke-dasharray" not in attrs:
                    new += f' stroke-dasharray="{p["dashes"][i]}"'
                attrs = attrs.replace(sm.group(0), new)
        return head + attrs + end

    out = _ATTR.sub(one, svg_text)
    if not used:
        return out
    defs = "<defs>" + "".join(_pattern_svg(f"pp-{i}", p["patterns"][i], p["colors"][i]) for i in sorted(used)) + "</defs>"
    # 첫 바탕 rect 뒤에 defs
    k = out.find("/>", out.find("<rect")) + 2
    return out[:k] + defs + out[k:]


# ── 흑백 인쇄: 오피스 차트 XML ─────────────────────────────────────────────
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
LINE_CHARTS = {"lineChart", "line3DChart", "scatterChart", "radarChart", "stockChart"}


def print_chart_xml(xml_bytes, spec_or_pal=None):
    """차트 파트(chartSpace) XML 에서 흑백 인쇄 회색을 찾아: 계열·점의 면 채우기 → a:pattFill(무늬), 선 계열의 선 → 진한 회색 + a:prstDash.
    회색 값이 팔레트에만 쓰는 값이라 색으로 계열 순서를 안다. 바뀐 것이 없으면 원래 바이트 그대로"""
    from lxml import etree
    p = BY_ID["print"]
    fills = {c.lstrip("#").upper(): i for i, c in enumerate(p["colors"])}
    root = etree.fromstring(xml_bytes)
    q = lambda ns, t: f"{{{ns}}}{t}"  # noqa: E731
    changed = False
    for sppr in root.iter(q(C_NS, "spPr")):
        par = sppr.getparent()
        if par is None or etree.QName(par).localname not in ("ser", "dPt", "upBars", "downBars"):
            continue
        chart_kind = None
        a = par
        while a is not None:
            ln_ = etree.QName(a).localname
            if ln_.endswith("Chart") and ln_ != "chart":
                chart_kind = ln_
                break
            a = a.getparent()
        # 면
        sf = sppr.find(q(A_NS, "solidFill"))
        if sf is not None:
            clr = sf.find(q(A_NS, "srgbClr"))
            v = clr.get("val", "").upper() if clr is not None else ""
            if v in fills and p["patterns"][fills[v]] and chart_kind not in LINE_CHARTS:
                i = fills[v]
                fg = "FFFFFF" if lum("#" + v) < 0.18 else "1A1A1A"
                pf = etree.Element(q(A_NS, "pattFill"), prst=OFFICE_PATT[p["patterns"][i]])
                fgc = etree.SubElement(pf, q(A_NS, "fgClr"))
                etree.SubElement(fgc, q(A_NS, "srgbClr"), val=fg)
                bgc = etree.SubElement(pf, q(A_NS, "bgClr"))
                etree.SubElement(bgc, q(A_NS, "srgbClr"), val=v)
                sppr.replace(sf, pf)
                changed = True
        # 선
        ln = sppr.find(q(A_NS, "ln"))
        if ln is not None and chart_kind in LINE_CHARTS:
            sf = ln.find(q(A_NS, "solidFill"))
            clr = sf.find(q(A_NS, "srgbClr")) if sf is not None else None
            v = clr.get("val", "").upper() if clr is not None else ""
            if v in fills:
                i = fills[v]
                clr.set("val", p["line_colors"][i].lstrip("#").upper())
                if p["dashes"][i] and ln.find(q(A_NS, "prstDash")) is None:
                    d = etree.Element(q(A_NS, "prstDash"), val=OFFICE_DASH[p["dashes"][i]])
                    sf.addnext(d)
                changed = True
    if not changed:
        return xml_bytes
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


CHART_PART = re.compile(r"^(xl|ppt|word)/charts/chart\d+\.xml$")


def print_office_file(path):
    """저장된 XLSX·PPTX·DOCX 안 차트 파트를 흑백 인쇄 모양으로 고쳐 다시 쓴다. 바꾼 차트 파트 수 반환"""
    import shutil
    import tempfile
    import zipfile
    n = 0
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    out = []
    for info, data in items:
        if CHART_PART.match(info.filename):
            nd = print_chart_xml(data)
            if nd is not data:
                n += 1
                data = nd
        out.append((info, data))
    if not n:
        return 0
    fd, tmp = tempfile.mkstemp(suffix=".zip", dir=__import__("os").path.dirname(path) or ".")
    __import__("os").close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    shutil.move(tmp, path)
    return n
