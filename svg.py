"""chart local — 그래프 명세 → SVG (자체 렌더, 외부 라이브러리·CDN 없음)

시각 규칙(dataviz 가이드): 범주색은 고정 순서(순환 금지, 최대 8), 막대 두께 ≤24px·데이터 끝 4px 둥글림·기준선 쪽은 직각,
선 2px, 점 r≥4 + 2px 바탕색 테, 누적 조각·인접 막대 사이 2px 바탕색 틈, 격자는 1px 실선 연회색,
글자는 잉크 색(계열색으로 칠하지 않음), 계열 2개 이상이면 범례, 값 라벨은 선택적으로.
이중축(dual)은 보고서 관례 때문에 사용자 결정으로 넣었다: 축 제목을 축마다 계열색으로, 범례에 '(오른쪽 축)', 두 축 눈금 개수·0선을 맞춘다.
"""
import math
from xml.sax.saxutils import escape

import core

PAL = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
SEQ = ["#cde2fb", "#b7d3f6", "#9ec5f4", "#86b6ef", "#6da7ec", "#5598e7", "#3987e5", "#2a78d6", "#256abf",
       "#1c5cab", "#184f95", "#104281", "#0d366b"]
SURFACE = "#ffffff"
INK, INK2, MUTED, GRID, BASE = "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7"
FONT = "'Pretendard','Malgun Gothic','맑은 고딕','Apple SD Gothic Neo','Noto Sans KR','Noto Sans CJK KR',sans-serif"


def tw(s, size):
    """글자 폭 어림 (한글·한자 1em, 숫자·영문 ~0.58em)"""
    w = 0.0
    for ch in str(s):
        o = ord(ch)
        if o >= 0x1100:
            w += 1.0
        elif ch in "il.,:;|!'`":
            w += 0.3
        elif ch in " ()[]-":
            w += 0.38
        elif ch.isupper() or ch in "mwMW%":
            w += 0.72
        else:
            w += 0.56
    return w * size


def clip(s, size, width):
    s = str(s)
    if tw(s, size) <= width:
        return s
    while s and tw(s + "…", size) > width:
        s = s[:-1]
    return s + "…"


def esc(s):
    return escape(str(s), {'"': "&quot;"})


def color_for(spec, idx):
    cs = spec.get("colors") or {}
    if isinstance(cs, list):
        c = cs[idx] if idx < len(cs) and cs[idx] else None
    else:
        c = cs.get(str(idx))
    return c or PAL[idx % len(PAL)]


def darken(hexc, f=0.78):
    """계열색을 글자로 쓸 때(이중축 축 제목) 밝은 색(노랑·청록)도 읽히게 어둡게"""
    h = hexc.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    if lum(hexc) < 0.2:
        return hexc
    return "#%02x%02x%02x" % (int(r * f), int(g * f), int(b * f))


def lum(hexc):
    h = hexc.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def text_on(fill):
    return "#ffffff" if lum(fill) < 0.35 else INK


class Svg:
    def __init__(self, w, h, title):
        self.w, self.h, self.parts = w, h, []
        self.title = title

    def add(self, s):
        self.parts.append(s)

    def text(self, x, y, s, size=12, fill=INK2, anchor="start", weight=400, extra=""):
        self.add(f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size}" fill="{fill}" text-anchor="{anchor}"'
                 f'{f' font-weight="{weight}"' if weight != 400 else ""} {extra}>{esc(s)}</text>')

    def line(self, x1, y1, x2, y2, stroke=GRID, width=1, extra=""):
        self.add(f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="{stroke}" stroke-width="{width}" {extra}/>')

    def out(self):
        return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {self.w} {self.h}" width="{self.w}" height="{self.h}" '
                f'role="img" aria-label="{esc(self.title)}" font-family="{FONT}" style="font-variant-numeric:tabular-nums">'
                f'<rect width="{self.w}" height="{self.h}" fill="{SURFACE}"/>' + "".join(self.parts) + "</svg>")


def tip(*parts):
    t = " · ".join(str(p) for p in parts if p not in (None, ""))
    return f'data-tip="{esc(t)}"', ""  # 툴팁은 화면(ui.html)이 data-tip 으로 띄운다 — <title> 은 브라우저 툴팁과 겹쳐서 뺌


def bar_path(x, y, w, h, r, end):
    """막대 — 데이터 끝(end: top|bottom|right|left)만 r 둥글림, 기준선 쪽 직각"""
    if w <= 0 or h <= 0:
        return ""
    r = max(0, min(r, w / 2, h / 2))
    if end == "top":
        return (f"M{x:.1f},{y + h:.1f}V{y + r:.1f}Q{x:.1f},{y:.1f} {x + r:.1f},{y:.1f}H{x + w - r:.1f}"
                f"Q{x + w:.1f},{y:.1f} {x + w:.1f},{y + r:.1f}V{y + h:.1f}Z")
    if end == "bottom":
        return (f"M{x:.1f},{y:.1f}V{y + h - r:.1f}Q{x:.1f},{y + h:.1f} {x + r:.1f},{y + h:.1f}H{x + w - r:.1f}"
                f"Q{x + w:.1f},{y + h:.1f} {x + w:.1f},{y + h - r:.1f}V{y:.1f}Z")
    if end == "right":
        return (f"M{x:.1f},{y:.1f}H{x + w - r:.1f}Q{x + w:.1f},{y:.1f} {x + w:.1f},{y + r:.1f}V{y + h - r:.1f}"
                f"Q{x + w:.1f},{y + h:.1f} {x + w - r:.1f},{y + h:.1f}H{x:.1f}Z")
    if end == "left":
        return (f"M{x + w:.1f},{y:.1f}H{x + r:.1f}Q{x:.1f},{y:.1f} {x:.1f},{y + r:.1f}V{y + h - r:.1f}"
                f"Q{x:.1f},{y + h:.1f} {x + r:.1f},{y + h:.1f}H{x + w:.1f}Z")
    return f"M{x:.1f},{y:.1f}h{w:.1f}v{h:.1f}h{-w:.1f}Z"


def nice_domain(lo, hi, zero=True, target=5):
    if zero:
        lo, hi = min(0, lo), max(0, hi)
    if lo == hi:
        lo, hi = (lo - 1, hi + 1) if lo else (0, 1)
    step = core.nice_step(hi - lo, target)
    a = math.floor(lo / step + 1e-9) * step
    b = math.ceil(hi / step - 1e-9) * step
    ticks = []
    t = a
    while t <= b + step * 1e-6:
        ticks.append(round(t, 10))
        t += step
    return a, b, ticks, step


def tick_fmt(ticks, step, unit=""):
    dec = 0 if step >= 1 else min(4, max(0, -int(math.floor(math.log10(step)))))
    if step >= 1 and any(abs(t - round(t)) > 1e-9 for t in ticks):
        dec = 1
    return [core.fmt(t, dec, "%" if unit == "%" else "") for t in ticks]


def axis_title(spec, key, default):
    v = spec.get(key)
    return default if v is None else v


# ── 레이아웃: 제목·범례 ─────────────────────────────────────────────────────
def frame(spec, w, h, legend_items):
    """제목·부제·범례를 그리고 플롯 영역 (x0, y0, x1, y1) 반환"""
    g = Svg(w, h, spec.get("title") or "")
    pad = 18
    y = pad
    title = spec.get("title") or ""
    if title:
        g.text(pad, y + 15, clip(title, 17, w - 2 * pad), 17, INK, weight=600)
        y += 24
    if spec.get("subtitle"):
        g.text(pad, y + 12, clip(spec["subtitle"], 12.5, w - 2 * pad), 12.5, INK2)
        y += 19
    y += 6
    x0, x1, y1 = pad, w - pad, h - pad
    pos = spec.get("legend") or "top"
    if len(legend_items) >= 2 and pos != "none":
        if pos == "right":
            lw = min(180, max(tw(n, 12) for n, _, _ in legend_items) + 30)
            ly = y + 8
            for n, c, shape in legend_items:
                _swatch(g, x1 - lw + 2, ly, c, shape)
                g.text(x1 - lw + 22, ly + 4.5, clip(n, 12, lw - 26), 12, INK2)
                ly += 20
            x1 -= lw + 8
        else:
            rows, cur, cw = [], [], 0
            for it in legend_items:
                iw = tw(it[0], 12) + 34
                if cur and cw + iw > x1 - x0:
                    rows.append(cur)
                    cur, cw = [], 0
                cur.append((it, iw))
                cw += iw
            rows.append(cur)
            if pos == "bottom":
                ly = y1 - 20 * len(rows) + 10
                y1 -= 20 * len(rows) + 4
            else:
                ly = y + 6
                y += 20 * len(rows) + 4
            for r in rows:
                lx = x0
                for (n, c, shape), iw in r:
                    _swatch(g, lx, ly, c, shape)
                    g.text(lx + 20, ly + 4.5, n, 12, INK2)
                    lx += iw
                ly += 20
    return g, (x0, y + 4, x1, y1)


def _swatch(g, x, y, c, shape):
    if shape == "line":
        g.add(f'<line x1="{x:.1f}" y1="{y:.1f}" x2="{x + 14:.1f}" y2="{y:.1f}" stroke="{c}" stroke-width="2.5" stroke-linecap="round"/>'
              f'<circle cx="{x + 7:.1f}" cy="{y:.1f}" r="3.5" fill="{c}" stroke="{SURFACE}" stroke-width="1.5"/>')
    elif shape == "dot":
        g.add(f'<circle cx="{x + 6:.1f}" cy="{y:.1f}" r="5" fill="{c}"/>')
    else:
        g.add(f'<rect x="{x:.1f}" y="{y - 6:.1f}" width="13" height="12" rx="3" fill="{c}"/>')


# ── y 축(값) — 세로형 ──────────────────────────────────────────────────────
def y_axis(g, box, lo, hi, ticks, step, unit, title, left_extra=0):
    """눈금 라벨 폭만큼 플롯 왼쪽을 줄이고 격자를 그린다. 반환: (새 x0, sy)"""
    x0, y0, x1, y1 = box
    labels = tick_fmt(ticks, step, unit)
    lw = max(tw(s, 11.5) for s in labels) + 8
    if title:
        g.text(x0, y0 + 2, title, 12, INK2)
        y0 += 14
    px0 = x0 + lw + left_extra
    sy = lambda v: y1 - (v - lo) / (hi - lo) * (y1 - y0)
    for t, s in zip(ticks, labels):
        yy = sy(t)
        if abs(t) > 1e-12:
            g.line(px0, yy, x1, yy, GRID, 1)
        g.text(px0 - 7, yy + 4, s, 11.5, MUTED, "end")
    if lo <= 0 <= hi:
        g.line(px0, sy(0), x1, sy(0), BASE, 1)
    return px0, y0, sy


def unit_title(name, unit):
    if name and unit:
        return f"{name} ({unit})"
    return name or (f"단위: {unit}" if unit else "")


def x_cat_labels(g, cats, cx, band, ybase, x0, x1, xtitle):
    """세로형 x 범주 라벨: 겹치면 기울이고, 그래도 많으면 건너뛴다. 반환: 사용한 높이"""
    if not cats:
        return 0
    maxw = max(tw(c, 11.5) for c in cats)
    used = 18
    if maxw <= band - 4:
        for c in cats:
            g.text(cx(c), ybase + 16, c, 11.5, MUTED, "middle")
    else:
        every = max(1, math.ceil(15 / max(band, 1)))
        lab_w = min(maxw, 110)
        for i, c in enumerate(cats):
            if i % every:
                continue
            x = cx(c)
            g.text(x, ybase + 12, clip(c, 11.5, 110), 11.5, MUTED, "end", extra=f'transform="rotate(-40 {x:.1f} {ybase + 12:.1f})"')
        used = 14 + lab_w * math.sin(math.radians(40)) + 4
    if xtitle:
        g.text((x0 + x1) / 2, ybase + used + 18, xtitle, 12, INK2, "middle")
        used += 22
    return used


def x_label_room(cats, band, xtitle):
    maxw = max((tw(c, 11.5) for c in cats), default=0)
    used = 20 if maxw <= band - 4 else 14 + min(maxw, 110) * math.sin(math.radians(40)) + 6
    return used + (22 if xtitle else 0)


# ── 범주형: 세로·가로 막대, 꺾은선, 영역 ─────────────────────────────────────
def render_category(table, spec, w, h, data=None):
    k = spec["kind"]
    d = data or core.category_data(table, spec)
    cats, series = d["cats"], [s for s in d["series"] if not s["hidden"]]
    if k == "col_stack100":
        tots = [sum(abs(s["values"][i] or 0) for s in series) for i in range(len(cats))]
        for s in series:
            s["values"] = [None if v is None else (100 * v / tots[i] if tots[i] else 0) for i, v in enumerate(s["values"])]
            s["unit"] = "%"
    if k == "index":
        for s in series:
            base = next((v for v in s["values"] if v), None)
            s["values"] = [None if v is None or not base else 100 * v / base for v in s["values"]]
    horiz = k in ("bar", "bar_group", "bar_stack")
    stacked = k in ("col_stack", "col_stack100", "bar_stack", "area_stack")
    is_line = k in ("line", "index")
    is_area = k in ("area", "area_stack")
    shape = "line" if is_line else "bar"
    legend = [(s["name"], color_for(spec, s["idx"]), shape) for s in series]
    g, box = frame(spec, w, h, legend)
    if not cats or not series:
        g.text(w / 2, h / 2, "표시할 값이 없습니다", 13, MUTED, "middle")
        return g.out()
    unit = series[0].get("unit", "")
    if k == "index":
        unit = ""
    vals = [v for s in series for v in s["values"] if v is not None]
    if stacked:
        pos = [sum(max(0, s["values"][i] or 0) for s in series) for i in range(len(cats))]
        neg = [sum(min(0, s["values"][i] or 0) for s in series) for i in range(len(cats))]
        lo, hi = min(neg + [0]), max(pos + [0])
    else:
        lo, hi = (min(vals), max(vals)) if vals else (0, 1)
    zero = not is_line or (lo >= 0 and lo < 0.35 * hi)
    if k == "index":
        zero = False
        lo, hi = min(lo, 100), max(hi, 100)
    if k == "col_stack100":
        lo, hi = 0, 100
    lo, hi, ticks, step = nice_domain(lo, hi, zero)
    ytitle_default = "지수 (첫 값=100)" if k == "index" else ("비율 (%)" if k == "col_stack100" else
                                                         unit_title("" if len(series) > 1 or spec.get("group") else series[0]["name"], unit))
    ytitle = axis_title(spec, "y_title", ytitle_default)
    if spec.get("y_title") and unit and k not in ("index", "col_stack100") and f"({unit})" not in ytitle:
        ytitle = f"{ytitle} ({unit})"
    xtitle = axis_title(spec, "x_title", spec.get("x") or "")
    labels = spec.get("labels") or "auto"
    dec = core.decimals_of(vals)
    vfmt = lambda v: core.fmt(v, dec if k not in ("col_stack100", "index") else 1, "%" if (unit == "%" or k == "col_stack100") else "")
    x0, y0, x1, y1 = box

    if horiz:
        return _render_horizontal(g, spec, k, cats, series, lo, hi, ticks, step, unit, ytitle, xtitle, labels, vfmt, box)

    n = len(cats)
    room = x_label_room(cats, (x1 - x0 - 60) / n, xtitle)
    px0, py0, sy = y_axis(g, (x0, y0, x1, y1 - room), lo, hi, ticks, step, unit if k not in ("col_stack100",) else "%", ytitle)
    py1 = y1 - room
    band = (x1 - px0) / n
    if is_line or is_area:
        cx = lambda c: px0 + band * (cats.index(c) + 0.5)
    else:
        cx = lambda c: px0 + band * (cats.index(c) + 0.5)
    x_cat_labels(g, cats, cx, band, py1, px0, x1, xtitle)
    if k == "index":
        g.line(px0, sy(100), x1, sy(100), BASE, 1.2)

    ns = len(series)
    if k in ("col", "col_group"):
        gw = min(band * 0.78, ns * 24 + (ns - 1) * 2) if ns > 1 else min(24, band * 0.6)
        bw = (gw - (ns - 1) * 2) / ns if ns > 1 else gw
        lab_all = labels == "all" or (labels == "auto" and ns == 1 and n <= 16) or (labels == "auto" and ns <= 3 and n * ns <= 12)
        for si, s in enumerate(series):
            c = color_for(spec, s["idx"])
            for i, v in enumerate(s["values"]):
                if v is None:
                    continue
                bx = px0 + band * i + (band - gw) / 2 + si * (bw + 2)
                ya, yb = sy(max(v, 0)), sy(min(v, 0))
                a, t = tip(cats[i], s["name"] if ns > 1 else "", vfmt(v) + ("" if unit == "%" else (" " + unit if unit else "")))
                g.add(f'<path d="{bar_path(bx, ya, bw, max(yb - ya, 0.5), 4, "top" if v >= 0 else "bottom")}" fill="{c}" {a}>{t}</path>')
                if labels != "none" and lab_all and (bw >= tw(vfmt(v), 10.5) - 6 or ns == 1):
                    ly = ya - 5 if v >= 0 else yb + 13
                    g.text(bx + bw / 2, ly, vfmt(v), 10.5 if ns > 1 else 11.5, INK, "middle")
    elif k in ("col_stack", "col_stack100"):
        bw = min(28, band * 0.62)
        for i in range(n):
            up, dn = 0.0, 0.0
            segs = []
            for s in series:
                v = s["values"][i]
                if v is None or v == 0:
                    continue
                if v > 0:
                    segs.append((s, v, up, up + v))
                    up += v
                else:
                    segs.append((s, v, dn + v, dn))
                    dn += v
            tops = [sg for sg in segs if sg[1] > 0]
            top = tops[-1] if tops else None
            for s, v, a0, a1 in segs:
                c = color_for(spec, s["idx"])
                ya, yb = sy(a1), sy(a0)
                hh = yb - ya - (2 if (s, v, a0, a1) is not top or len(segs) > 1 else 0)
                bx = px0 + band * i + (band - bw) / 2
                gap_top = 0 if (s, v, a0, a1) is top else 2
                a, t = tip(cats[i], s["name"], vfmt(v) + ("" if k == "col_stack100" or unit == "%" else (" " + unit if unit else "")))
                end = "top" if (s, v, a0, a1) is top else "none"
                g.add(f'<path d="{bar_path(bx, ya + gap_top, bw, max(yb - ya - gap_top, 0.5), 4 if end == "top" else 0, end if end == "top" else "x")}" fill="{c}" {a}>{t}</path>')
                if labels == "all" and yb - ya >= 16 and bw >= tw(vfmt(v), 10.5) + 4:
                    g.text(bx + bw / 2, (ya + yb) / 2 + 4, vfmt(v), 10.5, text_on(c), "middle")
            if labels in ("auto", "all") and k == "col_stack" and up > 0 and n <= 16:
                g.text(px0 + band * i + band / 2, sy(up) - 5, vfmt(up), 11, INK, "middle")
    elif is_area:
        base_acc = [0.0] * n
        prev_top = None
        for si, s in enumerate(series):
            c = color_for(spec, s["idx"])
            vs = [v or 0 for v in s["values"]]
            if k == "area_stack":
                top = [base_acc[i] + vs[i] for i in range(n)]
                bottom = base_acc[:]
            else:
                top, bottom = vs, [0] * n
            pts_t = [(cx(cats[i]), sy(top[i])) for i in range(n)]
            pts_b = [(cx(cats[i]), sy(bottom[i])) for i in reversed(range(n))]
            path = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts_t + pts_b) + "Z"
            op = 0.85 if k == "area_stack" else 0.12
            g.add(f'<path d="{path}" fill="{c}" fill-opacity="{op}"/>')
            stroke = SURFACE if k == "area_stack" else c
            g.add(f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts_t)}" fill="none" stroke="{stroke}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>')
            for i in range(n):
                a, t = tip(cats[i], s["name"] if ns > 1 else "", vfmt(s["values"][i]) + (" " + unit if unit and unit != "%" else ""))
                g.add(f'<rect x="{px0 + band * i:.1f}" y="{min(pts_t[i][1], sy(bottom[i])) - 6:.1f}" width="{band:.1f}" height="{abs(sy(bottom[i]) - pts_t[i][1]) + 12:.1f}" fill="transparent" {a}>{t}</rect>')
            if k == "area_stack":
                base_acc = top
        if labels != "none" and k == "area":
            s = series[0]
            last = max(i for i in range(n) if s["values"][i] is not None)
            g.text(cx(cats[last]) - 4, sy(s["values"][last]) - 8, vfmt(s["values"][last]), 11.5, INK, "end")
    else:  # line, index
        ends = []
        for si, s in enumerate(series):
            c = color_for(spec, s["idx"])
            pts = [(cx(cats[i]), sy(v), i) for i, v in enumerate(s["values"]) if v is not None]
            if not pts:
                continue
            segs, cur, lasti = [], [], None
            for p in pts:
                if lasti is not None and p[2] != lasti + 1:
                    segs.append(cur)
                    cur = []
                cur.append(p)
                lasti = p[2]
            segs.append(cur)
            for sg in segs:
                g.add(f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y, _ in sg)}" fill="none" stroke="{c}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>')
            dots = n <= 30
            for x, y, i in pts:
                a, t = tip(cats[i], s["name"] if ns > 1 else "", vfmt(s["values"][i]) + (" " + unit if unit and unit != "%" else ""))
                if dots:
                    g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{c}" stroke="{SURFACE}" stroke-width="2" {a}>{t}</circle>')
                else:
                    g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="8" fill="transparent" {a}>{t}</circle>')
                if labels == "all" and n <= 24 and ns <= 2:
                    g.text(x, y - 9, vfmt(s["values"][i]), 10.5, INK, "middle")
            ends.append((pts[-1][1], vfmt(s["values"][pts[-1][2]]), pts[-1][0]))
        if labels == "auto" and len(ends) <= 4:
            ys_sorted = sorted(e[0] for e in ends)
            if all(b - a >= 13 for a, b in zip(ys_sorted, ys_sorted[1:])):
                for y, s, x in ends:
                    g.text(x + 8, y + 4, s, 11.5, INK, "start") if x + 8 + tw(s, 11.5) < w else g.text(x, y - 10, s, 11.5, INK, "middle")
    return g.out()


def _render_horizontal(g, spec, k, cats, series, lo, hi, ticks, step, unit, vtitle, ctitle, labels, vfmt, box):
    x0, y0, x1, y1 = box
    n, ns = len(cats), len(series)
    labw = min(160, max(tw(c, 12) for c in cats) + 10)
    tick_labels = tick_fmt(ticks, step, unit)
    vt = vtitle
    bottom_room = 20 + (18 if vt else 0)
    px0 = x0 + labw
    right_pad = 46 if labels != "none" else 10
    px1 = x1 - right_pad
    py0, py1 = y0 + 2, y1 - bottom_room
    sx = lambda v: px0 + (v - lo) / (hi - lo) * (px1 - px0)
    band = (py1 - py0) / n
    for t, s in zip(ticks, tick_labels):
        xx = sx(t)
        if abs(t) > 1e-12:
            g.line(xx, py0, xx, py1, GRID, 1)
        g.text(xx, py1 + 15, s, 11.5, MUTED, "middle")
    if lo <= 0 <= hi:
        g.line(sx(0), py0, sx(0), py1, BASE, 1)
    if vt:
        g.text((px0 + px1) / 2, py1 + 33, vt, 12, INK2, "middle")
    for i, c in enumerate(cats):
        g.text(px0 - 8, py0 + band * (i + .5) + 4, clip(c, 12, labw - 10), 12, INK2, "end")
    if k in ("bar", "bar_group"):
        gh = min(band * 0.78, ns * 22 + (ns - 1) * 2) if ns > 1 else min(24, band * 0.62)
        bh = (gh - (ns - 1) * 2) / ns if ns > 1 else gh
        for si, s in enumerate(series):
            col = color_for(spec, s["idx"])
            for i, v in enumerate(s["values"]):
                if v is None:
                    continue
                by = py0 + band * i + (band - gh) / 2 + si * (bh + 2)
                xa, xb = sx(min(v, 0)), sx(max(v, 0))
                a, t = tip(cats[i], s["name"] if ns > 1 else "", vfmt(v) + (" " + unit if unit and unit != "%" else ""))
                g.add(f'<path d="{bar_path(xa, by, max(xb - xa, 0.5), bh, 4, "right" if v >= 0 else "left")}" fill="{col}" {a}>{t}</path>')
                if labels == "all" or (labels == "auto" and (ns == 1 or n * ns <= 12)):
                    if v >= 0:
                        g.text(xb + 5, by + bh / 2 + 4, vfmt(v), 11 if ns > 1 else 11.5, INK)
                    else:
                        g.text(xa - 5, by + bh / 2 + 4, vfmt(v), 11, INK, "end")
    else:  # bar_stack
        bh = min(24, band * 0.62)
        for i in range(n):
            acc = 0.0
            segs = [(s, s["values"][i]) for s in series if s["values"][i]]
            segs = [(s, v) for s, v in segs if v > 0]
            for j, (s, v) in enumerate(segs):
                col = color_for(spec, s["idx"])
                xa, xb = sx(acc), sx(acc + v)
                last = j == len(segs) - 1
                by = py0 + band * i + (band - bh) / 2
                a, t = tip(cats[i], s["name"], vfmt(v) + (" " + unit if unit and unit != "%" else ""))
                g.add(f'<path d="{bar_path(xa, by, max(xb - xa - (0 if last else 2), 0.5), bh, 4 if last else 0, "right" if last else "x")}" fill="{col}" {a}>{t}</path>')
                if labels == "all" and xb - xa >= tw(vfmt(v), 10.5) + 8:
                    g.text((xa + xb) / 2, by + bh / 2 + 4, vfmt(v), 10.5, text_on(col), "middle")
                acc += v
            if labels in ("auto", "all") and acc > 0:
                g.text(sx(acc) + 5, py0 + band * (i + .5) + 4, vfmt(acc), 11.5, INK)
    return g.out()


# ── 원·도넛 ────────────────────────────────────────────────────────────────
def render_pie(table, spec, w, h):
    d = core.category_data(table, spec)
    s = d["series"][0]
    items = [(c, v) for c, v in zip(d["cats"], s["values"]) if v and v > 0]
    hidden = set(spec.get("hidden_cats") or [])
    items = [(c, v) for c, v in items if c not in hidden]
    g, box = frame(dict(spec, legend="none"), w, h, [])
    x0, y0, x1, y1 = box
    if not items:
        g.text(w / 2, h / 2, "양수 값이 없습니다", 13, MUTED, "middle")
        return g.out()
    tot = sum(v for _, v in items)
    unit = s.get("unit", "")
    dec = core.decimals_of([v for _, v in items])
    pct_only = unit == "%" and abs(tot - 100) < 0.6
    vtxt = lambda v: f"{v / tot * 100:.1f}%" if pct_only else f"{core.fmt(v, dec)}{'%' if unit == '%' else ''}  ·  {v / tot * 100:.1f}%"
    name_w = max(tw(c, 12.5) for c, _ in items)
    num_w = max(tw(vtxt(v), 12) for _, v in items)
    leg_w = min(380, name_w + num_w + 46)
    r = min((x1 - x0 - leg_w - 30) / 2, (y1 - y0) / 2 - 4)
    cx, cy = x0 + 10 + r, (y0 + y1) / 2
    hole = 0.58 if spec["kind"] == "doughnut" else 0
    ang = -math.pi / 2
    labels = spec.get("labels") or "auto"
    for i, (c, v) in enumerate(items):
        col = color_for(spec, i)
        frac = v / tot
        a2 = ang + frac * 2 * math.pi
        large = 1 if frac > 0.5 else 0
        p = lambda a, rr: (cx + rr * math.cos(a), cy + rr * math.sin(a))
        (ax, ay), (bx, by) = p(ang, r), p(a2, r)
        tp = tip(c, core.fmt(v, dec) + (" " + unit if unit and unit != "%" else ("%" if unit == "%" else "")), f"{frac * 100:.1f}%")
        if frac >= 0.9999:
            path = f"M{cx - r:.1f},{cy:.1f}a{r:.1f},{r:.1f} 0 1,0 {2 * r:.1f},0a{r:.1f},{r:.1f} 0 1,0 {-2 * r:.1f},0Z"
        elif hole:
            (cx2, cy2), (dx, dy) = p(a2, r * hole), p(ang, r * hole)
            path = (f"M{ax:.1f},{ay:.1f}A{r:.1f},{r:.1f} 0 {large},1 {bx:.1f},{by:.1f}L{cx2:.1f},{cy2:.1f}"
                    f"A{r * hole:.1f},{r * hole:.1f} 0 {large},0 {dx:.1f},{dy:.1f}Z")
        else:
            path = f"M{cx:.1f},{cy:.1f}L{ax:.1f},{ay:.1f}A{r:.1f},{r:.1f} 0 {large},1 {bx:.1f},{by:.1f}Z"
        g.add(f'<path d="{path}" fill="{col}" stroke="{SURFACE}" stroke-width="2" stroke-linejoin="round" {tp[0]}>{tp[1]}</path>')
        if labels != "none" and frac >= 0.06:
            mid = (ang + a2) / 2
            rr = r * ((1 + hole) / 2 if hole else 0.64)
            lx, ly = p(mid, rr)
            g.text(lx, ly + 4.5, f"{frac * 100:.0f}%" if frac >= 0.1 else f"{frac * 100:.1f}%", 12.5, text_on(col), "middle", 600)
        ang = a2
    if hole:
        g.text(cx, cy - 2, core.fmt(tot, dec) + ("%" if unit == "%" else ""), 20, INK, "middle", 600)
        g.text(cx, cy + 17, f"합계{(' (' + unit + ')') if unit and unit != '%' else ''}", 12, INK2, "middle")
    lx = cx + r + 34
    ly = cy - len(items) * 22 / 2 + 11
    for i, (c, v) in enumerate(items):
        col = color_for(spec, i)
        g.add(f'<rect x="{lx:.1f}" y="{ly - 6:.1f}" width="12" height="12" rx="3" fill="{col}"/>')
        g.text(lx + 19, ly + 4.5, clip(c, 12.5, x1 - lx - num_w - 34), 12.5, INK)
        g.text(x1, ly + 4.5, vtxt(v), 12, INK2, "end")
        ly += 22
    return g.out()


# ── 산점도 ────────────────────────────────────────────────────────────────
def render_scatter(table, spec, w, h):
    sers = [s for s in core.scatter_data(table, spec) if not s["hidden"]]
    multi = len(sers) > 1 or spec.get("group")
    legend = [(s["name"], color_for(spec, s["idx"]), "dot") for s in sers] if multi else []
    trend = spec.get("trend")
    g, box = frame(spec, w, h, legend)
    x0, y0, x1, y1 = box
    pts = [p for s in sers for p in s["points"]]
    if not pts:
        g.text(w / 2, h / 2, "점이 없습니다", 13, MUTED, "middle")
        return g.out()
    xu, yu = core.unit_of(table, spec["x"]), core.unit_of(table, spec["ys"][0])
    xlo, xhi, xt, xs = nice_domain(min(p[0] for p in pts), max(p[0] for p in pts), zero=False)
    ylo, yhi, yt, ys_ = nice_domain(min(p[1] for p in pts), max(p[1] for p in pts), zero=False)
    ytitle = axis_title(spec, "y_title", unit_title(spec["ys"][0] if not spec.get("group") or True else "", yu))
    xtitle = axis_title(spec, "x_title", unit_title(spec["x"], xu))
    room = 20 + (18 if xtitle else 0)
    px0, py0, sy = y_axis(g, (x0, y0, x1 - 8, y1 - room), ylo, yhi, yt, ys_, yu, ytitle)
    py1 = y1 - room
    px1 = x1 - 8
    sx = lambda v: px0 + (v - xlo) / (xhi - xlo) * (px1 - px0)
    for t, s in zip(xt, tick_fmt(xt, xs, xu)):
        g.line(sx(t), py0, sx(t), py1, GRID, 1) if t != xlo else None
        g.text(sx(t), py1 + 15, s, 11.5, MUTED, "middle")
    g.line(px0, py1, px1, py1, BASE, 1)
    if xtitle:
        g.text((px0 + px1) / 2, py1 + 33, xtitle, 12, INK2, "middle")
    for s in sers:
        c = color_for(spec, s["idx"])
        for x, y in s["points"]:
            a, t = tip(s["name"] if multi else "", f"{spec['x']} {core.fmt(x)}", f"{spec['ys'][0]} {core.fmt(y)}")
            g.add(f'<circle cx="{sx(x):.1f}" cy="{sy(y):.1f}" r="4.5" fill="{c}" fill-opacity="0.9" stroke="{SURFACE}" stroke-width="2" {a}>{t}</circle>')
    if trend:
        targets = sers if multi else sers[:1]
        for s in targets:
            lr = core.linreg([p[0] for p in s["points"]], [p[1] for p in s["points"]])
            if not lr:
                continue
            xa, xb = min(p[0] for p in s["points"]), max(p[0] for p in s["points"])
            ya, yb = lr["slope"] * xa + lr["intercept"], lr["slope"] * xb + lr["intercept"]
            g.add(f'<line x1="{sx(xa):.1f}" y1="{sy(ya):.1f}" x2="{sx(xb):.1f}" y2="{sy(yb):.1f}" stroke="{INK2 if not multi else color_for(spec, s["idx"])}" stroke-width="1.5" stroke-linecap="round" opacity="0.85"/>')
            if not multi:
                lbl = f"추세선  R² = {lr['r2']:.2f}"
                # 기울기 반대편 구석(양의 기울기 → 오른쪽 아래)에 둬서 점·선과 겹치지 않게
                ty = py1 - 10 if lr["slope"] >= 0 else py0 + 14
                g.add(f'<line x1="{px1 - tw(lbl, 11.5) - 26:.1f}" y1="{ty - 4:.1f}" x2="{px1 - tw(lbl, 11.5) - 8:.1f}" y2="{ty - 4:.1f}" stroke="{INK2}" stroke-width="1.5"/>')
                g.text(px1 - 2, ty, lbl, 11.5, INK2, "end")
    return g.out()


# ── 히스토그램·상자그림 ─────────────────────────────────────────────────────
def render_hist(table, spec, w, h):
    s = core.dist_data(table, spec)[0]
    bins = core.hist_bins(s["values"], int(spec.get("bins") or 0) or None)
    g, box = frame(spec, w, h, [])
    x0, y0, x1, y1 = box
    if not bins:
        return g.out()
    unit = core.unit_of(table, s["name"])
    lo, hi, ticks, step = nice_domain(0, max(b[2] for b in bins), True)
    if step < 1:
        lo, hi, ticks, step = 0, max(1, hi), list(range(0, int(max(1, hi)) + 1)), 1
    xtitle = axis_title(spec, "x_title", unit_title(s["name"], unit))
    room = 20 + (18 if xtitle else 0)
    px0, py0, sy = y_axis(g, (x0, y0, x1, y1 - room), lo, hi, ticks, step, "", axis_title(spec, "y_title", "빈도 (개)"))
    py1 = y1 - room
    a0, a1 = bins[0][0], bins[-1][1]
    sx = lambda v: px0 + (v - a0) / ((a1 - a0) or 1) * (x1 - px0)
    c = color_for(spec, 0)
    edges = [b[0] for b in bins] + [bins[-1][1]]
    ew = core.decimals_of(edges)
    every = max(1, math.ceil(len(edges) * tw(core.fmt(max(edges, key=abs), ew), 11.5) / (x1 - px0 - 20)))
    for i, e in enumerate(edges):
        if i % every == 0:
            g.text(sx(e), py1 + 15, core.fmt(e, ew), 11.5, MUTED, "middle")
    for lo_, hi_, cnt in bins:
        xa, xb = sx(lo_) + 1, sx(hi_) - 1
        yy = sy(cnt)
        a, t = tip(f"{core.fmt(lo_, ew)} ~ {core.fmt(hi_, ew)}{(' ' + unit) if unit else ''}", f"{cnt}개")
        if cnt:
            g.add(f'<path d="{bar_path(xa, yy, max(xb - xa, 0.5), py1 - yy, 3, "top")}" fill="{c}" {a}>{t}</path>')
            if (spec.get("labels") or "auto") != "none" and xb - xa >= tw(str(cnt), 11) + 2 and len(bins) <= 16:
                g.text((xa + xb) / 2, yy - 5, str(cnt), 11, INK, "middle")
    g.line(px0, py1, x1, py1, BASE, 1)
    if xtitle:
        g.text((px0 + x1) / 2, py1 + 33, xtitle, 12, INK2, "middle")
    st = core.box_stats(s["values"])
    if st:
        g.text(x1, y0 + 2, f"n={st['n']} · 평균 {core.fmt(st['mean'], ew + 1)} · 중앙값 {core.fmt(st['median'], ew + 1)}", 11.5, INK2, "end")
    return g.out()


def render_box(table, spec, w, h):
    sers = [s for s in core.dist_data(table, spec) if not s["hidden"]]
    stats = [(s, core.box_stats(s["values"])) for s in sers]
    stats = [(s, st) for s, st in stats if st]
    g, box = frame(spec, w, h, [])
    x0, y0, x1, y1 = box
    if not stats:
        return g.out()
    unit = core.unit_of(table, spec["ys"][0])
    allv = [v for _, st in stats for v in (st["min"], st["max"])]
    lo, hi, ticks, step = nice_domain(min(allv), max(allv), zero=False)
    ytitle = axis_title(spec, "y_title", unit_title(spec["ys"][0] if spec.get("group") or len(spec["ys"]) == 1 else "", unit))
    cats = [s["name"] for s, _ in stats]
    xtitle = axis_title(spec, "x_title", spec.get("group") or "")
    n = len(stats)
    room = x_label_room(cats, (x1 - x0 - 60) / n, xtitle)
    px0, py0, sy = y_axis(g, (x0, y0, x1, y1 - room), lo, hi, ticks, step, unit, ytitle)
    py1 = y1 - room
    band = (x1 - px0) / n
    x_cat_labels(g, cats, lambda c: px0 + band * (cats.index(c) + .5), band, py1, px0, x1, xtitle)
    g.line(px0, py1, x1, py1, BASE, 1)
    dec = core.decimals_of(allv)
    one_color = not spec.get("group") and len({core.unit_of(table, y) for y in spec["ys"]}) == 1 and False
    for i, (s, st) in enumerate(stats):
        c = color_for(spec, 0 if one_color else s["idx"])
        cx = px0 + band * (i + .5)
        bw = min(44, band * 0.5)
        a, t = tip(s["name"], f"n={st['n']}", f"최소 {core.fmt(st['min'], dec)}", f"Q1 {core.fmt(st['q1'], dec)}",
                   f"중앙값 {core.fmt(st['median'], dec)}", f"Q3 {core.fmt(st['q3'], dec)}", f"최대 {core.fmt(st['max'], dec)}")
        g.line(cx, sy(st["whi"]), cx, sy(st["q3"]), INK2, 1.5)
        g.line(cx, sy(st["q1"]), cx, sy(st["wlo"]), INK2, 1.5)
        g.line(cx - bw * .25, sy(st["whi"]), cx + bw * .25, sy(st["whi"]), INK2, 1.5)
        g.line(cx - bw * .25, sy(st["wlo"]), cx + bw * .25, sy(st["wlo"]), INK2, 1.5)
        yq3, yq1 = sy(st["q3"]), sy(st["q1"])
        g.add(f'<rect x="{cx - bw / 2:.1f}" y="{yq3:.1f}" width="{bw:.1f}" height="{max(yq1 - yq3, 1):.1f}" rx="4" fill="{c}" fill-opacity="0.22" stroke="{c}" stroke-width="2" {a}>{t}</rect>')
        g.line(cx - bw / 2 + 1, sy(st["median"]), cx + bw / 2 - 1, sy(st["median"]), INK, 2.5)
        for o in st["out"]:
            g.add(f'<circle cx="{cx:.1f}" cy="{sy(o):.1f}" r="3.5" fill="{SURFACE}" stroke="{c}" stroke-width="1.8"><title>이상값 {core.fmt(o, dec)}</title></circle>')
        if (spec.get("labels") or "auto") != "none":
            g.text(cx + bw / 2 + 5, sy(st["median"]) + 4, core.fmt(st["median"], dec), 11, INK, "start")
    return g.out()


# ── 히트맵 ────────────────────────────────────────────────────────────────
def render_heatmap(table, spec, w, h):
    d = core.category_data(table, spec)
    cats, series = d["cats"], [s for s in d["series"] if not s["hidden"]]
    g, box = frame(spec, w, h, [])
    x0, y0, x1, y1 = box
    vals = [v for s in series for v in s["values"] if v is not None]
    if not vals:
        return g.out()
    lo, hi = min(vals), max(vals)
    unit = series[0].get("unit", "")
    dec = core.decimals_of(vals)
    labw = min(150, max(tw(s["name"], 12) for s in series) + 12)
    xtitle = axis_title(spec, "x_title", spec.get("x") or "")
    gx0 = x0 + labw
    leg_h = 34
    n, m = len(cats), len(series)
    cw = (x1 - gx0) / n
    room = x_label_room(cats, cw, xtitle)
    ch = min(58, (y1 - y0 - room - leg_h) / m)
    gy0 = y0 + 2
    for j, s in enumerate(series):
        g.text(gx0 - 8, gy0 + ch * (j + .5) + 4, clip(s["name"], 12, labw - 12), 12, INK2, "end")
        for i, v in enumerate(s["values"]):
            x, y = gx0 + cw * i, gy0 + ch * j
            if v is None:
                g.add(f'<rect x="{x + 1:.1f}" y="{y + 1:.1f}" width="{cw - 2:.1f}" height="{ch - 2:.1f}" rx="3" fill="#f0efec"/>')
                continue
            f = (v - lo) / ((hi - lo) or 1)
            col = SEQ[min(len(SEQ) - 1, int(round(f * (len(SEQ) - 1))))]
            a, t = tip(cats[i], s["name"], core.fmt(v, dec) + (" " + unit if unit and unit != "%" else ("%" if unit == "%" else "")))
            g.add(f'<rect x="{x + 1:.1f}" y="{y + 1:.1f}" width="{max(cw - 2, .5):.1f}" height="{max(ch - 2, .5):.1f}" rx="3" fill="{col}" {a}>{t}</rect>')
            lbl = core.fmt(v, dec)
            if (spec.get("labels") or "auto") != "none" and cw >= tw(lbl, 11) + 8 and ch >= 18:
                g.text(x + cw / 2, y + ch / 2 + 4, lbl, 11, text_on(col), "middle")
    gy1 = gy0 + ch * m
    x_cat_labels(g, cats, lambda c: gx0 + cw * (cats.index(c) + .5), cw, gy1 - 4, gx0, x1, xtitle)
    # 색 범례
    ly = min(y1 - 14, gy1 + room + 20)
    lw = min(220, x1 - gx0)
    lx = x1 - lw
    stops = "".join(f'<stop offset="{i / (len(SEQ) - 1):.3f}" stop-color="{c}"/>' for i, c in enumerate(SEQ))
    g.add(f'<defs><linearGradient id="hmg" x1="0" x2="1" y1="0" y2="0">{stops}</linearGradient></defs>'
          f'<rect x="{lx:.1f}" y="{ly - 8:.1f}" width="{lw:.1f}" height="8" rx="2" fill="url(#hmg)"/>')
    g.text(lx, ly + 11, core.fmt(lo, dec), 11, MUTED, "start")
    g.text(lx + lw, ly + 11, core.fmt(hi, dec) + (f" {unit}" if unit else ""), 11, MUTED, "end")
    g.h = min(h, int(ly + 30))  # 행이 적으면 아래 빈 공간을 잘라 낸다
    return g.out()


# ── 단위별 2단 ────────────────────────────────────────────────────────────
def render_panels(table, spec, w, h):
    """단위가 다른 두 계열: 같은 x 축을 공유하는 위·아래 두 그래프 (이중축 대신)"""
    d = core.category_data(table, spec)
    cats, series = d["cats"], [s for s in d["series"] if not s["hidden"]][:2]
    g, box = frame(spec, w, h, [])
    x0, y0, x1, y1 = box
    if not series:
        return g.out()
    is_line = d["xtype"] == "date" and len(cats) > 6 or spec.get("panel_shape") == "line"
    xtitle = axis_title(spec, "x_title", spec.get("x") or "")
    n = len(cats)
    room = x_label_room(cats, (x1 - x0 - 60) / n, xtitle)
    gap = 24
    ph = (y1 - y0 - room - gap * (len(series) - 1)) / len(series)
    # 두 패널의 왼쪽 여백을 맞춘다
    doms = []
    for s in series:
        vs = [v for v in s["values"] if v is not None] or [0]
        lo, hi = min(vs), max(vs)
        doms.append(nice_domain(lo, hi, zero=not is_line or (lo >= 0 and lo < 0.35 * hi), target=4))
    lw = max(max(tw(t, 11.5) for t in tick_fmt(dm[2], dm[3], s.get("unit", ""))) for dm, s in zip(doms, series)) + 8
    px0 = x0 + lw
    band = (x1 - px0) / n
    cx = lambda i: px0 + band * (i + .5)
    for k, (s, (lo, hi, ticks, step)) in enumerate(zip(series, doms)):
        top = y0 + k * (ph + gap)
        unit = s.get("unit", "")
        c = color_for(spec, s["idx"])
        g.add(f'<rect x="{x0:.1f}" y="{top + 1:.1f}" width="12" height="10" rx="2" fill="{c}"/>')
        g.text(x0 + 18, top + 10, unit_title(s["name"], unit), 12, INK2, weight=600)
        ptop, pbot = top + 20, top + ph
        sy = lambda v, lo=lo, hi=hi, ptop=ptop, pbot=pbot: pbot - (v - lo) / (hi - lo) * (pbot - ptop)
        for t, lab in zip(ticks, tick_fmt(ticks, step, unit)):
            if abs(t) > 1e-12:
                g.line(px0, sy(t), x1, sy(t), GRID, 1)
            g.text(px0 - 7, sy(t) + 4, lab, 11.5, MUTED, "end")
        g.line(px0, sy(max(lo, 0)) if lo <= 0 <= hi else pbot, x1, sy(max(lo, 0)) if lo <= 0 <= hi else pbot, BASE, 1)
        dec = core.decimals_of([v for v in s["values"] if v is not None])
        if is_line:
            pts = [(cx(i), sy(v), v, i) for i, v in enumerate(s["values"]) if v is not None]
            g.add(f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y, _, _ in pts)}" fill="none" stroke="{c}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>')
            for x, y, v, i in pts:
                a, t = tip(cats[i], s["name"], core.fmt(v, dec, unit) + (" " + unit if unit and unit != "%" else ""))
                g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{c}" stroke="{SURFACE}" stroke-width="2" {a}>{t}</circle>')
            if pts and (spec.get("labels") or "auto") != "none":
                x, y, v, _ = pts[-1]
                g.text(x - 6, y - 9, core.fmt(v, dec, unit), 11.5, INK, "end")
        else:
            bw = min(24, band * 0.6)
            for i, v in enumerate(s["values"]):
                if v is None:
                    continue
                ya, yb = sy(max(v, 0) if lo <= 0 else v), sy(min(v, 0) if lo <= 0 else lo)
                a, t = tip(cats[i], s["name"], core.fmt(v, dec, unit) + (" " + unit if unit and unit != "%" else ""))
                g.add(f'<path d="{bar_path(cx(i) - bw / 2, ya, bw, max(yb - ya, .5), 4, "top")}" fill="{c}" {a}>{t}</path>')
                if (spec.get("labels") or "auto") != "none" and n <= 12:
                    g.text(cx(i), ya - 5, core.fmt(v, dec, unit), 11, INK, "middle")
    x_cat_labels(g, cats, lambda c: cx(cats.index(c)), band, y1 - room, px0, x1, xtitle)
    return g.out()


# ── 이중축 ────────────────────────────────────────────────────────────────
def dual_domains(lv, rv, zero=True, left_zero=True):
    """왼쪽·오른쪽 축 범위. 둘 다 0 을 품으면 0 아래·위 눈금 칸 수를 맞춰 0선과 격자를 공유한다.
    zero=False 면 오른쪽 축은 데이터 범위대로(정렬 안 함). 반환: ((lo, hi, ticks, step), (lo, hi, ticks, step), aligned)"""
    lv = [v for v in lv if v is not None] or [0]
    rv = [v for v in rv if v is not None] or [0]
    L = nice_domain(min(lv), max(lv), left_zero, 5)
    R = nice_domain(min(rv), max(rv), zero, 5)
    if not (L[0] <= 0 <= L[1] and R[0] <= 0 <= R[1]):
        return L, R, False
    cnt = lambda d: (round(d[1] / d[3]), round(-d[0] / d[3]))
    (pl, nl), (pr, nr) = cnt(L), cnt(R)
    P, N = max(pl, pr), max(nl, nr)
    mk = lambda st: (-N * st, P * st, [round(i * st, 10) for i in range(-N, P + 1)], st)
    return mk(L[3]), mk(R[3]), True


def render_dual(table, spec, w, h):
    d = core.category_data(table, spec)
    cats = d["cats"]
    vis = [s for s in d["series"] if not s["hidden"]]
    left = [s for s in vis if not s.get("right")]
    right = [s for s in vis if s.get("right")]
    shape = spec.get("dual_shape") or "bar_line"
    legend = [(s["name"], color_for(spec, s["idx"]), "bar" if shape == "bar_line" else "line") for s in left] + \
             [(s["name"] + " (오른쪽 축)", color_for(spec, s["idx"]), "line") for s in right]
    g, box = frame(spec, w, h, legend)
    x0, y0, x1, y1 = box
    if not cats or not (left or right):
        g.text(w / 2, h / 2, "표시할 값이 없습니다", 13, MUTED, "middle")
        return g.out()
    lu = left[0]["unit"] if left else ""
    ru = right[0]["unit"] if right else ""
    zero = spec.get("zero_base", True) is not False
    L, R, aligned = dual_domains([v for s in left for v in s["values"]], [v for s in right for v in s["values"]], zero,
                                 left_zero=shape == "bar_line" or zero)
    llab, rlab = tick_fmt(L[2], L[3], lu), tick_fmt(R[2], R[3], ru)
    xtitle = axis_title(spec, "x_title", spec.get("x") or "")
    n = len(cats)
    lw = max(tw(s, 11.5) for s in llab) + 8
    rw = max(tw(s, 11.5) for s in rlab) + 8 if right else 0
    room = x_label_room(cats, (x1 - x0 - lw - rw) / n, xtitle)
    # 축 제목: 축마다 계열 색 표시(계열이 하나뿐인 축은 제목 글자도 그 색), 단위 포함
    lt = axis_title(spec, "y_title", unit_title(left[0]["name"] if len(left) == 1 else "", lu)) if left else ""
    rt = axis_title(spec, "y2_title", unit_title(right[0]["name"], ru)) if right else ""
    ty = y0 + 2
    if lt:
        c = darken(color_for(spec, left[0]["idx"])) if len(left) == 1 else INK2
        for j, s in enumerate(left):  # 왼쪽 축 계열마다 색 견본
            g.add(f'<rect x="{x0 + j * 13:.1f}" y="{ty - 9:.1f}" width="10" height="10" rx="2" fill="{color_for(spec, s["idx"])}"/>')
        g.text(x0 + 13 * len(left) + 2, ty, lt, 12, c, weight=600)
    if rt:
        c = color_for(spec, right[0]["idx"])
        g.text(x1 - 14, ty, rt, 12, darken(c), "end", 600)
        g.add(f'<line x1="{x1 - 10:.1f}" y1="{ty - 4:.1f}" x2="{x1:.1f}" y2="{ty - 4:.1f}" stroke="{c}" stroke-width="2.5" stroke-linecap="round"/>')
    py0, py1 = y0 + 16, y1 - room
    px0, px1 = x0 + lw, x1 - rw
    syL = lambda v: py1 - (v - L[0]) / (L[1] - L[0]) * (py1 - py0)
    syR = lambda v: py1 - (v - R[0]) / (R[1] - R[0]) * (py1 - py0)
    for t, s in zip(L[2], llab):
        if abs(t) > 1e-12:
            g.line(px0, syL(t), px1, syL(t), GRID, 1)
        g.text(px0 - 7, syL(t) + 4, s, 11.5, MUTED, "end")
    if right:
        for t, s in zip(R[2], rlab):
            if not aligned and abs(t) > 1e-12:
                pass  # 정렬하지 않으면 오른쪽 축은 격자 없이 눈금 글자만
            g.text(px1 + 7, syR(t) + 4, s, 11.5, MUTED, "start")
    if L[0] <= 0 <= L[1]:
        g.line(px0, syL(0), px1, syL(0), BASE, 1)
    band = (px1 - px0) / n
    cx = lambda i: px0 + band * (i + .5)
    x_cat_labels(g, cats, lambda c: cx(cats.index(c)), band, py1, px0, px1, xtitle)
    labels = spec.get("labels") or "auto"
    decL = core.decimals_of([v for s in left for v in s["values"] if v is not None])
    decR = core.decimals_of([v for s in right for v in s["values"] if v is not None])
    fL = lambda v: core.fmt(v, decL, lu)
    fR = lambda v: core.fmt(v, decR, ru)
    sfx = lambda u: "" if not u or u == "%" else " " + u
    if shape == "bar_line":
        k = len(left)
        gw = min(band * 0.7, k * 24 + (k - 1) * 2) if k > 1 else min(24, band * 0.6)
        bw = (gw - (k - 1) * 2) / k if k > 1 else gw
        for si, s in enumerate(left):
            c = color_for(spec, s["idx"])
            for i, v in enumerate(s["values"]):
                if v is None:
                    continue
                bx = px0 + band * i + (band - gw) / 2 + si * (bw + 2)
                ya, yb = syL(max(v, 0)), syL(min(v, 0))
                a, t = tip(cats[i], s["name"] + " (왼쪽 축)", fL(v) + sfx(lu))
                g.add(f'<path d="{bar_path(bx, ya, bw, max(yb - ya, .5), 4, "top" if v >= 0 else "bottom")}" fill="{c}" {a}>{t}</path>')
                last_clash = labels == "auto" and right and i == n - 1  # 오른쪽 선 끝 라벨과 겹치지 않게
                if labels == "all" or (labels == "auto" and k == 1 and n <= 10 and not last_clash):
                    g.text(bx + bw / 2, ya - 5 if v >= 0 else syL(0) - 5, fL(v), 10.5, INK, "middle")
    lines = ([(s, syR, fR, ru, "오른쪽") for s in right] +
             ([(s, syL, fL, lu, "왼쪽") for s in left] if shape == "line_line" else []))
    for s, sy, f, u, side in lines:
        c = color_for(spec, s["idx"])
        pts = [(cx(i), sy(v), v, i) for i, v in enumerate(s["values"]) if v is not None]
        if not pts:
            continue
        g.add(f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y, _, _ in pts)}" fill="none" stroke="{c}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>')
        for x, y, v, i in pts:
            a, t = tip(cats[i], f"{s['name']} ({side} 축)", f(v) + sfx(u))
            g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{c}" stroke="{SURFACE}" stroke-width="2" {a}>{t}</circle>')
            if labels == "all":
                g.text(x, y - 9, f(v), 10.5, INK, "middle")
        if labels == "auto":
            x, y, v, _ = pts[-1]
            g.text(x, y - 10, f(v), 11.5, INK, "middle")
    return g.out()


def render(table, spec, w=720, h=440):
    k = spec.get("kind")
    try:
        if k in ("pie", "doughnut"):
            return render_pie(table, spec, w, h)
        if k == "scatter":
            return render_scatter(table, spec, w, h)
        if k == "hist":
            return render_hist(table, spec, w, h)
        if k == "box":
            return render_box(table, spec, w, h)
        if k == "heatmap":
            return render_heatmap(table, spec, w, h)
        if k == "panels":
            return render_panels(table, spec, w, h)
        if k == "dual":
            return render_dual(table, spec, w, h)
        return render_category(table, spec, w, h)
    except (ValueError, KeyError, IndexError, ZeroDivisionError) as e:
        g = Svg(w, h, spec.get("title") or "")
        g.text(w / 2, h / 2, f"그릴 수 없음: {e}", 13, MUTED, "middle")
        return g.out()
