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


def on_bg(fill):
    """색 위에 얹은 값 라벨 표시 — 화면(ui.html)이 어두운 테마에서 바탕색에 맞춰 글자색을 고를 수 있게. 그림 모양은 같음"""
    return f'data-bg="{fill}"'


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
    elif shape == "dash":
        g.add(f'<line x1="{x:.1f}" y1="{y:.1f}" x2="{x + 16:.1f}" y2="{y:.1f}" stroke="{c}" stroke-width="2.5" stroke-dasharray="5 3"/>')
    elif shape == "dot":
        g.add(f'<circle cx="{x + 6:.1f}" cy="{y:.1f}" r="5" fill="{c}"/>')
    elif shape == "tick":
        g.add(f'<line x1="{x:.1f}" y1="{y:.1f}" x2="{x + 14:.1f}" y2="{y:.1f}" stroke="{c}" stroke-width="3.5" stroke-linecap="butt"/>')
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
    trends = core.bar_trends(spec, series) if k in core.TREND_KINDS else []
    legend += [(t["name"], trend_color(spec, t), "dash") for t in trends]
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
    tv = [v for t in trends for v in t["values"] if v is not None]
    if tv:
        lo, hi = min(lo, min(tv)), max(hi, max(tv))
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
        return _render_horizontal(g, spec, k, cats, series, lo, hi, ticks, step, unit, ytitle, xtitle, labels, vfmt, box, trends)

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
        for t in trends:
            si = next((j for j, s in enumerate(series) if s["idx"] == t["idx"]), 0)
            _trend_path(g, spec, t, [px0 + band * i + (band - gw) / 2 + si * (bw + 2) + bw / 2 for i in range(n)], sy, vfmt, cats, False)
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
                    g.text(bx + bw / 2, (ya + yb) / 2 + 4, vfmt(v), 10.5, text_on(c), "middle", extra=on_bg(c))
            if labels in ("auto", "all") and k == "col_stack" and up > 0 and n <= 16:
                g.text(px0 + band * i + band / 2, sy(up) - 5, vfmt(up), 11, INK, "middle")
        for t in trends:
            _trend_path(g, spec, t, [px0 + band * (i + .5) for i in range(n)], sy, vfmt, cats, False)
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


def trend_color(spec, t):
    return darken(color_for(spec, t["idx"]), 0.72) if t.get("idx") is not None else INK2


def _trend_path(g, spec, t, pos, scale, vfmt, cats, horiz):
    """추세선: 계열색을 어둡게 한 점선(아래 흰 테두리) — pos = 범주별 막대 가운데, scale = 값 → 좌표"""
    c = trend_color(spec, t)
    pts = [(p, scale(v)) if not horiz else (scale(v), p) for p, v in zip(pos, t["values"]) if v is not None]
    if len(pts) < 2:
        return
    d = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    g.add(f'<path d="{d}" fill="none" stroke="{SURFACE}" stroke-width="4.5" stroke-linejoin="round" stroke-opacity="0.85"/>')
    g.add(f'<path d="{d}" fill="none" stroke="{c}" stroke-width="2" stroke-dasharray="6 4" stroke-linejoin="round" stroke-linecap="round"/>')
    for (x, y), (i, v) in zip(pts, [(i, v) for i, v in enumerate(t["values"]) if v is not None]):
        a, tt = tip(cats[i], t["name"], vfmt(v))
        g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="6" fill="transparent" {a}>{tt}</circle>')


def _render_horizontal(g, spec, k, cats, series, lo, hi, ticks, step, unit, vtitle, ctitle, labels, vfmt, box, trends=()):
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
        for t in trends:
            si = next((j for j, s in enumerate(series) if s["idx"] == t["idx"]), 0)
            _trend_path(g, spec, t, [py0 + band * i + (band - gh) / 2 + si * (bh + 2) + bh / 2 for i in range(n)], sx, vfmt, cats, True)
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
                    g.text((xa + xb) / 2, by + bh / 2 + 4, vfmt(v), 10.5, text_on(col), "middle", extra=on_bg(col))
                acc += v
            if labels in ("auto", "all") and acc > 0:
                g.text(sx(acc) + 5, py0 + band * (i + .5) + 4, vfmt(acc), 11.5, INK)
        for t in trends:
            _trend_path(g, spec, t, [py0 + band * (i + .5) for i in range(n)], sx, vfmt, cats, True)
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
    semi = spec["kind"] == "semi_doughnut"
    if semi:  # 반원: 왼쪽(180°)에서 위를 지나 오른쪽까지, 가운데 아래에 합계
        r = min((x1 - x0 - leg_w - 30) / 2, y1 - y0 - 46)
        cx, cy = x0 + 10 + r, y0 + 8 + r
    else:
        r = min((x1 - x0 - leg_w - 30) / 2, (y1 - y0) / 2 - 4)
        cx, cy = x0 + 10 + r, (y0 + y1) / 2
    hole = 0.58 if spec["kind"] in ("doughnut", "semi_doughnut") else 0
    sweep = math.pi if semi else 2 * math.pi
    ang = math.pi if semi else -math.pi / 2
    labels = spec.get("labels") or "auto"
    for i, (c, v) in enumerate(items):
        col = color_for(spec, i)
        frac = v / tot
        a2 = ang + frac * sweep
        large = 1 if frac * sweep > math.pi else 0
        p = lambda a, rr: (cx + rr * math.cos(a), cy + rr * math.sin(a))
        (ax, ay), (bx, by) = p(ang, r), p(a2, r)
        tp = tip(c, core.fmt(v, dec) + (" " + unit if unit and unit != "%" else ("%" if unit == "%" else "")), f"{frac * 100:.1f}%")
        if frac >= 0.9999 and not semi:
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
            g.text(lx, ly + 4.5, f"{frac * 100:.0f}%" if frac >= 0.1 else f"{frac * 100:.1f}%", 12.5, text_on(col), "middle", 600, on_bg(col))
        ang = a2
    if semi:
        g.text(cx, cy - 8, core.fmt(tot, dec) + ("%" if unit == "%" else ""), 20, INK, "middle", 600)
        g.text(cx, cy + 12, f"합계{(' (' + unit + ')') if unit and unit != '%' else ''}", 12, INK2, "middle")
    elif hole:
        g.text(cx, cy - 2, core.fmt(tot, dec) + ("%" if unit == "%" else ""), 20, INK, "middle", 600)
        g.text(cx, cy + 17, f"합계{(' (' + unit + ')') if unit and unit != '%' else ''}", 12, INK2, "middle")
    lx = cx + r + 34
    ly = (cy - r / 2 if semi else cy) - len(items) * 22 / 2 + 11
    for i, (c, v) in enumerate(items):
        col = color_for(spec, i)
        g.add(f'<rect x="{lx:.1f}" y="{ly - 6:.1f}" width="12" height="12" rx="3" fill="{col}"/>')
        g.text(lx + 19, ly + 4.5, clip(c, 12.5, x1 - lx - num_w - 34), 12.5, INK)
        g.text(x1, ly + 4.5, vtxt(v), 12, INK2, "end")
        ly += 22
    if semi:
        g.h = min(h, int(max(cy + 34, ly + 10)))
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
                g.text(x + cw / 2, y + ch / 2 + 4, lbl, 11, text_on(col), "middle", extra=on_bg(col))
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


# ── 새 종류 공통: 범주 x 축 + 값 y 축 ─────────────────────────────────────────
def _cat_axes(g, box, cats, lo, hi, ticks, step, unit, ytitle, xtitle):
    """세로형 범주 플롯 틀. 반환: (px0, py0, py1, px1, band, cx(i), sy)"""
    x0, y0, x1, y1 = box
    n = max(1, len(cats))
    room = x_label_room(cats, (x1 - x0 - 60) / n, xtitle)
    px0, py0, sy = y_axis(g, (x0, y0, x1, y1 - room), lo, hi, ticks, step, unit, ytitle)
    py1 = y1 - room
    band = (x1 - px0) / n
    cx = lambda i: px0 + band * (i + 0.5)
    x_cat_labels(g, cats, lambda c: cx(cats.index(c)), band, py1, px0, x1, xtitle)
    return px0, py0, py1, x1, band, cx, sy


def _vis(d, role=None):
    return [s for s in d["series"] if not s["hidden"] and s.get("role") == role]


def _sfx(unit):
    return "" if not unit or unit == "%" else " " + unit


def _empty(g, w, h, msg="표시할 값이 없습니다"):
    g.text(w / 2, h / 2, msg, 13, MUTED, "middle")
    return g.out()


def _ytitle(spec, series, unit):
    t = axis_title(spec, "y_title", unit_title(series[0]["name"] if len(series) == 1 else "", unit))
    if spec.get("y_title") and unit and f"({unit})" not in t:
        t = f"{t} ({unit})"
    return t


# ── 방사형 ────────────────────────────────────────────────────────────────
def render_radar(table, spec, w, h):
    d = core.category_data(table, spec)
    cats, series = d["cats"], _vis(d)
    legend = [(s["name"], color_for(spec, s["idx"]), "line") for s in series]
    g, (x0, y0, x1, y1) = frame(spec, w, h, legend)
    n = len(cats)
    if n < 3 or not series:
        return _empty(g, w, h, "방사형은 항목(축)이 3개 이상 필요합니다")
    vals = [v for s in series for v in s["values"] if v is not None]
    if not vals:
        return _empty(g, w, h)
    unit = series[0].get("unit", "")
    lo, hi, ticks, step = nice_domain(min(vals), max(vals), True, 4)
    labw = min(120, max(tw(c, 12) for c in cats) + 10)
    r = max(40, min((x1 - x0) / 2 - labw, (y1 - y0) / 2 - 24))
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2 + 2
    ang = lambda i: -math.pi / 2 + 2 * math.pi * i / n
    rad = lambda v: r * (v - lo) / ((hi - lo) or 1)
    pt = lambda i, rr: (cx + rr * math.cos(ang(i)), cy + rr * math.sin(ang(i)))
    for t in ticks:
        if t <= lo + 1e-12:
            continue
        poly = " ".join(f"{x:.1f},{y:.1f}" for x, y in (pt(i, rad(t)) for i in range(n)))
        g.add(f'<polygon points="{poly}" fill="none" stroke="{GRID if t < hi else BASE}" stroke-width="1"/>')
    for i in range(n):
        ex, ey = pt(i, r)
        g.line(cx, cy, ex, ey, GRID, 1)
    for i, c in enumerate(cats):
        lx, ly = pt(i, r + 10)
        ca = math.cos(ang(i))
        anchor = "start" if ca > 0.25 else ("end" if ca < -0.25 else "middle")
        dy = 4 + (6 * math.sin(ang(i)) if anchor == "middle" else 0)
        g.text(lx, ly + dy, clip(c, 12, labw - 6), 12, INK2, anchor)
    labels = spec.get("labels") or "auto"
    dec = core.decimals_of(vals)
    op = 0.14 if len(series) <= 3 else 0.07
    for s in series:
        col = color_for(spec, s["idx"])
        pts = [pt(i, rad(v if v is not None else lo)) for i, v in enumerate(s["values"])]
        g.add(f'<polygon points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts)}" fill="{col}" fill-opacity="{op}" stroke="{col}" stroke-width="2" stroke-linejoin="round"/>')
        for i, ((x, y), v) in enumerate(zip(pts, s["values"])):
            a, t = tip(cats[i], s["name"] if len(series) > 1 else "", "값 없음" if v is None else core.fmt(v, dec, unit) + _sfx(unit))
            g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="{col}" stroke="{SURFACE}" stroke-width="1.5" {a}>{t}</circle>')
            if v is not None and (labels == "all" or (labels == "auto" and len(series) == 1 and n <= 10)):
                ox, oy = math.cos(ang(i)) * 12, math.sin(ang(i)) * 12
                g.text(x + ox, y + oy + 4, core.fmt(v, dec, unit), 10.5, INK, "middle")
    halo = f'stroke="{SURFACE}" stroke-width="3" paint-order="stroke"'
    for t, lab in zip(ticks, tick_fmt(ticks, step, unit)):  # 눈금 글자는 맨 위에 바탕색 테를 둘러 선·점과 겹쳐도 읽히게
        if t <= lo + 1e-12:
            continue
        g.text(cx + 5, cy - rad(t) + 12, lab, 10.5, MUTED, "start", extra=halo)
    return g.out()


# ── 스파크라인 ──────────────────────────────────────────────────────────────
def render_sparkline(table, spec, w, h):
    d = core.category_data(table, spec)
    cats, series = d["cats"], _vis(d)
    g, (x0, y0, x1, y1) = frame(spec, w, h, [])
    if not cats or not series:
        return _empty(g, w, h)
    labels = spec.get("labels") or "auto"
    show_v = labels != "none"
    m = len(series)
    head_h = 20
    row_h = max(30, min(56, (y1 - y0 - head_h) / m))
    need = y0 + head_h + row_h * m + 14
    if need > h:
        g.h = int(need)
    names = [unit_title(s["name"], s.get("unit", "")) for s in series]
    nw = min(190, max(tw(nm, 12) for nm in names) + 12)
    vw = 150 if show_v else 0
    sx0, sx1 = x0 + nw, x1 - vw - 10
    n = len(cats)
    sx = lambda i: sx0 + (sx1 - sx0) * (i / (n - 1) if n > 1 else 0.5)
    g.text(x0, y0 + 12, "계열", 11, MUTED)
    g.text(sx0, y0 + 12, cats[0], 11, MUTED)
    g.text(sx1, y0 + 12, cats[-1], 11, MUTED, "end")
    if show_v:
        g.text(x1 - vw + 70, y0 + 12, "마지막 값", 11, MUTED, "end")
        g.text(x1, y0 + 12, "처음 대비", 11, MUTED, "end")
    top = y0 + head_h
    for j, s in enumerate(series):
        ry = top + row_h * j
        g.line(x0, ry, x1, ry, GRID, 1)
        col = color_for(spec, s["idx"])
        unit = s.get("unit", "")
        vs = s["values"]
        pv = [(i, v) for i, v in enumerate(vs) if v is not None]
        g.text(x0, ry + row_h / 2 + 4, clip(names[j], 12, nw - 10), 12, INK2)
        if not pv:
            continue
        lo, hi = min(v for _, v in pv), max(v for _, v in pv)
        pad = 6
        sy = lambda v: ry + row_h - pad - (v - lo) / ((hi - lo) or 1) * (row_h - 2 * pad) if hi > lo else ry + row_h / 2
        dec = core.decimals_of([v for _, v in pv])
        segs, cur, last = [], [], None
        for i, v in pv:
            if last is not None and i != last + 1:
                segs.append(cur)
                cur = []
            cur.append((sx(i), sy(v)))
            last = i
        segs.append(cur)
        for sg in segs:
            g.add(f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in sg)}" fill="none" stroke="{col}" stroke-width="1.75" stroke-linejoin="round" stroke-linecap="round"/>')
        imax = max(pv, key=lambda p: p[1])[0]
        imin = min(pv, key=lambda p: p[1])[0]
        for i, v in pv:
            a, t = tip(cats[i], s["name"], core.fmt(v, dec, unit) + _sfx(unit))
            g.add(f'<circle cx="{sx(i):.1f}" cy="{sy(v):.1f}" r="6" fill="transparent" {a}>{t}</circle>')
        if hi > lo:
            for i in (imax, imin):
                g.add(f'<circle cx="{sx(i):.1f}" cy="{sy(vs[i]):.1f}" r="2.5" fill="{SURFACE}" stroke="{col}" stroke-width="1.5"/>')
        li, lv = pv[-1]
        g.add(f'<circle cx="{sx(li):.1f}" cy="{sy(lv):.1f}" r="3.5" fill="{col}" stroke="{SURFACE}" stroke-width="1.5"/>')
        if show_v:
            g.text(x1 - vw + 70, ry + row_h / 2 + 4, core.fmt(lv, dec, unit) + ("" if unit == "%" else ""), 12, INK, "end", 600)
            fv = pv[0][1]
            if fv:
                ch = (lv - fv) / abs(fv) * 100
                arrow = "▲" if ch > 0.05 else ("▼" if ch < -0.05 else "–")
                g.text(x1, ry + row_h / 2 + 4, f"{arrow} {abs(ch):.1f}%", 11.5, INK2, "end")
    g.line(x0, top + row_h * m, x1, top + row_h * m, GRID, 1)
    g.text(x0, top + row_h * m + 13, "계열마다 눈금이 따로입니다(높이로 계열끼리 크기를 비교하지 마세요) · ○ 최고·최저, ● 마지막 값", 10.5, MUTED)
    g.h = max(g.h if g.h > h else 0, min(h, int(top + row_h * m + 26))) or h
    return g.out()


# ── 캔들스틱·OHLC 막대 (+ 이동평균·볼린저·매물대·거래량/MACD/RSI 칸) ──────────────
UP_COL, DOWN_COL = "#e34948", "#2a78d6"  # 한국 관례: 오름 빨강, 내림 파랑
UP_INTL, DOWN_INTL = "#1baf7a", "#e34948"  # 국제 관례: 오름 초록, 내림 빨강
MA_COLS = ["#eda100", "#e87ba4", "#4a3aa7", "#008300", "#eb6834", "#52514e"]  # 이동평균 선(짧은 창부터) — 상승·하락 색과 겹치지 않게
BB_COL, BB_FILL = "#898781", "#c3c2b7"
MACD_COL, SIG_COL, RSI_COL = "#4a3aa7", "#eda100", "#4a3aa7"
VP_COL, VP_VA, POC_COL = "#d8d6cd", "#a9a79c", "#eb6834"
PANEL_W = {"volume": 0.26, "macd": 0.36, "rsi": 0.32}  # 가격 칸 1.0 대비 높이
PANEL_EXTRA = {"volume": 0, "macd": 120, "rsi": 110}   # 칸을 붙일 때 그림 높이를 늘리는 양(px)


def updown_default(spec):
    return (UP_INTL, DOWN_INTL) if spec.get("colors_updown") == "intl" else (UP_COL, DOWN_COL)


def candle_colors(spec):
    """(상승색, 하락색) — 직접 고른 색(colors 0·1)이 먼저, 없으면 관례(colors_updown: kr 빨강/파랑 · intl 초록/빨강)"""
    cs = spec.get("colors") or {}
    get = (lambda i: cs[i] if isinstance(cs, list) and i < len(cs) and cs[i] else None) if isinstance(cs, list) else (lambda i: cs.get(str(i)))
    u, d = updown_default(spec)
    return get(0) or u, get(1) or d


def ma_color(i):
    return MA_COLS[i % len(MA_COLS)]


def _path(xs, ys):
    """빈 값에서 끊기는 선 경로"""
    out, pen = [], False
    for x, y in zip(xs, ys):
        if y is None:
            pen = False
            continue
        out.append(f"{'L' if pen else 'M'}{x:.1f},{y:.1f}")
        pen = True
    return "".join(out)


def render_candle(table, spec, w, h):
    import fin
    k = spec.get("kind")
    up, dn = candle_colors(spec)
    try:
        F = fin.compute(table, spec)
    except (ValueError, KeyError):
        F = None
    op = F["opt"] if F else fin.options(spec, 0, False)
    legend = [("상승 (종가 ≥ 시가)", up, "bar"), ("하락 (종가 < 시가)", dn, "bar")]
    if F:
        legend += [(f"이동평균 {wnd}", ma_color(i), "tick") for i, (wnd, _) in enumerate(F["ma"])]
        if F["bb"]:
            legend.append((f"볼린저 밴드 ({op['bb']['period']}, {op['bb']['k']:g}σ)", BB_FILL, "bar"))
        if F["vp"]:
            legend.append(("매물대 POC", POC_COL, "bar"))
    panels = op["panels"] if F else []
    H = h + sum(PANEL_EXTRA[p] for p in panels)
    g, (x0, y0, x1, y1) = frame(spec, w, H, legend)
    if not F:
        return _empty(g, w, H, "시가·고가·저가·종가 4개 열이 필요합니다")
    cats, O, Hh, L, C = F["cats"], F["O"], F["H"], F["L"], F["C"]
    unit = F["unit"]
    vol = F["vol"]
    lows = [v for v in L + O + C if v is not None]
    highs = [v for v in Hh + O + C if v is not None]
    if not lows:
        return _empty(g, w, H)
    over = [v for _, s in F["ma"] for v in s if v is not None]
    if F["bb"]:
        over += [v for v in F["bb"][1] + F["bb"][2] if v is not None]
    lo, hi, ticks, step = nice_domain(min(lows + over), max(highs + over), False, 7 if over else 5)
    xtitle = axis_title(spec, "x_title", spec.get("x") or "")
    ytitle = axis_title(spec, "y_title", unit_title("가격", unit))
    n = len(cats)
    dec = core.decimals_of(lows + highs)
    f = lambda v: core.fmt(v, dec, unit)
    # 칸 나누기: 가격 칸 + 아래 칸들(같은 x 축), 맨 아래에 x 라벨
    room = 20 + (22 if xtitle else 0)
    gap = 16
    if ytitle:
        g.text(x0, y0 + 2, ytitle, 12, INK2)
        y0 += 14
    avail = y1 - room - y0 - gap * len(panels)
    wsum = 1 + sum(PANEL_W[p] for p in panels)
    ph = {p: avail * PANEL_W[p] / wsum for p in panels}
    py0, py1 = y0 + 4, y0 + avail / wsum
    # 눈금 라벨 폭 — 모든 칸이 같은 왼쪽 끝을 쓰게
    labs = tick_fmt(ticks, step, unit)
    lw = max(tw(s, 11.5) for s in labs) + 8
    sub = {}
    if "volume" in panels:
        vv = [v for v in F["V"] if v is not None] or [0]
        sub["volume"] = nice_domain(0, max(vv), True, 2)
    if "macd" in panels:
        mv = [v for s in F["macd"] for v in s if v is not None] or [0]
        sub["macd"] = nice_domain(min(mv), max(mv), True, 3)
    if "rsi" in panels:
        sub["rsi"] = (0, 100, [30, 70], 10)
    for p, (a_, b_, t_, s_) in sub.items():
        if p != "rsi":
            lw = max(lw, max(tw(s, 10.5) for s in tick_fmt(t_, s_, vol.get("unit", "") if p == "volume" else "")) + 8)
    px0 = x0 + lw
    vpw = min(150, (x1 - px0) * 0.2) if F["vp"] else 0
    px1 = x1 - (vpw + 10 if vpw else 0)
    band = (px1 - px0) / n
    cx = lambda i: px0 + band * (i + .5)
    bw = max(1.2, min(14, band * 0.62))
    lab_every = _x_every(cats, band)
    sy = lambda v: py1 - (v - lo) / ((hi - lo) or 1) * (py1 - py0)
    # 가격 칸 격자·눈금
    for t, s in zip(ticks, labs):
        g.line(px0, sy(t), px1, sy(t), GRID, 1)
        g.text(px0 - 7, sy(t) + 4, s, 11.5, MUTED, "end")
    for i in range(0, n, lab_every):  # 세로 보조선(날짜 라벨 자리) — 모든 칸에
        g.line(cx(i), py0, cx(i), y1 - room, "#f0efea", 1)
    # 볼린저 밴드(채운 띠 + 위·아래 점선 + 중심선)
    if F["bb"]:
        mid, bu, bl = F["bb"]
        seg = [i for i in range(n) if bu[i] is not None and bl[i] is not None]
        if seg:
            top = " ".join(f"{cx(i):.1f},{sy(bu[i]):.1f}" for i in seg)
            bot = " ".join(f"{cx(i):.1f},{sy(bl[i]):.1f}" for i in reversed(seg))
            g.add(f'<polygon points="{top} {bot}" fill="{BB_FILL}" fill-opacity="0.28" stroke="none"/>')
            for s_, dash in ((bu, "4 3"), (bl, "4 3"), (mid, "")):
                g.add(f'<path d="{_path([cx(i) for i in range(n)], [None if v is None else sy(v) for v in s_])}" fill="none" stroke="{BB_COL}" '
                      f'stroke-width="1.25"{f" stroke-dasharray={chr(34)}{dash}{chr(34)}" if dash else ""}/>')
    # 캔들 / OHLC 막대
    for i in range(n):
        o, hh, ll, c = O[i], Hh[i], L[i], C[i]
        if None in (o, c):
            continue
        hh = max(x for x in (hh, o, c) if x is not None)
        ll = min(x for x in (ll, o, c) if x is not None)
        col = up if c >= o else dn
        extra = [f"이동평균 {wnd} {f(s[i])}" for wnd, s in F["ma"] if s[i] is not None]
        a, t = tip(cats[i], f"시가 {f(o)}", f"고가 {f(hh)}", f"저가 {f(ll)}", f"종가 {f(c)}",
                   f"{vol['name']} {core.fmt(F['V'][i])}" if vol and F["V"][i] is not None else "", *extra)
        if k == "ohlc":
            sw = 1.6 if band >= 4 else 1.1
            tk = max(1.5, bw / 2)
            g.add(f'<path d="M{cx(i):.1f},{sy(hh):.1f}V{sy(ll):.1f}M{cx(i) - tk:.1f},{sy(o):.1f}H{cx(i):.1f}M{cx(i):.1f},{sy(c):.1f}H{cx(i) + tk:.1f}" '
                  f'stroke="{col}" stroke-width="{sw}" fill="none" {a}>{t}</path>')
        else:
            g.line(cx(i), sy(hh), cx(i), sy(ll), col, 1.25 if band >= 4 else 1)
            ya, yb = sy(max(o, c)), sy(min(o, c))
            g.add(f'<rect x="{cx(i) - bw / 2:.1f}" y="{ya:.1f}" width="{bw:.1f}" height="{max(yb - ya, 1.2):.1f}" rx="{1 if bw > 3 else 0}" fill="{col}" {a}>{t}</rect>')
    # 이동평균
    for j, (wnd, s_) in enumerate(F["ma"]):
        g.add(f'<path d="{_path([cx(i) for i in range(n)], [None if v is None else sy(v) for v in s_])}" fill="none" stroke="{ma_color(j)}" '
              f'stroke-width="1.75" stroke-linejoin="round" stroke-linecap="round"/>')
    if (spec.get("labels") or "auto") != "none":
        li = max((i for i in range(n) if C[i] is not None), default=None)
        if li is not None:
            top_ = max(x for x in (Hh[li], O[li], C[li]) if x is not None)
            g.text(min(cx(li), px1 - tw(f(C[li]), 11.5) / 2), max(py0 + 10, sy(top_) - 7), f"종가 {f(C[li])}", 11.5, INK,
                   "end" if cx(li) + 40 > px1 else "middle")
    # 매물대 (오른쪽, 같은 가격 눈금)
    if F["vp"]:
        vp = F["vp"]
        vx0 = px1 + 10
        mx = max(vp["vols"]) or 1
        g.text(vx0, py0 - 2, clip(f"매물대 ({vol['name']})", 11, vpw), 11, INK2)
        d_, u_ = vp["va"]
        vdec = core.decimals_of(vp["edges"])
        for i, v in enumerate(vp["vols"]):
            a_, b_ = vp["edges"][i], vp["edges"][i + 1]
            ya, yb = sy(b_), sy(a_)
            col = POC_COL if i == vp["poc"] else (VP_VA if (op["vprofile"]["va"] and d_ <= i <= u_) else VP_COL)
            at, t = tip(f"{core.fmt(a_, vdec)}~{core.fmt(b_, vdec)}", f"{vol['name']} {core.fmt(round(v))}",
                        "POC (거래 최다)" if i == vp["poc"] else ("가치 영역 70%" if op["vprofile"]["va"] and d_ <= i <= u_ else ""))
            g.add(f'<rect x="{vx0:.1f}" y="{ya + 0.5:.1f}" width="{max(0.5, v / mx * (vpw - 4)):.1f}" height="{max(0.8, yb - ya - 1):.1f}" fill="{col}" {at}>{t}</rect>')
        pm = (vp["edges"][vp["poc"]] + vp["edges"][vp["poc"] + 1]) / 2
        g.line(px0, sy(pm), px1, sy(pm), POC_COL, 1, 'stroke-dasharray="2 3"')
        g.text(px1 - 4, sy(pm) - 4, f"POC {f(pm)}", 10.5, INK, "end")
        if op["vprofile"]["va"]:
            g.text(x1, py1 + 12, f"가치 영역 {f(vp['edges'][d_])}~{f(vp['edges'][u_ + 1])}", 10, MUTED, "end")
    # 아래 칸들
    yy = py1
    for p in panels:
        ya, yb = yy + gap, yy + gap + ph[p]
        yy = yb
        g.line(px0, ya, px1, ya, GRID, 1)
        lo_, hi_, t_, s_ = sub[p]
        ss = lambda v, lo_=lo_, hi_=hi_, ya=ya, yb=yb: yb - (v - lo_) / ((hi_ - lo_) or 1) * (yb - ya - 14)
        if p == "volume":
            g.text(px0 + 4, ya + 11, unit_title(vol["name"], vol.get("unit", "")), 11, INK2)
            for t, lab in zip(t_, tick_fmt(t_, s_, vol.get("unit", ""))):
                if t > 0 and t == t_[-1]:
                    g.line(px0, ss(t), px1, ss(t), GRID, 1)
                    g.text(px0 - 7, ss(t) + 4, lab, 10.5, MUTED, "end")
            for i, v in enumerate(F["V"]):
                if v is None:
                    continue
                o, c = O[i], C[i]
                col = up if (o is None or c is None or c >= o) else dn
                a, t = tip(cats[i], vol["name"], core.fmt(v) + _sfx(vol.get("unit", "")))
                g.add(f'<rect x="{cx(i) - bw / 2:.1f}" y="{ss(v):.1f}" width="{bw:.1f}" height="{max(yb - ss(v), .5):.1f}" fill="{col}" fill-opacity="0.45" {a}>{t}</rect>')
            g.line(px0, yb, px1, yb, BASE, 1)
        elif p == "macd":
            ml, sg, hs = F["macd"]
            fa, sl, sn = op["macd"]
            for t, lab in zip(t_, tick_fmt(t_, s_)):
                if lo_ <= t <= hi_:
                    g.line(px0, ss(t), px1, ss(t), BASE if abs(t) < 1e-12 else GRID, 1)
                    g.text(px0 - 7, ss(t) + 4, lab, 10.5, MUTED, "end")
            mdec = max(1, core.decimals_of([v for v in ml if v is not None], 2))
            for i, v in enumerate(hs):
                if v is None:
                    continue
                y_a, y_b = sorted((ss(0), ss(v)))
                a, t = tip(cats[i], f"MACD {core.fmt(ml[i], mdec)}", f"시그널 {core.fmt(sg[i], mdec)}", f"히스토그램 {core.fmt(v, mdec)}")
                g.add(f'<rect x="{cx(i) - bw / 2:.1f}" y="{y_a:.1f}" width="{bw:.1f}" height="{max(y_b - y_a, .5):.1f}" fill="{up if v >= 0 else dn}" fill-opacity="0.5" {a}>{t}</rect>')
            for s__, col in ((ml, MACD_COL), (sg, SIG_COL)):
                g.add(f'<path d="{_path([cx(i) for i in range(n)], [None if v is None else ss(v) for v in s__])}" fill="none" stroke="{col}" stroke-width="1.6"/>')
            tx = px0 + 4
            g.text(tx, ya + 11, f"MACD ({fa}, {sl}, {sn})", 11, INK2)
            tx += tw(f"MACD ({fa}, {sl}, {sn})", 11) + 14
            for nm, col in (("MACD", MACD_COL), ("시그널", SIG_COL)):
                g.line(tx, ya + 7, tx + 14, ya + 7, col, 2.5)
                g.text(tx + 18, ya + 11, nm, 10.5, INK2)
                tx += tw(nm, 10.5) + 32
            g.text(tx, ya + 11, "막대 = 히스토그램(MACD − 시그널)", 10.5, MUTED)
        elif p == "rsi":
            rv = F["rsi"]
            g.add(f'<rect x="{px0:.1f}" y="{ss(70):.1f}" width="{px1 - px0:.1f}" height="{ss(30) - ss(70):.1f}" fill="#f0efea"/>')
            for t in (30, 70):
                g.line(px0, ss(t), px1, ss(t), MUTED, 1, 'stroke-dasharray="4 3"')
                g.text(px0 - 7, ss(t) + 4, str(t), 10.5, MUTED, "end")
            g.add(f'<path d="{_path([cx(i) for i in range(n)], [None if v is None else ss(v) for v in rv])}" fill="none" stroke="{RSI_COL}" stroke-width="1.6"/>')
            for i, v in enumerate(rv):
                if v is not None and (i % max(1, n // 60) == 0 or i == n - 1):
                    a, t = tip(cats[i], f"RSI {v:.1f}")
                    g.add(f'<circle cx="{cx(i):.1f}" cy="{ss(v):.1f}" r="3" fill="transparent" {a}>{t}</circle>')
            last = next((v for v in reversed(rv) if v is not None), None)
            g.text(px0 + 4, ya + 11, f"RSI ({op['rsi']}, 와일더)" + (f" · 마지막 {last:.1f}" if last is not None else "") + " · 70 과열 / 30 침체", 11, INK2)
    _x_labels(g, cats, cx, lab_every, y1 - room, px0, px1, xtitle)
    return g.out()


def _x_every(cats, band):
    """긴 날짜 축: 가로 라벨이 겹치지 않게 몇 칸마다 하나씩"""
    maxw = max((tw(c, 11) for c in cats), default=0)
    return max(1, math.ceil((maxw + 16) / max(band, 0.5)))


def _x_labels(g, cats, cx, every, ybase, x0, x1, xtitle):
    g.line(x0, ybase, x1, ybase, BASE, 1)
    for i in range(0, len(cats), every):
        g.text(cx(i), ybase + 15, cats[i], 11, MUTED, "middle")
    if xtitle:
        g.text((x0 + x1) / 2, ybase + 38, xtitle, 12, INK2, "middle")


# ── 폭포 ──────────────────────────────────────────────────────────────────
WF_COLS = ("#1baf7a", "#e34948", "#2a78d6")  # 증가·감소·합계


def wf_colors(spec):
    cs = spec.get("colors") or {}
    out = []
    for i, dflt in enumerate(WF_COLS):
        c = (cs[i] if isinstance(cs, list) and i < len(cs) and cs[i] else None) if isinstance(cs, list) else cs.get(str(i))
        out.append(c or dflt)
    return out


def render_waterfall(table, spec, w, h):
    d = core.category_data(table, spec)
    steps = core.waterfall_steps(table, spec, d)
    cu, cd, ct = wf_colors(spec)
    legend = [("증가", cu, "bar"), ("감소", cd, "bar"), ("합계", ct, "bar")]
    g, box = frame(spec, w, h, legend)
    if not steps:
        return _empty(g, w, h)
    sers = _vis(d)
    unit = sers[0].get("unit", "") if sers else ""
    ends = [x for st in steps for x in (st["a"], st["b"])]
    lo, hi, ticks, step = nice_domain(min(ends), max(ends), True)
    cats = [st["cat"] for st in steps]
    ytitle = _ytitle(spec, sers[:1], unit)
    px0, py0, py1, px1, band, cx, sy = _cat_axes(g, box, cats, lo, hi, ticks, step, unit, ytitle, axis_title(spec, "x_title", spec.get("x") or ""))
    bw = min(34, band * 0.62)
    dec = core.decimals_of([st["value"] for st in steps])
    labels = spec.get("labels") or "auto"
    for i, st in enumerate(steps):
        col = {"up": cu, "down": cd, "total": ct}[st["kind"]]
        top, bot = max(st["a"], st["b"]), min(st["a"], st["b"])
        ya, yb = sy(top), sy(bot)
        sign = "+" if st["kind"] == "up" else ""
        txt = (sign + core.fmt(st["value"], dec, unit)) if st["kind"] != "total" else core.fmt(st["value"], dec, unit)
        a, t = tip(st["cat"], {"up": "증가", "down": "감소", "total": "합계"}[st["kind"]], txt + _sfx(unit),
                   f"누적 {core.fmt(st['b'], dec, unit)}" if st["kind"] != "total" else "")
        g.add(f'<rect x="{cx(i) - bw / 2:.1f}" y="{ya:.1f}" width="{bw:.1f}" height="{max(yb - ya, 1):.1f}" rx="2" fill="{col}" {a}>{t}</rect>')
        if i + 1 < len(steps):
            nxt = steps[i + 1]
            lvl = st["b"]
            if nxt["kind"] != "total" or abs(nxt["b"] - lvl) < 1e-9:
                g.line(cx(i) + bw / 2, sy(lvl), cx(i + 1) - bw / 2, sy(lvl), BASE, 1, 'stroke-dasharray="3 3"')
        if labels != "none" and (labels == "all" or len(steps) <= 20) and band >= tw(txt, 10.5) - 4:
            fs = 10.5 if len(steps) > 10 else 11
            below = st["value"] < 0 and yb + 15 < py1  # 아래 자리가 없으면(축 바닥) 막대 위에
            g.text(cx(i), yb + 13 if below else ya - 5, txt, fs, INK, "middle")
    return g.out()


# ── 깔때기 ────────────────────────────────────────────────────────────────
def render_funnel(table, spec, w, h):
    d = core.category_data(table, spec)
    sers = _vis(d)
    g, (x0, y0, x1, y1) = frame(spec, w, h, [])
    if not sers:
        return _empty(g, w, h)
    s = sers[0]
    items = [(c, v) for c, v in zip(d["cats"], s["values"]) if v is not None and v > 0]
    if not items:
        return _empty(g, w, h, "양수 값이 없습니다")
    unit = s.get("unit", "")
    labels = spec.get("labels") or "auto"
    show = labels != "none"
    labw = min(170, max(tw(c, 12.5) for c, _ in items) + 14)
    rw = 150 if show else 0
    fx0, fx1 = x0 + labw, x1 - rw
    mid = (fx0 + fx1) / 2
    n = len(items)
    band = min(64, (y1 - y0) / n)
    bh = band * 0.72
    vmax = max(v for _, v in items)
    col = color_for(spec, s["idx"])
    dec = core.decimals_of([v for _, v in items])
    half = lambda v: (fx1 - fx0) / 2 * v / vmax
    if show:
        g.text(x1, y0 + 2, "이전 단계 대비 · 처음 대비", 11, MUTED, "end")
    for i, (c, v) in enumerate(items):
        top = y0 + 8 + band * i + (band - bh) / 2
        hw = half(v)
        if i + 1 < n:  # 다음 단계로 이어지는 연한 사다리꼴
            nv = items[i + 1][1]
            nt = y0 + 8 + band * (i + 1) + (band - bh) / 2
            g.add(f'<path d="M{mid - hw:.1f},{top + bh:.1f}L{mid + hw:.1f},{top + bh:.1f}L{mid + half(nv):.1f},{nt:.1f}L{mid - half(nv):.1f},{nt:.1f}Z" fill="{col}" fill-opacity="0.13"/>')
        a, t = tip(c, core.fmt(v, dec, unit) + _sfx(unit), f"처음 대비 {v / items[0][1] * 100:.1f}%")
        g.add(f'<rect x="{mid - hw:.1f}" y="{top:.1f}" width="{max(2 * hw, 1):.1f}" height="{bh:.1f}" rx="3" fill="{col}" {a}>{t}</rect>')
        g.text(fx0 - 10, top + bh / 2 + 4.5, clip(c, 12.5, labw - 14), 12.5, INK2, "end")
        if show:
            vt = core.fmt(v, dec, unit) + _sfx(unit)
            if 2 * hw >= tw(vt, 12) + 10:
                g.text(mid, top + bh / 2 + 4.5, vt, 12, text_on(col), "middle", 600, on_bg(col))
            else:
                g.text(mid + hw + 6, top + bh / 2 + 4.5, vt, 12, INK, "start", 600)
            conv = "" if i == 0 else f"{v / items[i - 1][1] * 100:.1f}% · {v / items[0][1] * 100:.1f}%"
            if conv:
                g.text(x1, top + bh / 2 + 4.5, conv, 11.5, INK2, "end")
    g.h = min(h, int(y0 + 8 + band * n + 16))
    return g.out()


# ── 거품 ──────────────────────────────────────────────────────────────────
def render_bubble(table, spec, w, h):
    sers = [s for s in core.bubble_data(table, spec) if not s["hidden"]]
    multi = bool(spec.get("group")) and len(sers) > 1
    legend = [(s["name"], color_for(spec, s["idx"]), "dot") for s in sers] if multi else []
    g, (x0, y0, x1, y1) = frame(spec, w, h, legend)
    pts = [p for s in sers for p in s["points"]]
    if not pts:
        return _empty(g, w, h, "점이 없습니다 (크기 열은 양수여야 함)")
    xu, yu = core.unit_of(table, spec["x"]), core.unit_of(table, spec["ys"][0])
    zname = spec.get("size") or ""
    zu = core.unit_of(table, zname) if zname else ""
    zmax = max(p[2] for p in pts)
    leg_w = 110 if zname else 0
    rmax = max(8, min(32, 0.17 * (x1 - x0 - leg_w) / math.sqrt(max(1, len(pts)))))  # 점이 많으면 작게
    rr = lambda z: max(3, rmax * math.sqrt(z / zmax))
    xlo, xhi, xt, xs_ = nice_domain(min(p[0] for p in pts), max(p[0] for p in pts), zero=False)
    ylo, yhi, yt, ys_ = nice_domain(min(p[1] for p in pts), max(p[1] for p in pts), zero=False)
    ytitle = axis_title(spec, "y_title", unit_title(spec["ys"][0], yu))
    xtitle = axis_title(spec, "x_title", unit_title(spec["x"], xu))
    room = 20 + (18 if xtitle else 0)
    px0, py0, sy0 = y_axis(g, (x0, y0, x1 - leg_w - 8, y1 - room), ylo, yhi, yt, ys_, yu, ytitle)
    py1 = y1 - room
    px1 = x1 - leg_w - 8
    # 거품이 플롯 밖으로 나가지 않게 안쪽 여백
    ins = rmax * 0.6
    sx = lambda v: px0 + ins + (v - xlo) / ((xhi - xlo) or 1) * (px1 - px0 - 2 * ins)
    sy = lambda v: py1 - ins - (v - ylo) / ((yhi - ylo) or 1) * (py1 - py0 - 2 * ins)
    for t, lab in zip(xt, tick_fmt(xt, xs_, xu)):
        g.line(sx(t), py0, sx(t), py1, GRID, 1)
        g.text(sx(t), py1 + 15, lab, 11.5, MUTED, "middle")
    g.line(px0, py1, px1, py1, BASE, 1)
    if xtitle:
        g.text((px0 + px1) / 2, py1 + 33, xtitle, 12, INK2, "middle")
    labels = spec.get("labels") or "auto"
    allp = sorted(((s, p) for s in sers for p in s["points"]), key=lambda sp: -sp[1][2])
    big = {id(p) for _, p in allp[:10]}
    for s, (x, y, z, lab) in allp:
        c = color_for(spec, s["idx"])
        r = rr(z)
        a, t = tip(lab, s["name"] if multi else "", f"{spec['x']} {core.fmt(x)}", f"{spec['ys'][0]} {core.fmt(y)}", f"{zname} {core.fmt(z)}" if zname else "")
        g.add(f'<circle cx="{sx(x):.1f}" cy="{sy(y):.1f}" r="{r:.1f}" fill="{c}" fill-opacity="0.55" stroke="{c}" stroke-width="1.5" {a}>{t}</circle>')
    placed = []  # 이미 쓴 글자 상자 — 겹치면 건너뜀('모두'일 때만 다 씀)
    for s, p in allp:
        x, y, z, lab = p
        if lab and labels != "none" and (labels == "all" or (id(p) in big and rr(z) >= 7)):
            r = rr(z)
            inside = tw(lab, 11) <= 2 * r - 4
            tx, ty = sx(x), (sy(y) + 4 if inside else sy(y) - r - 4)
            box_ = (tx - tw(lab, 11) / 2 - 2, ty - 11, tx + tw(lab, 11) / 2 + 2, ty + 2)
            if labels != "all" and any(not (box_[2] < b[0] or box_[0] > b[2] or box_[3] < b[1] or box_[1] > b[3]) for b in placed):
                continue
            placed.append(box_)
            g.text(tx, ty, lab, 11, INK, "middle")
    if zname:  # 크기 범례
        lx = x1 - leg_w / 2
        g.text(lx, py0 + 4, clip(unit_title(zname, zu), 11, leg_w), 11, INK2, "middle")
        for k, frac in enumerate((1.0, 0.25)):
            z = zmax * frac
            r = rr(z)
            cyb = py0 + 20 + 2 * rmax
            g.add(f'<circle cx="{lx:.1f}" cy="{cyb - r:.1f}" r="{r:.1f}" fill="none" stroke="{MUTED}" stroke-width="1"/>')
            g.text(lx, cyb - 2 * r - 3 if k == 0 else cyb + 13, core.fmt(z), 10.5, MUTED, "middle")
    return g.out()


# ── 계단선·막대사탕·덤벨·불릿·누적+선 (범주 x) ─────────────────────────────────
def render_step(table, spec, w, h):
    d = core.category_data(table, spec)
    cats, series = d["cats"], _vis(d)
    legend = [(s["name"], color_for(spec, s["idx"]), "line") for s in series]
    g, box = frame(spec, w, h, legend)
    vals = [v for s in series for v in s["values"] if v is not None]
    if not cats or not vals:
        return _empty(g, w, h)
    unit = series[0].get("unit", "")
    lo, hi = min(vals), max(vals)
    lo, hi, ticks, step = nice_domain(lo, hi, lo >= 0 and lo < 0.35 * hi)
    px0, py0, py1, px1, band, cx, sy = _cat_axes(g, box, cats, lo, hi, ticks, step, unit, _ytitle(spec, series, unit),
                                                 axis_title(spec, "x_title", spec.get("x") or ""))
    dec = core.decimals_of(vals)
    labels = spec.get("labels") or "auto"
    n = len(cats)
    for s in series:
        c = color_for(spec, s["idx"])
        pv = [(i, v) for i, v in enumerate(s["values"]) if v is not None]
        if not pv:
            continue
        dpath = ""
        for k, (i, v) in enumerate(pv):
            xa = cx(i)
            xb = cx(pv[k + 1][0]) if k + 1 < len(pv) else min(px1, cx(i) + band / 2)
            dpath += (f"M{xa:.1f},{sy(v):.1f}" if k == 0 else f"V{sy(v):.1f}") + f"H{xb:.1f}"
        g.add(f'<path d="{dpath}" fill="none" stroke="{c}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>')
        prev = None
        for i, v in pv:
            a, t = tip(cats[i], s["name"] if len(series) > 1 else "", core.fmt(v, dec, unit) + _sfx(unit))
            changed = prev is None or v != prev
            ring = f'stroke="{SURFACE}" stroke-width="2"' if changed else ""
            fill = c if changed else "transparent"
            g.add(f'<circle cx="{cx(i):.1f}" cy="{sy(v):.1f}" r="{4 if changed else 7}" fill="{fill}" {ring} {a}>{t}</circle>')
            if changed and labels != "none" and (labels == "all" or (len(series) == 1 and n <= 40)):
                g.text(cx(i) + 3, sy(v) - 8, core.fmt(v, dec, unit), 10.5, INK, "start")
            prev = v
    return g.out()


def render_lollipop(table, spec, w, h):
    d = core.category_data(table, spec)
    cats, series = d["cats"], _vis(d)[:1]
    g, box = frame(spec, w, h, [])
    vals = [v for s in series for v in s["values"] if v is not None]
    if not cats or not vals:
        return _empty(g, w, h)
    s = series[0]
    unit = s.get("unit", "")
    lo, hi, ticks, step = nice_domain(min(vals), max(vals), True)
    px0, py0, py1, px1, band, cx, sy = _cat_axes(g, box, cats, lo, hi, ticks, step, unit, _ytitle(spec, series, unit),
                                                 axis_title(spec, "x_title", spec.get("x") or ""))
    c = color_for(spec, s["idx"])
    dec = core.decimals_of(vals)
    labels = spec.get("labels") or "auto"
    r = max(3.5, min(6.5, band * 0.28))
    for i, v in enumerate(s["values"]):
        if v is None:
            continue
        g.line(cx(i), sy(0), cx(i), sy(v), c, 2, 'stroke-linecap="round" opacity="0.75"')
        a, t = tip(cats[i], core.fmt(v, dec, unit) + _sfx(unit))
        g.add(f'<circle cx="{cx(i):.1f}" cy="{sy(v):.1f}" r="{r:.1f}" fill="{c}" stroke="{SURFACE}" stroke-width="2" {a}>{t}</circle>')
        lab = core.fmt(v, dec, unit)
        if labels != "none" and (labels == "all" or (len(cats) <= 20 and band >= tw(lab, 10.5) - 2)):
            g.text(cx(i), sy(v) - r - 5 if v >= 0 else sy(v) + r + 13, lab, 10.5, INK, "middle")
    return g.out()


def render_dumbbell(table, spec, w, h):
    d = core.category_data(table, spec)
    cats, series = d["cats"], _vis(d)[:2]
    legend = [(s["name"], color_for(spec, s["idx"]), "dot") for s in series]
    g, box = frame(spec, w, h, legend)
    vals = [v for s in series for v in s["values"] if v is not None]
    if not cats or len(series) < 2 or not vals:
        return _empty(g, w, h, "덤벨은 비교할 계열 2개가 필요합니다")
    unit = series[0].get("unit", "")
    lo, hi = min(vals), max(vals)
    lo, hi, ticks, step = nice_domain(lo, hi, lo >= 0 and lo < 0.35 * hi)
    px0, py0, py1, px1, band, cx, sy = _cat_axes(g, box, cats, lo, hi, ticks, step, unit, _ytitle(spec, [], unit),
                                                 axis_title(spec, "x_title", spec.get("x") or ""))
    a_s, b_s = series
    dec = core.decimals_of(vals)
    labels = spec.get("labels") or "auto"
    ca, cb = color_for(spec, a_s["idx"]), color_for(spec, b_s["idx"])
    for i in range(len(cats)):
        va, vb = a_s["values"][i], b_s["values"][i]
        if va is not None and vb is not None:
            g.line(cx(i), sy(va), cx(i), sy(vb), BASE, 3.5, 'stroke-linecap="round"')
        for s, v, col in ((a_s, va, ca), (b_s, vb, cb)):
            if v is None:
                continue
            diff = "" if va is None or vb is None else f"차이 {'+' if vb - va >= 0 else ''}{core.fmt(vb - va, dec, unit)}"
            at, t = tip(cats[i], s["name"], core.fmt(v, dec, unit) + _sfx(unit), diff)
            g.add(f'<circle cx="{cx(i):.1f}" cy="{sy(v):.1f}" r="6" fill="{col}" stroke="{SURFACE}" stroke-width="2" {at}>{t}</circle>')
        if labels != "none" and va is not None and vb is not None and (labels == "all" or len(cats) <= 12):
            upper, lower = (va, vb) if va >= vb else (vb, va)
            g.text(cx(i), sy(upper) - 10, core.fmt(upper, dec, unit), 10.5, INK, "middle")
            g.text(cx(i), sy(lower) + 18, core.fmt(lower, dec, unit), 10.5, INK2, "middle")
    return g.out()


def render_bullet(table, spec, w, h):
    d = core.category_data(table, spec)
    cats = d["cats"]
    act = _vis(d)[:1]
    tgt = next((s for s in d["series"] if s.get("role") == "target" and not s["hidden"]), None)
    tcol = INK2
    legend = [(s["name"], color_for(spec, s["idx"]), "bar") for s in act] + ([(tgt["name"], tcol, "tick")] if tgt else [])
    g, box = frame(spec, w, h, legend)
    vals = [v for s in act + ([tgt] if tgt else []) for v in s["values"] if v is not None]
    if not cats or not act or not vals:
        return _empty(g, w, h)
    s = act[0]
    unit = s.get("unit", "")
    lo, hi, ticks, step = nice_domain(min(vals), max(vals), True)
    px0, py0, py1, px1, band, cx, sy = _cat_axes(g, box, cats, lo, hi, ticks, step, unit, _ytitle(spec, act, unit),
                                                 axis_title(spec, "x_title", spec.get("x") or ""))
    c = color_for(spec, s["idx"])
    bw = min(24, band * 0.42)
    dec = core.decimals_of(vals)
    labels = spec.get("labels") or "auto"
    for i in range(len(cats)):
        v = s["values"][i]
        tv = tgt["values"][i] if tgt else None
        if tv is not None:  # 목표까지 연한 바탕
            g.add(f'<rect x="{cx(i) - bw * 0.95:.1f}" y="{sy(max(tv, 0)):.1f}" width="{bw * 1.9:.1f}" height="{abs(sy(0) - sy(tv)):.1f}" rx="2" fill="#f0efec"/>')
        if v is not None:
            ya, yb = sy(max(v, 0)), sy(min(v, 0))
            a, t = tip(cats[i], f"{s['name']} {core.fmt(v, dec, unit)}{_sfx(unit)}",
                       f"{tgt['name']} {core.fmt(tv, dec, unit)}{_sfx(unit)}" if tv is not None else "",
                       f"달성률 {v / tv * 100:.1f}%" if tv else "")
            g.add(f'<path d="{bar_path(cx(i) - bw / 2, ya, bw, max(yb - ya, .5), 3, "top" if v >= 0 else "bottom")}" fill="{c}" {a}>{t}</path>')
        if tv is not None:
            g.line(cx(i) - bw * 0.95, sy(tv), cx(i) + bw * 0.95, sy(tv), tcol, 3.5)
        if labels != "none" and v is not None and tv:
            pct = f"{v / tv * 100:.0f}%"
            top = min(sy(v), sy(tv))
            g.text(cx(i), top - 7, pct if labels == "auto" else f"{core.fmt(v, dec, unit)} ({pct})", 11, INK, "middle", 600 if v >= tv else 400)
    return g.out()


def render_stack_line(table, spec, w, h):
    d = core.category_data(table, spec)
    cats = d["cats"]
    bars = _vis(d)
    line = next((s for s in d["series"] if s.get("role") == "line" and not s["hidden"]), None)
    legend = [(s["name"], color_for(spec, s["idx"]), "bar") for s in bars] + ([(line["name"], color_for(spec, line["idx"]), "line")] if line else [])
    g, box = frame(spec, w, h, legend)
    if not cats or not (bars or line):
        return _empty(g, w, h)
    n = len(cats)
    pos = [sum(max(0, s["values"][i] or 0) for s in bars) for i in range(n)]
    neg = [sum(min(0, s["values"][i] or 0) for s in bars) for i in range(n)]
    lv = [v for v in (line["values"] if line else []) if v is not None]
    lo, hi, ticks, step = nice_domain(min(neg + lv + [0]), max(pos + lv + [0]), True)
    unit = (bars or [line])[0].get("unit", "")
    px0, py0, py1, px1, band, cx, sy = _cat_axes(g, box, cats, lo, hi, ticks, step, unit, _ytitle(spec, [], unit),
                                                 axis_title(spec, "x_title", spec.get("x") or ""))
    dec = core.decimals_of([v for s in bars for v in s["values"] if v is not None] + lv)
    f = lambda v: core.fmt(v, dec, unit)
    labels = spec.get("labels") or "auto"
    bw = min(28, band * 0.6)
    for i in range(n):
        up = 0.0
        segs = [(s, s["values"][i]) for s in bars if s["values"][i]]
        segs = [(s, v) for s, v in segs if v > 0]
        for j, (s, v) in enumerate(segs):
            c = color_for(spec, s["idx"])
            last = j == len(segs) - 1
            ya, yb = sy(up + v), sy(up)
            gap = 0 if last else 2
            a, t = tip(cats[i], s["name"], f(v) + _sfx(unit))
            g.add(f'<path d="{bar_path(cx(i) - bw / 2, ya + gap, bw, max(yb - ya - gap, .5), 4 if last else 0, "top" if last else "x")}" fill="{c}" {a}>{t}</path>')
            if labels == "all" and yb - ya >= 16 and bw >= tw(f(v), 10.5) + 4:
                g.text(cx(i), (ya + yb) / 2 + 4, f(v), 10.5, text_on(c), "middle", extra=on_bg(c))
            up += v
        if labels in ("auto", "all") and up > 0 and n <= 16 and not line:
            g.text(cx(i), sy(up) - 5, f(up), 11, INK, "middle")
    if line:
        c = color_for(spec, line["idx"])
        pts = [(cx(i), sy(v), v, i) for i, v in enumerate(line["values"]) if v is not None]
        g.add(f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y, _, _ in pts)}" fill="none" stroke="{c}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>')
        for x, y, v, i in pts:
            a, t = tip(cats[i], line["name"], f(v) + _sfx(unit), f"막대 합계 {f(pos[i])}")
            g.add(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="{c}" stroke="{SURFACE}" stroke-width="2" {a}>{t}</circle>')
        if pts and labels != "none":
            x, y, v, _ = pts[-1]
            g.text(x + 8, y + 4, f(v), 11.5, INK, "start") if x + 8 + tw(f(v), 11.5) < w else g.text(x, y - 10, f(v), 11.5, INK, "middle")
    return g.out()


# ── 트리맵·로즈 ─────────────────────────────────────────────────────────────
def render_treemap(table, spec, w, h):
    d = core.category_data(table, spec)
    sers = _vis(d)
    g, (x0, y0, x1, y1) = frame(spec, w, h, [])
    if not sers:
        return _empty(g, w, h)
    s = sers[0]
    hidden = set(spec.get("hidden_cats") or [])
    items = sorted(((c, v) for c, v in zip(d["cats"], s["values"]) if v and v > 0 and c not in hidden), key=lambda cv: -cv[1])
    if not items:
        return _empty(g, w, h, "양수 값이 없습니다")
    unit = s.get("unit", "")
    tot = sum(v for _, v in items)
    dec = core.decimals_of([v for _, v in items])
    lo, hi = min(v for _, v in items), max(v for _, v in items)
    labels = spec.get("labels") or "auto"
    base = spec.get("colors")
    custom = (base[0] if isinstance(base, list) and base and base[0] else None) if isinstance(base, list) else (base or {}).get("0")
    for k, v, x, y, ww, hh in core.squarify(items, x0, y0 + 2, x1 - x0, y1 - y0 - 2):
        f = (v - lo) / ((hi - lo) or 1)
        col = custom or SEQ[min(len(SEQ) - 1, 3 + int(round(f * (len(SEQ) - 4))))]
        a, t = tip(k, core.fmt(v, dec, unit) + _sfx(unit) + ("" if unit != "%" else ""), f"{v / tot * 100:.1f}%")
        g.add(f'<rect x="{x + 1:.1f}" y="{y + 1:.1f}" width="{max(ww - 2, .5):.1f}" height="{max(hh - 2, .5):.1f}" rx="3" fill="{col}" {a}>{t}</rect>')
        if labels == "none":
            continue
        name = clip(k, 12.5, ww - 12)
        if ww >= 34 and hh >= 22 and name != "…":
            g.text(x + 8, y + 18, name, 12.5, text_on(col), "start", 600, on_bg(col))
            vt = f"{core.fmt(v, dec, unit)}{'%' if unit == '%' else ''} · {v / tot * 100:.1f}%"
            if hh >= 40 and tw(vt, 11.5) <= ww - 14:
                g.text(x + 8, y + 34, vt, 11.5, text_on(col), "start", extra=on_bg(col))
    return g.out()


def render_rose(table, spec, w, h):
    d = core.category_data(table, spec)
    sers = _vis(d)
    g, (x0, y0, x1, y1) = frame(spec, w, h, [])
    if not sers:
        return _empty(g, w, h)
    s = sers[0]
    cats, vals = d["cats"], s["values"]
    vs = [v for v in vals if v is not None and v >= 0]
    if not vs or max(vs) <= 0:
        return _empty(g, w, h)
    unit = s.get("unit", "")
    n = len(cats)
    labw = min(70, max(tw(c, 12) for c in cats) + 8)
    R = max(40, min((x1 - x0) / 2 - labw - 10, (y1 - y0) / 2 - 22))
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2 + 2
    vmax = max(vs)
    _, top, ticks, step = nice_domain(0, vmax, True, 4)
    rad = lambda v: R * math.sqrt(max(v, 0) / top)  # 넓이 ∝ 값
    for t, lab in zip(ticks, tick_fmt(ticks, step, unit)):
        if t <= 0:
            continue
        g.add(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{rad(t):.1f}" fill="none" stroke="{GRID}" stroke-width="1"/>')
        g.text(cx + 3, cy - rad(t) + 11, lab, 10.5, MUTED, "start")
    col = color_for(spec, s["idx"])
    dec = core.decimals_of(vs)
    labels = spec.get("labels") or "auto"
    da = 2 * math.pi / n
    for i, (c, v) in enumerate(zip(cats, vals)):
        a1 = -math.pi / 2 + da * i
        a2 = a1 + da
        p = lambda a, rr: (cx + rr * math.cos(a), cy + rr * math.sin(a))
        lx, ly = p((a1 + a2) / 2, R + 12)
        ca = math.cos((a1 + a2) / 2)
        g.text(lx, ly + 4, clip(c, 12, labw), 12, INK2, "start" if ca > 0.25 else ("end" if ca < -0.25 else "middle"))
        if v is None or v <= 0:
            continue
        r = rad(v)
        (ax, ay), (bx, by) = p(a1, r), p(a2, r)
        at, t = tip(c, core.fmt(v, dec, unit) + _sfx(unit))
        g.add(f'<path d="M{cx:.1f},{cy:.1f}L{ax:.1f},{ay:.1f}A{r:.1f},{r:.1f} 0 0,1 {bx:.1f},{by:.1f}Z" fill="{col}" fill-opacity="0.85" stroke="{SURFACE}" stroke-width="2" stroke-linejoin="round" {at}>{t}</path>')
        if labels != "none" and (labels == "all" or v >= vmax * 0.25):
            mx, my = p((a1 + a2) / 2, r - 12 if r > 34 else r + 10)
            g.text(mx, my + 4, core.fmt(v, dec, unit), 10.5, text_on(col) if r > 34 else INK, "middle", extra=on_bg(col) if r > 34 else "")
    return g.out()


def render(table, spec, w=720, h=440):
    k = spec.get("kind")
    try:
        if k in ("pie", "doughnut", "semi_doughnut"):
            return render_pie(table, spec, w, h)
        fn = {"radar": render_radar, "sparkline": render_sparkline, "candlestick": render_candle, "ohlc": render_candle, "waterfall": render_waterfall,
              "funnel": render_funnel, "bubble": render_bubble, "step": render_step, "lollipop": render_lollipop,
              "dumbbell": render_dumbbell, "bullet": render_bullet, "stack_line": render_stack_line,
              "treemap": render_treemap, "rose": render_rose}.get(k)
        if fn:
            return fn(table, spec, w, h)
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
