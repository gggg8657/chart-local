"""chart local — 선택한 그래프를 받은 쪽에서 고칠 수 있는 문서로 내보내기

XLSX : '데이터' 시트(원본) + 그림마다 시트 하나(차트용 표 = 데이터 시트를 가리키는 수식) + 엑셀 네이티브 차트
PPTX : python-pptx 네이티브 차트 (내장 워크북 → PowerPoint '데이터 편집')
DOCX : 같은 python-pptx 차트 파트(chartSpace XML + 내장 xlsx)를 word/charts 로 옮겨 심은 네이티브 차트 (Word '데이터 편집')
HWPX : kordoc 의 ```chart 펜스 → 한글 네이티브 차트(Chart/chartN.xml), 데이터 표는 한글 표로
히트맵은 어느 형식이든 '색 칠한 표'(XLSX 는 조건부 서식 색조) — 차트 개체가 아니라 편집 가능한 표.
반환되는 report 에 그림마다 무엇으로 들어갔는지(native|table|image)를 정직하게 적는다.
"""
import copy
import io
import os
import re
import shutil
import subprocess
import tempfile

import core
import svg

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = "맑은 고딕"
PAL = [c.lstrip("#").upper() for c in svg.PAL]


def hexcol(spec, idx):
    return svg.color_for(spec, idx).lstrip("#").upper()


def numfmt(vals, unit=""):
    d = core.decimals_of([v for v in vals if v is not None])
    f = "#,##0" + ("." + "0" * d if d else "")
    return f + '"%"' if unit == "%" else f


# ── 새 종류: 블록 ────────────────────────────────────────────────────────────
# 오피스 차트에 해당 형식이 없는 종류 → 그림(PNG) + 데이터 표 (rsvg-convert 가 없으면 표만)
IMAGE_KINDS = {"step", "treemap", "rose"}
# 한글(kordoc) 차트가 없는 종류 → 그림 + 표 (radar 만 한글 네이티브)
HWP_IMAGE_KINDS = {"sparkline", "candlestick", "waterfall", "funnel", "bubble", "step", "lollipop", "dumbbell", "treemap",
                   "bullet", "rose", "stack_line", "semi_doughnut"}
NEW_CT = {"radar", "stock", "waterfall", "funnel", "bubble", "lollipop", "dumbbell", "bullet", "stack_line", "semi_doughnut"}
KIND_KO = {"up": "증가", "down": "감소", "total": "합계"}


def image_block(table, spec, base):
    """그림으로 넣는 블록 — 데이터 표용 범주·계열(역할 계열 포함)을 함께 든다"""
    d = core.category_data(table, spec)
    sers = [s for s in d["series"] if not s["hidden"]]
    return dict(base, mode="image", cats=d["cats"], series=sers, unit=sers[0]["unit"] if sers else "",
                x_title=spec.get("x_title", spec.get("x") or ""), note=d["note"])


def wf_parts(st):
    """폭포 막대 하나 → (기준, 본 막대, 0 아래 조각). 0 을 가로지르는 막대만 0 아래 조각이 생긴다(누적 막대는 음수를 아래로 따로 쌓음)"""
    lo, hi = min(st["a"], st["b"]), max(st["a"], st["b"])
    if lo >= 0:
        return lo, hi - lo, 0.0
    if hi <= 0:
        return hi, lo - hi, 0.0
    return 0.0, hi, lo


def wf_series(steps, spec):
    """폭포 → 누적 막대 계열: 기준(투명) + 증가·감소·합계 (+ 0 을 가로지를 때만 '(0 아래)' 계열)"""
    cols = [c.lstrip("#").upper() for c in svg.wf_colors(spec)]
    parts = [wf_parts(st) for st in steps]
    out = [{"name": "기준", "values": [p[0] for p in parts], "role": "base", "color": None}]
    cross = any(p[2] for p in parts)
    for k, col in zip(("up", "down", "total"), cols):
        out.append({"name": KIND_KO[k], "values": [p[1] if st["kind"] == k else 0.0 for st, p in zip(steps, parts)], "color": col, "kind": k})
    if cross:
        for k, col in zip(("up", "down", "total"), cols):
            out.append({"name": KIND_KO[k] + " (0 아래)", "values": [p[2] if st["kind"] == k else 0.0 for st, p in zip(steps, parts)],
                        "color": col, "kind": k, "role": "extra"})
    return out


def new_blocks(table, spec, base, title):
    """새 종류의 블록. 해당 없으면 None"""
    k = spec["kind"]
    xt = spec.get("x_title", spec.get("x") or "")
    if k in IMAGE_KINDS:
        return [image_block(table, spec, base)]
    if k == "sparkline":
        d = core.category_data(table, spec)
        sers = [s for s in d["series"] if not s["hidden"]]
        return [dict(base, mode="spark", cats=d["cats"], series=sers, unit="", x_title=xt, note=d["note"])]
    if k == "candlestick":
        d = core.category_data(table, spec)
        px = [s for s in d["series"] if not s.get("role")][:4]
        if len(px) < 4:
            raise ValueError("캔들스틱은 시가·고가·저가·종가 4개 열이 필요합니다")
        vol = next((s for s in d["series"] if s.get("role") == "volume" and not s["hidden"]), None)
        unit = px[3]["unit"]
        up, dn = svg.candle_colors(spec)
        out = [dict(base, mode="chart", ctype="stock", cats=d["cats"], series=px, unit=unit, x_title=xt, up=up, down=dn,
                    y_title=spec.get("y_title") or svg.unit_title("가격", unit), note=d["note"])]
        if vol:
            out[0]["panel"] = 1
            vspec = dict(spec, colors={str(vol["idx"]): "#9ec5f4"})
            out.append(dict(base, spec=vspec, mode="chart", ctype="col", cats=d["cats"], series=[vol], unit=vol["unit"], x_title=xt,
                            y_title=svg.unit_title(vol["name"], vol["unit"]), title=f"{title} — {vol['name']}", panel=2, labels="none"))
        return out
    if k == "waterfall":
        d = core.category_data(table, spec)
        steps = core.waterfall_steps(table, spec, d)
        src = next((s for s in d["series"] if not s["hidden"] and not s.get("role")), None)
        if not steps or not src:
            raise ValueError("폭포에 넣을 값이 없습니다")
        return [dict(base, mode="chart", ctype="waterfall", cats=[st["cat"] for st in steps], steps=steps, src=src["name"],
                     series=wf_series(steps, spec), unit=src["unit"], x_title=xt,
                     y_title=spec.get("y_title") or svg.unit_title(src["name"], src["unit"]), note=d["note"])]
    if k == "funnel":
        d = core.category_data(table, spec)
        s = next((s for s in d["series"] if not s["hidden"] and not s.get("role")), None)
        items = [(c, v) for c, v in zip(d["cats"], s["values"] if s else []) if v is not None and v > 0]
        if not items:
            raise ValueError("깔때기에 넣을 양수 값이 없습니다")
        mx = max(v for _, v in items)
        return [dict(base, mode="chart", ctype="funnel", cats=[c for c, _ in items], src=s["name"], unit=s["unit"], x_title=xt,
                     series=[{"name": "여백", "values": [(mx - v) / 2 for _, v in items], "role": "base", "color": None},
                             {"name": s["name"], "values": [v for _, v in items], "idx": s["idx"]}], note=d["note"])]
    if k == "bubble":
        sers = [s for s in core.bubble_data(table, spec) if not s["hidden"] and s["points"]]
        if not sers:
            raise ValueError("거품에 넣을 점이 없습니다 (크기 열은 양수여야 함)")
        return [dict(base, mode="chart", ctype="bubble", series=sers, size=spec.get("size") or "",
                     x_title=spec.get("x_title") or svg.unit_title(spec["x"], core.unit_of(table, spec["x"])),
                     y_title=spec.get("y_title") or svg.unit_title(spec["ys"][0], core.unit_of(table, spec["ys"][0])))]
    if k == "semi_doughnut":
        d = core.category_data(table, spec)
        s = next((s for s in d["series"] if not s["hidden"] and not s.get("role")), None)
        items = [(c, v) for c, v in zip(d["cats"], s["values"] if s else []) if v and v > 0]
        hidden = set(spec.get("hidden_cats") or [])
        items = [(c, v) for c, v in items if c not in hidden]
        if not items:
            raise ValueError("반원 도넛에 넣을 양수 값이 없습니다")
        return [dict(base, mode="chart", ctype="semi_doughnut", cats=[c for c, _ in items],
                     series=[{"name": s["name"], "values": [v for _, v in items], "idx": 0}], unit=s["unit"])]
    return None


# ── 새 종류: PPTX(·DOCX) 네이티브 차트 ─────────────────────────────────────────
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS2 = f'xmlns:c="{C_NS}" xmlns:a="{A_NS}"'


def _x(s):
    from lxml import etree
    return etree.fromstring(s)


def _ln(hex6, w=19050, alpha=None):
    al = f'<a:alpha val="{alpha}"/>' if alpha else ""
    return f'<a:ln w="{w}"><a:solidFill><a:srgbClr val="{hex6}">{al}</a:srgbClr></a:solidFill></a:ln>'


def _set_sppr(ser_el, inner):
    """계열 c:spPr 를 통째로 바꿈(c:tx 바로 뒤)"""
    from pptx.oxml.ns import qn
    old = ser_el.find(qn("c:spPr"))
    new = _x(f'<c:spPr {NS2}>{inner}</c:spPr>')
    if old is not None:
        ser_el.replace(old, new)
    else:
        (ser_el.find(qn("c:tx")) if ser_el.find(qn("c:tx")) is not None else ser_el.find(qn("c:order"))).addnext(new)


def _set_marker(ser_el, symbol, size=7, fill=None, line="FFFFFF"):
    from pptx.oxml.ns import qn
    old = ser_el.find(qn("c:marker"))
    body = f'<c:symbol val="{symbol}"/>' + (f'<c:size val="{size}"/>' if symbol != "none" else "")
    if symbol != "none" and fill:
        body += f'<c:spPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{_ln(line, 9525) if line else "<a:ln><a:noFill/></a:ln>"}</c:spPr>'
    new = _x(f'<c:marker {NS2}>{body}</c:marker>')
    if old is not None:
        ser_el.replace(old, new)
    else:
        ser_el.find(qn("c:spPr")).addnext(new)


def _legend_hide(chart_space, idxs):
    """범례에서 이 계열(또는 조각) 번호만 지움 — 투명 기준 계열·숨김 조각"""
    from pptx.oxml.ns import qn
    leg = chart_space.find(qn("c:chart")).find(qn("c:legend"))
    if leg is None or not idxs:
        return
    pos = leg.find(qn("c:legendPos"))
    for i in sorted(idxs, reverse=True):
        e = _x(f'<c:legendEntry {NS2}><c:idx val="{i}"/><c:delete val="1"/></c:legendEntry>')
        if pos is not None:
            pos.addnext(e)
        else:
            leg.insert(0, e)


def _split_to_line(plot_el, idxs, styles):
    """막대 그룹에서 계열 몇 개를 빼 같은 축을 쓰는 c:lineChart 로 옮김(불릿 목표·누적+선)"""
    from lxml import etree
    from pptx.oxml.ns import qn
    sers = plot_el.findall(qn("c:ser"))
    axids = [e.get("val") for e in plot_el.findall(qn("c:axId"))]
    xml = []
    for i, st in zip(idxs, styles):
        sr = sers[i]
        parts = [etree.tostring(sr.find(qn(t))).decode() for t in ("c:idx", "c:order", "c:tx")]
        cat, val = sr.find(qn("c:cat")), sr.find(qn("c:val"))
        xml.append(f'<c:ser {NS2}>' + "".join(parts) + st + etree.tostring(cat).decode() + etree.tostring(val).decode() + '<c:smooth val="0"/></c:ser>')
        plot_el.remove(sr)
    line = _x(f'<c:lineChart {NS2}><c:grouping val="standard"/><c:varyColors val="0"/>' + "".join(xml) +
              '<c:marker val="1"/>' + "".join(f'<c:axId val="{a}"/>' for a in axids) + '</c:lineChart>')
    plot_el.addnext(line)
    return line


def _pptx_style(ch, block, legend):
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.util import Pt
    ch.font.size = Pt(10)
    ch.font.name = FONT
    ch.font.color.rgb = RGBColor.from_string("52514E")
    ch.has_title = True
    ch.chart_title.text_frame.text = block["title"]
    tp = ch.chart_title.text_frame.paragraphs[0]
    tp.font.size, tp.font.bold, tp.font.name = Pt(13), True, FONT
    tp.font.color.rgb = RGBColor.from_string("0B0B0B")
    on = legend and block["legend"] != "none"
    ch.has_legend = on
    if on:
        ch.legend.position = {"top": XL_LEGEND_POSITION.TOP, "bottom": XL_LEGEND_POSITION.BOTTOM,
                              "right": XL_LEGEND_POSITION.RIGHT}.get(block["legend"], XL_LEGEND_POSITION.TOP)
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(10)


def _pptx_axes(ch, block, numf, titles=True, cat_grid=False):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor.from_string("E1E0D9")
    va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(9)
    va.tick_labels.font.color.rgb = RGBColor.from_string("898781")
    if numf:
        va.tick_labels.number_format = numf
        va.tick_labels.number_format_is_linked = False
    ca = ch.category_axis
    ca.format.line.color.rgb = RGBColor.from_string("C3C2B7")
    ca.tick_labels.font.size = Pt(9)
    ca.tick_labels.font.color.rgb = RGBColor.from_string("52514E")
    ca.has_major_gridlines = cat_grid
    if cat_grid:
        ca.major_gridlines.format.line.color.rgb = RGBColor.from_string("E1E0D9")
    if titles:
        for ax, txt in ((va, block.get("y_title")), (ca, block.get("x_title"))):
            if txt:
                ax.has_title = True
                ax.axis_title.text_frame.text = txt
                f = ax.axis_title.text_frame.paragraphs[0].font
                f.size, f.bold, f.name = Pt(9.5), False, FONT
                f.color.rgb = RGBColor.from_string("52514E")
    return va, ca


def _pptx_labels(plot, numf, pos=None, color="0B0B0B", size=9):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(size)
    dl.font.color.rgb = RGBColor.from_string(color)
    dl.number_format = numf
    dl.number_format_is_linked = False
    dl.show_value = True
    if pos is not None:
        dl.position = pos
    return dl


def _ser_labels(ser, numf, pos=None, color="0B0B0B", bold=False):
    """계열 하나에만 값 라벨"""
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    dl = ser.data_labels
    dl.show_value = True
    dl.number_format = numf
    dl.number_format_is_linked = False
    dl.font.size = Pt(9)
    dl.font.bold = bold
    dl.font.color.rgb = RGBColor.from_string(color)
    if pos is not None:
        dl.position = pos
    return dl


def _pptx_chart_new(slide_shapes, block, x, y, cx, cy):
    from pptx.chart.data import BubbleChartData, CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE as T, XL_LABEL_POSITION as LP, XL_MARKER_STYLE
    from pptx.oxml.ns import qn
    from pptx.util import Pt

    ct = block["ctype"]
    spec = block["spec"]
    unit = block.get("unit", "")
    lab = block["labels"]
    sers = block.get("series") or []
    allv = [v for s in sers for v in (s.get("values") or []) if v is not None]
    nf = numfmt(allv, unit)
    if ct == "bubble":
        cd = BubbleChartData()
        for s in sers:
            se = cd.add_series(s["name"])
            for px_, py_, pz, _lab in s["points"]:
                se.add_data_point(px_, py_, pz)
        gf = slide_shapes.add_chart(T.BUBBLE, x, y, cx, cy, cd)
    else:
        cd = CategoryChartData(number_format=nf)
        cd.categories = block["cats"] + (["(숨김) 합계"] if ct == "semi_doughnut" else [])
        for s in sers:
            vals = list(s["values"]) + ([sum(s["values"])] if ct == "semi_doughnut" else [])
            cd.add_series(s["name"], vals, number_format=numfmt(s["values"], s.get("unit", unit)))
        typ = {"radar": T.RADAR_MARKERS, "stock": T.LINE, "waterfall": T.COLUMN_STACKED, "funnel": T.BAR_STACKED,
               "lollipop": T.LINE_MARKERS, "dumbbell": T.LINE_MARKERS, "bullet": T.COLUMN_CLUSTERED,
               "stack_line": T.COLUMN_STACKED, "semi_doughnut": T.DOUGHNUT}[ct]
        gf = slide_shapes.add_chart(typ, x, y, cx, cy, cd)
    ch = gf.chart
    cs = ch._chartSpace
    plot = ch.plots[0]
    pel = plot._element
    col_of = lambda s, i: (s.get("color") or hexcol(spec, s.get("idx", i)))
    want_leg = {"radar": len(sers) >= 2, "stock": True, "waterfall": True, "funnel": False, "bubble": len(sers) >= 2,
                "lollipop": False, "dumbbell": True, "bullet": True, "stack_line": True, "semi_doughnut": True}[ct]
    _pptx_style(ch, block, want_leg)
    if ct == "semi_doughnut":
        from pptx.enum.chart import XL_LEGEND_POSITION
        plot.vary_by_categories = True
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.RIGHT
        ser = plot.series[0]
        n = len(block["cats"])
        for i in range(n + 1):
            pt = ser.points[i]
            if i < n:
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = RGBColor.from_string(hexcol(spec, i))
                pt.format.line.color.rgb = RGBColor.from_string("FFFFFF")
                pt.format.line.width = Pt(1.5)
            else:  # 아래쪽 반 = 합계와 같은 투명 조각
                pt.format.fill.background()
                pt.format.line.fill.background()
        if lab != "none":
            dl = _pptx_labels(plot, nf, None, "FFFFFF", 10)
            dl.font.bold = True
            dl.show_percentage = False
            dle = dl._element
            dle.insert(0, _x(f'<c:dLbl {NS2}><c:idx val="{n}"/><c:delete val="1"/></c:dLbl>'))
        pel.find(qn("c:firstSliceAng")).set("val", "270")
        hs = pel.find(qn("c:holeSize"))
        if hs is not None:
            hs.set("val", "50")
        _legend_hide(cs, [n])
        return gf
    if ct == "bubble":
        va, ca = _pptx_axes(ch, block, numfmt([p[1] for s in sers for p in s["points"]]), cat_grid=True)
        ca.tick_labels.number_format = numfmt([p[0] for s in sers for p in s["points"]])
        ca.tick_labels.number_format_is_linked = False
        bs = pel.find(qn("c:bubbleScale"))
        if bs is not None:
            bs.set("val", "60")
        for i, (s, ser) in enumerate(zip(sers, plot.series)):
            c = col_of(s, i)
            _set_sppr(ser._element, f'<a:solidFill><a:srgbClr val="{c}"><a:alpha val="60000"/></a:srgbClr></a:solidFill>{_ln(c, 12700)}')
            if lab != "none":  # 점 이름(큰 것부터 10개, '모두'면 전부)
                order = sorted(range(len(s["points"])), key=lambda j: -s["points"][j][2])
                pick = set(order if lab == "all" else order[:10])
                for j, p in enumerate(s["points"]):
                    if j in pick and p[3]:
                        tf = ser.points[j].data_label.text_frame
                        tf.text = str(p[3])
                        tf.paragraphs[0].font.size = Pt(8.5)
                        tf.paragraphs[0].font.color.rgb = RGBColor.from_string("0B0B0B")
                        ser.points[j].data_label.position = LP.CENTER
        return gf
    if ct == "radar":
        va, ca = _pptx_axes(ch, block, nf, titles=False)
        for i, (s, ser) in enumerate(zip(sers, plot.series)):
            c = RGBColor.from_string(col_of(s, i))
            ser.format.line.color.rgb = c
            ser.format.line.width = Pt(2.25)
            ser.marker.style = XL_MARKER_STYLE.CIRCLE
            ser.marker.size = 6
            ser.marker.format.fill.solid()
            ser.marker.format.fill.fore_color.rgb = c
            ser.marker.format.line.color.rgb = RGBColor.from_string("FFFFFF")
        if lab == "all" or (lab == "auto" and len(sers) == 1 and len(block["cats"]) <= 10):
            _pptx_labels(plot, nf)
        return gf
    va, ca = _pptx_axes(ch, block, nf if ct != "funnel" else None, titles=ct != "funnel")
    if ct == "stock":
        up, dn = block["up"].lstrip("#").upper(), block["down"].lstrip("#").upper()
        lo = min(v for s in sers for v in s["values"] if v is not None)
        hi = max(v for s in sers for v in s["values"] if v is not None)
        a_, b_, _, st_ = svg.nice_domain(lo, hi, zero=False)
        va.minimum_scale, va.maximum_scale = a_, b_
        pel.tag = qn("c:stockChart")
        for t in ("c:grouping", "c:varyColors", "c:marker", "c:dLbls"):
            for e in pel.findall(qn(t)):
                pel.remove(e)
        for sr in pel.findall(qn("c:ser")):
            _set_sppr(sr, '<a:ln w="19050"><a:noFill/></a:ln>')
            _set_marker(sr, "none")
        first_ax = pel.find(qn("c:axId"))
        first_ax.addprevious(_x(f'<c:hiLowLines {NS2}><c:spPr>{_ln("52514E", 12700)}</c:spPr></c:hiLowLines>'))
        first_ax.addprevious(_x(
            f'<c:upDownBars {NS2}><c:gapWidth val="60"/>'
            f'<c:upBars><c:spPr><a:solidFill><a:srgbClr val="{up}"/></a:solidFill>{_ln(up, 9525)}</c:spPr></c:upBars>'
            f'<c:downBars><c:spPr><a:solidFill><a:srgbClr val="{dn}"/></a:solidFill>{_ln(dn, 9525)}</c:spPr></c:downBars></c:upDownBars>'))
        ch.has_legend = False  # 시가·고가·저가·종가 계열 범례는 뜻이 없음 — 색 뜻은 표 아래 주석에
        return gf
    if ct in ("waterfall", "funnel"):
        plot.gap_width = 40 if ct == "waterfall" else 25
        plot.overlap = 100
        hide = []
        for i, (s, ser) in enumerate(zip(sers, plot.series)):
            if s.get("role") == "base":
                ser.format.fill.background()
                ser.format.line.fill.background()
                hide.append(i)
                continue
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = RGBColor.from_string(col_of(s, i))
            if s.get("role") == "extra":
                hide.append(i)
            if lab != "none" and s.get("role") != "extra":
                if ct == "funnel":  # 아주 짧은 막대는 글자가 밖으로 나가므로 잉크색
                    _ser_labels(ser, nf, LP.CENTER, "0B0B0B", True)
                else:
                    f = {"up": '"+"#,##0;"+"#,##0;;', "down": '"−"#,##0;"−"#,##0;;', "total": "#,##0;-#,##0;;"}[s["kind"]]
                    d = core.decimals_of(allv)
                    if d:
                        f = f.replace("#,##0", "#,##0." + "0" * d)
                    _ser_labels(ser, f, LP.CENTER, "FFFFFF" if s["kind"] == "total" else "0B0B0B", True)
        if ct == "funnel":
            ca.reverse_order = True
            va.visible = False
            va.has_major_gridlines = False
            ch.has_legend = False
        else:
            _legend_hide(cs, hide)
            # 0 을 가로지르는 막대는 '본 막대' 라벨이 일부만 가리킴 → 그 점은 실제 증감으로 직접 씀
            for j, st in enumerate(block["steps"]):
                if wf_parts(st)[2] and lab != "none":
                    k = 1 + ("up", "down", "total").index(st["kind"])
                    tf = plot.series[k].points[j].data_label.text_frame
                    tf.text = ("+" if st["kind"] == "up" else "") + core.fmt(st["value"], core.decimals_of(allv)).replace("−", "-")
        return gf
    if ct in ("lollipop", "dumbbell"):
        for i, (s, ser) in enumerate(zip(sers, plot.series)):
            c = col_of(s, i)
            _set_sppr(ser._element, '<a:ln w="19050"><a:noFill/></a:ln>')
            _set_marker(ser._element, "circle", 9 if ct == "lollipop" else 10, c)
            if ct == "lollipop":  # 줄기 = 아래로 100% 오차 막대
                ser._element.find(qn("c:cat")).addprevious(_x(
                    f'<c:errBars {NS2}><c:errDir val="y"/><c:errBarType val="minus"/><c:errValType val="percentage"/><c:noEndCap val="1"/>'
                    f'<c:val val="100"/><c:spPr>{_ln(c, 22225)}</c:spPr></c:errBars>'))
        if ct == "dumbbell":
            mk = pel.find(qn("c:marker"))
            hl = _x(f'<c:hiLowLines {NS2}><c:spPr>{_ln("C3C2B7", 38100)}</c:spPr></c:hiLowLines>')
            (mk.addprevious(hl) if mk is not None else pel.find(qn("c:axId")).addprevious(hl))
        if lab != "none" and (lab == "all" or len(block["cats"]) <= 16):
            for ser in plot.series:
                _ser_labels(ser, nf, LP.ABOVE)
        return gf
    if ct in ("bullet", "stack_line"):
        role = "target" if ct == "bullet" else "line"
        moves = [i for i, s in enumerate(sers) if s.get("role") == role]
        bars = [i for i, s in enumerate(sers) if s.get("role") != role]
        for i in bars:
            ser = plot.series[i]
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = RGBColor.from_string(col_of(sers[i], i))
            if ct == "stack_line":
                ser.format.line.color.rgb = RGBColor.from_string("FFFFFF")
                ser.format.line.width = Pt(1)
        plot.gap_width = 120 if ct == "bullet" else 80
        if lab != "none" and (lab == "all" or len(block["cats"]) <= 16) and ct == "bullet":
            for i in bars:
                _ser_labels(plot.series[i], nf, LP.INSIDE_END, "FFFFFF", True)
        styles = []
        for i in moves:
            if ct == "bullet":  # 목표 = 선 없는 가로 막대 표식
                styles.append(f'<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="dash"/><c:size val="20"/>'
                              f'<c:spPr><a:solidFill><a:srgbClr val="52514E"/></a:solidFill>{_ln("52514E", 9525)}</c:spPr></c:marker>')
            else:
                c = col_of(sers[i], i)
                styles.append(f'<c:spPr>{_ln(c, 28575)}</c:spPr><c:marker><c:symbol val="circle"/><c:size val="7"/>'
                              f'<c:spPr><a:solidFill><a:srgbClr val="{c}"/></a:solidFill>{_ln("FFFFFF", 12700)}</c:spPr></c:marker>')
        if moves:
            _split_to_line(pel, moves, styles)
        return gf
    raise ValueError(f"알 수 없는 차트: {ct}")


def _pptx_spark_rows(slide, block, x, y, w, h):
    """스파크라인(PPTX): 계열마다 [이름 | 작은 꺾은선 차트(축·격자 없음) | 마지막 값 · 처음 대비] 한 줄"""
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE as T
    from pptx.util import Emu, Pt
    sers = block["series"]
    m = max(1, len(sers))
    rh = int(min(Emu(914400 * 0.62), h / m))
    nw, vw = int(w * 0.22), int(w * 0.2)
    out = 0
    for j, s in enumerate(sers):
        ry = y + rh * j
        tb = slide.shapes.add_textbox(x, ry, nw, rh)
        tb.text_frame.text = svg.unit_title(s["name"], s.get("unit", ""))
        p = tb.text_frame.paragraphs[0]
        p.font.size, p.font.name = Pt(11), FONT
        p.font.color.rgb = RGBColor.from_string("52514E")
        cd = CategoryChartData(number_format=numfmt(s["values"], s.get("unit", "")))
        cd.categories = block["cats"]
        cd.add_series(s["name"], s["values"])
        gf = slide.shapes.add_chart(T.LINE_MARKERS, x + nw, ry, w - nw - vw, rh, cd)
        _spark_chart_style(gf.chart, hexcol(block["spec"], s.get("idx", j)))
        out += 1
        pv = [v for v in s["values"] if v is not None]
        if pv:
            dec = core.decimals_of(pv)
            ch_ = f"  {'▲' if pv[-1] > pv[0] else ('▼' if pv[-1] < pv[0] else '–')} {abs((pv[-1] - pv[0]) / pv[0] * 100):.1f}%" if pv[0] else ""
            tv = slide.shapes.add_textbox(x + w - vw, ry, vw, rh)
            tv.text_frame.text = core.fmt(pv[-1], dec, s.get("unit", "")).replace("−", "-") + ch_
            p = tv.text_frame.paragraphs[0]
            p.font.size, p.font.name, p.font.bold = Pt(11), FONT, True
    return out


def _spark_chart_style(ch, col):
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_MARKER_STYLE
    from pptx.oxml.ns import qn
    from pptx.util import Pt
    ch.has_title = False
    ch.has_legend = False
    ch.font.size = Pt(8)
    va, ca = ch.value_axis, ch.category_axis
    for ax in (va, ca):
        ax.visible = False
        ax.has_major_gridlines = False
        ax.has_minor_gridlines = False
    vals = [v for v in ch.plots[0].series[0].values if v is not None]
    if vals:  # 작은 그림이라 0 부터가 아니라 값 범위로
        lo, hi = min(vals), max(vals)
        pad = (hi - lo) * 0.08 or abs(hi) * 0.05 or 1
        va.minimum_scale, va.maximum_scale = lo - pad, hi + pad
    ser = ch.plots[0].series[0]
    c = RGBColor.from_string(col)
    ser.format.line.color.rgb = c
    ser.format.line.width = Pt(1.75)
    ser.smooth = False
    ser.marker.style = XL_MARKER_STYLE.NONE
    n = len(list(ser.values))
    if n:  # 마지막 점만 표식
        pt = ser.points[n - 1]
        pt.marker.style = XL_MARKER_STYLE.CIRCLE
        pt.marker.size = 6
        pt.marker.format.fill.solid()
        pt.marker.format.fill.fore_color.rgb = c
        pt.marker.format.line.color.rgb = RGBColor.from_string("FFFFFF")
    ch._chartSpace.find(qn("c:chart")).find(qn("c:autoTitleDeleted")).set("val", "1")


def png_for(table, spec, wd, name, w=960, h=560):
    """그림 모드용 PNG (rsvg-convert). 없으면 None"""
    p = os.path.join(wd, name)
    return p if svg_to_png(svg.render(table, spec, w, h), p, 2) else None


# ── 새 종류: XLSX ─────────────────────────────────────────────────────────────
def _xl_hide_legend(ch, idxs):
    from openpyxl.chart.legend import LegendEntry
    if ch.legend is not None and idxs:
        ch.legend.legendEntry = [LegendEntry(idx=i, delete=True) for i in idxs]


def _xl_sp(fill=None, line=None, nofill=False, noline=False, w=None):
    from openpyxl.chart.shapes import GraphicalProperties
    gp = GraphicalProperties()
    if nofill:
        gp.noFill = True
    elif fill:
        gp.solidFill = fill
    if noline:
        gp.line.noFill = True
    elif line:
        gp.line.solidFill = line
        if w:
            gp.line.width = w
    return gp


def _xlsx_new(sh, table, spec, b, r0, colref, rowref, head_fill, cell_ref):
    """새 종류의 그림 시트 표 + openpyxl 차트. 반환: (크기 정보, 차트|None, 스파크라인 목록)"""
    from openpyxl.chart import BarChart, BubbleChart, Reference, Series
    from openpyxl.chart.label import DataLabelList
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter as L

    tf = Font(bold=True, name=FONT)
    ct = b.get("ctype")

    def head(cols):
        for j, h in enumerate(cols, 1):
            if h:
                sh.cell(r0, j, h).font = tf
                sh.cell(r0, j).fill = head_fill

    if b["mode"] == "spark":  # 계열 = 행, 기간 = 열 → 끝 열에 엑셀 스파크라인
        cats, sers = b["cats"], b["series"]
        head([b.get("x_title") or "계열"] + [str(c) for c in cats] + ["추세", "마지막 값"])
        for i, s in enumerate(sers, 1):
            r = r0 + i
            sh.cell(r, 1, svg.unit_title(s["name"], s.get("unit", "")))
            for j, c in enumerate(cats):
                sh.cell(r, j + 2, cell_ref(s, c, j))
                sh.cell(r, j + 2).number_format = numfmt(s["values"], s.get("unit", ""))
            last = L(len(cats) + 1)
            sh.row_dimensions[r].height = 26
            sh.cell(r, len(cats) + 3, f'=LOOKUP(2,1/({L(2)}{r}:{last}{r}<>""),{L(2)}{r}:{last}{r})')
            sh.cell(r, len(cats) + 3).number_format = numfmt(s["values"], s.get("unit", ""))
        sh.column_dimensions[L(len(cats) + 2)].width = 26
        sparks = [(f"{_q(sh.title)}!{L(2)}{r0 + i}:{L(len(cats) + 1)}{r0 + i}", f"{L(len(cats) + 2)}{r0 + i}",
                   hexcol(b["spec"], s.get("idx", i - 1))) for i, s in enumerate(sers, 1)]
        return {"nrows": len(sers) + 1, "ncols": len(cats) + 3}, None, sparks
    if ct == "waterfall":
        steps = b["steps"]
        head([b.get("x_title") or "항목", b["src"], "막대 시작", "막대 끝", "기준(투명)", "증가", "감소", "합계",
              "증가 (0 아래)", "감소 (0 아래)", "합계 (0 아래)"])
        xs = core.col_values(table, spec["x"]) if spec.get("x") else None
        for i, st in enumerate(steps, 1):
            r = r0 + i
            sh.cell(r, 1, st["cat"])
            tot = st["kind"] == "total"
            if st.get("auto"):
                sh.cell(r, 2, f"=D{r - 1}" if i > 1 else 0)
            elif xs is not None and xs.count(st["cat"]) == 1:
                sh.cell(r, 2, rowref(b["src"], xs.index(st["cat"])))
            else:
                sh.cell(r, 2, st["value"])
            sh.cell(r, 3, 0 if (tot or i == 1) else f"=D{r - 1}")
            sh.cell(r, 4, f"=B{r}" if tot else f"=C{r}+B{r}")
            lo, hi = f"MIN(C{r},D{r})", f"MAX(C{r},D{r})"
            main = f"IF({lo}>=0,{hi}-{lo},IF({hi}<=0,{lo}-{hi},{hi}))"
            ext = f"IF(AND({lo}<0,{hi}>0),{lo},0)"
            sh.cell(r, 5, f"=IF({lo}>=0,{lo},IF({hi}<=0,{hi},0))")
            if tot:
                vals = ["0", "0", main, "0", "0", ext]
            else:
                vals = [f"IF(B{r}>=0,{main},0)", f"IF(B{r}<0,{main},0)", "0", f"IF(B{r}>=0,{ext},0)", f"IF(B{r}<0,{ext},0)", "0"]
            for j, f in enumerate(vals, 6):
                sh.cell(r, j, "=" + f)
            for j in range(2, 12):
                sh.cell(r, j).number_format = numfmt([st["value"]], b.get("unit", ""))
        n = len(steps)
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap, ch.gapWidth = "col", "stacked", 100, 40
        ch.add_data(Reference(sh, min_col=5, max_col=11, min_row=r0, max_row=r0 + n), titles_from_data=True)
        ch.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + n))
        cols = [c.lstrip("#").upper() for c in svg.wf_colors(spec)]
        for i, s in enumerate(ch.series):
            if i == 0:
                s.graphicalProperties = _xl_sp(nofill=True, noline=True)
            else:
                c = cols[(i - 1) % 3]
                s.graphicalProperties = _xl_sp(fill=c, line=c)
        dec = core.decimals_of([st["value"] for st in steps])
        z = "#,##0" + ("." + "0" * dec if dec else "")
        if b["labels"] != "none":
            for i, f in ((1, f'"+"{z};"+"{z};;'), (2, f'"−"{z};"−"{z};;'), (3, f"{z};-{z};;")):
                ch.series[i].dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
                ch.series[i].dLbls.numFmt = f
        _xl_common(ch, dict(b, labels="none"), 3, "waterfall")
        _xl_hide_legend(ch, [0, 4, 5, 6])
        return {"nrows": n + 1, "ncols": 11}, ch, []
    if ct == "funnel":
        cats, vals = b["cats"], b["series"][1]["values"]
        head([b.get("x_title") or "단계", b["src"], "여백(투명)", "이전 단계 대비", "처음 대비"])
        xs = core.col_values(table, spec["x"]) if spec.get("x") else None
        n = len(cats)
        for i, (c, v) in enumerate(zip(cats, vals), 1):
            r = r0 + i
            sh.cell(r, 1, c)
            sh.cell(r, 2, rowref(b["src"], xs.index(c)) if xs is not None and xs.count(c) == 1 else v)
            sh.cell(r, 3, f"=(MAX($B${r0 + 1}:$B${r0 + n})-B{r})/2")
            sh.cell(r, 4, f"=B{r}/B{r - 1}" if i > 1 else None)
            sh.cell(r, 5, f"=B{r}/$B${r0 + 1}")
            sh.cell(r, 2).number_format = numfmt(vals, b.get("unit", ""))
            sh.cell(r, 3).number_format = numfmt(vals)
            sh.cell(r, 4).number_format = sh.cell(r, 5).number_format = "0.0%"
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap, ch.gapWidth = "bar", "stacked", 100, 25
        ch.add_data(Reference(sh, min_col=3, min_row=r0, max_row=r0 + n), titles_from_data=True)
        ch.add_data(Reference(sh, min_col=2, min_row=r0, max_row=r0 + n), titles_from_data=True)
        ch.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + n))
        ch.series[0].graphicalProperties = _xl_sp(nofill=True, noline=True)
        c = hexcol(b["spec"], b["series"][1].get("idx", 0))
        ch.series[1].graphicalProperties = _xl_sp(fill=c, line=c)
        if b["labels"] != "none":
            ch.series[1].dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
        ch.x_axis.scaling.orientation = "maxMin"
        _xl_common(ch, dict(b, labels="none", y_title=""), 1, "funnel")
        ch.y_axis.delete = True
        ch.y_axis.majorGridlines = None
        ch.legend = None
        return {"nrows": n + 1, "ncols": 5}, ch, []
    if ct == "bubble":
        ch = BubbleChart()
        ch.style = 18
        ch.bubbleScale = 60
        x, y, z, lab = spec["x"], spec["ys"][0], spec.get("size"), spec.get("label")
        gv = core.col_values(table, spec["group"]) if spec.get("group") else None
        xs_all, ys_all = core.col_nums(table, x), core.col_nums(table, y)
        zs_all = core.col_nums(table, z) if z else [1] * len(xs_all)
        col = 1
        maxn = 0
        for si, s in enumerate(b["series"]):
            hh = [lab or "이름", x, y, z or "크기"]
            for j, h in enumerate(hh):
                sh.cell(r0, col + j, (f"{s['name']} · " if gv else "") + h if j else h).font = tf
            rows = [i for i in range(len(xs_all)) if xs_all[i] is not None and ys_all[i] is not None and zs_all[i] is not None
                    and zs_all[i] > 0 and (gv is None or gv[i] == s["name"])]
            for k, i in enumerate(rows, 1):
                sh.cell(r0 + k, col, rowref(lab, i) if lab else k)
                sh.cell(r0 + k, col + 1, rowref(x, i))
                sh.cell(r0 + k, col + 2, rowref(y, i))
                sh.cell(r0 + k, col + 3, rowref(z, i) if z else 1)
            n = len(rows)
            maxn = max(maxn, n)
            se = Series(values=Reference(sh, min_col=col + 2, min_row=r0 + 1, max_row=r0 + n),
                        xvalues=Reference(sh, min_col=col + 1, min_row=r0 + 1, max_row=r0 + n),
                        zvalues=Reference(sh, min_col=col + 3, min_row=r0 + 1, max_row=r0 + n), title=s["name"])
            c = hexcol(b["spec"], s.get("idx", si))
            se.graphicalProperties = _xl_sp(fill=c, line=c)
            ch.series.append(se)
            col += 5
        _xl_common(ch, b, len(b["series"]), "bubble")
        if len(b["series"]) < 2:
            ch.legend = None
        return {"nrows": maxn + 1, "ncols": col - 2}, ch, []
    return None


def _xlsx_new_cat(sh, spec, b, r0, ncat, nser):
    """범주형 표가 이미 쓰인 뒤(제목 행 r0, A열=범주, B…=계열) 새 범주형 종류의 차트"""
    from openpyxl.chart import BarChart, DoughnutChart, LineChart, RadarChart, Reference, StockChart
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.error_bar import ErrorBars
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.updown_bars import UpDownBars
    ct = b["ctype"]
    sers = b["series"]
    data = lambda c0, c1: Reference(sh, min_col=c0, max_col=c1, min_row=r0, max_row=r0 + ncat)
    catref = Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + ncat)
    col_of = lambda i: hexcol(b["spec"], sers[i].get("idx", i))

    def line_style(s, c, line=True, symbol="circle", size=7):
        s.graphicalProperties = _xl_sp(line=c, w=28575) if line else _xl_sp(noline=True)
        s.marker = Marker(symbol=symbol, size=size)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=c)
        s.marker.graphicalProperties.line.solidFill = "FFFFFF" if symbol != "dash" else c
        s.smooth = False
    if ct == "radar":
        ch = RadarChart()
        ch.type = "marker"
        ch.add_data(data(2, nser + 1), titles_from_data=True)
        ch.set_categories(catref)
        for i, s in enumerate(ch.series):
            line_style(s, col_of(i))
        _xl_common(ch, dict(b, x_title="", y_title=""), nser, "radar")
        return ch
    if ct == "stock":
        ch = StockChart()
        ch.add_data(data(2, 5), titles_from_data=True)
        ch.set_categories(catref)
        for s in ch.series:
            s.graphicalProperties = _xl_sp(noline=True)
            s.marker = Marker(symbol="none")
            s.smooth = False
        ch.hiLowLines = ChartLines(spPr=_xl_sp(line="52514E", w=12700))
        up, dn = b["up"].lstrip("#").upper(), b["down"].lstrip("#").upper()
        ch.upDownBars = UpDownBars(gapWidth=60, upBars=ChartLines(spPr=_xl_sp(fill=up, line=up)), downBars=ChartLines(spPr=_xl_sp(fill=dn, line=dn)))
        vals = [v for s in sers for v in s["values"] if v is not None]
        a_, b_, _, _ = svg.nice_domain(min(vals), max(vals), zero=False)
        _xl_common(ch, dict(b, labels="none"), 1, "stock")
        ch.y_axis.scaling.min, ch.y_axis.scaling.max = a_, b_
        ch.legend = None
        return ch
    if ct == "semi_doughnut":
        r = r0 + ncat + 1
        sh.cell(r, 1, "(숨김) 합계")
        sh.cell(r, 2, f"=SUM(B{r0 + 1}:B{r0 + ncat})")
        ch = DoughnutChart(holeSize=50, firstSliceAng=270)
        ch.add_data(Reference(sh, min_col=2, min_row=r0, max_row=r), titles_from_data=True)
        ch.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r))
        s = ch.series[0]
        for p in range(ncat + 1):
            pt = DataPoint(idx=p)
            if p < ncat:
                pt.graphicalProperties.solidFill = hexcol(b["spec"], p)
                pt.graphicalProperties.line.solidFill = "FFFFFF"
            else:
                pt.graphicalProperties.noFill = True
                pt.graphicalProperties.line.noFill = True
            s.dPt.append(pt)
        _xl_common(ch, dict(b, labels="none"), 1, "doughnut")
        if b["labels"] != "none":
            from openpyxl.chart.label import DataLabel
            s.dLbls = DataLabelList(showVal=True, showPercent=False, showSerName=False, showCatName=False, showLegendKey=False)
            s.dLbls.dLbl = [DataLabel(idx=ncat, showVal=False, showPercent=False, showSerName=False, showCatName=False, showLegendKey=False)]
        _xl_hide_legend(ch, [ncat])
        return ch
    if ct in ("lollipop", "dumbbell"):
        ch = LineChart()
        ch.add_data(data(2, nser + 1), titles_from_data=True)
        ch.set_categories(catref)
        for i, s in enumerate(ch.series):
            c = col_of(i)
            line_style(s, c, line=False, size=9)
            if ct == "lollipop":
                s.errBars = ErrorBars(errDir="y", errBarType="minus", errValType="percentage", val=100, noEndCap=True,
                                      spPr=_xl_sp(line=c, w=22225))
        if ct == "dumbbell":
            ch.hiLowLines = ChartLines(spPr=_xl_sp(line="C3C2B7", w=38100))
        _xl_common(ch, b, nser if ct == "dumbbell" else 1, "line")
        return ch
    if ct in ("bullet", "stack_line"):
        role = "target" if ct == "bullet" else "line"
        nb = sum(1 for s in sers if s.get("role") != role)
        c1 = BarChart()
        c1.type = "col"
        c1.grouping = "clustered" if ct == "bullet" else "stacked"
        if ct == "stack_line":
            c1.overlap = 100
        c1.gapWidth = 120 if ct == "bullet" else 80
        c1.add_data(data(2, nb + 1), titles_from_data=True)
        c1.set_categories(catref)
        for i, s in enumerate(c1.series):
            c = col_of(i)
            s.graphicalProperties = _xl_sp(fill=c, line="FFFFFF" if ct == "stack_line" else c)
        c2 = LineChart()
        c2.add_data(data(nb + 2, nser + 1), titles_from_data=True)
        c2.set_categories(catref)
        for i, s in enumerate(c2.series):
            if ct == "bullet":
                line_style(s, "52514E", line=False, symbol="dash", size=20)
            else:
                line_style(s, col_of(nb + i))
        _xl_common(c1, dict(b, labels="none" if ct == "stack_line" else b["labels"]), nser, ct)
        c1 += c2  # 같은 축(axId 10·100)을 함께 씀
        return c1
    raise ValueError(f"알 수 없는 차트: {ct}")


def _xlsx_add_sparklines(path, sheet_sparks):
    """저장한 XLSX 손질: (1) 차트 값 라벨 numFmt 에 sourceLinked="0" (2) openpyxl 은 스파크라인을 못 쓴다 →
    시트 XML 끝(extLst)에 x14:sparklineGroups 를 넣는다(XlsxWriter 가 쓰는 것과 같은 구조).
    sheet_sparks: {시트 번호(0부터): [(데이터 범위, 놓을 칸, 색), ...]}"""
    import zipfile
    tmp = path + ".tmp"
    n = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"xl/charts/chart\d+\.xml", item.filename):
                # openpyxl 은 값 라벨 서식(numFmt)에 sourceLinked 를 안 써서 일부 프로그램이 '원본 서식'으로 읽음 → 명시
                data = re.sub(rb'(<numFmt formatCode="[^"]*")(/>)', rb'\1 sourceLinked="0"\2', data)
            m = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if m and int(m.group(1)) - 1 in sheet_sparks:
                groups = []
                for rng, loc, col in sheet_sparks[int(m.group(1)) - 1]:
                    c = "FF" + col.upper()
                    groups.append(
                        f'<x14:sparklineGroup displayEmptyCellsAs="gap" markers="0" high="1" low="1" last="1" lineWeight="1.5">'
                        f'<x14:colorSeries rgb="{c}"/><x14:colorNegative rgb="FFE34948"/><x14:colorAxis rgb="FF000000"/>'
                        f'<x14:colorMarkers rgb="{c}"/><x14:colorFirst rgb="{c}"/><x14:colorLast rgb="{c}"/>'
                        f'<x14:colorHigh rgb="FF52514E"/><x14:colorLow rgb="FF898781"/>'
                        f'<x14:sparklines><x14:sparkline><xm:f>{svg.esc(rng)}</xm:f><xm:sqref>{loc}</xm:sqref></x14:sparkline></x14:sparklines>'
                        f'</x14:sparklineGroup>')
                    n += 1
                ext = ('<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">'
                       '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' + "".join(groups) +
                       '</x14:sparklineGroups></ext>')
                s = data.decode("utf-8")
                if "<extLst>" in s:
                    s = s.replace("</extLst>", ext + "</extLst>", 1)
                else:
                    s = s.replace("</worksheet>", "<extLst>" + ext + "</extLst></worksheet>", 1)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)
    return n


# ── 공통: 명세 → 문서용 블록 ────────────────────────────────────────────────
def blocks_for(table, spec):
    """한 그림 → 문서에 넣을 블록 목록. kind 별로 차트/표 블록을 만든다.
    chart 블록: {mode: chart, ctype, cats, series:[{name, values, color}], unit, ...}"""
    k = spec["kind"]
    title = spec.get("title") or core.default_title(table, spec)
    base = {"title": title, "spec": spec, "legend": spec.get("legend") or "top", "labels": spec.get("labels") or "auto"}
    nb = new_blocks(table, spec, base, title)
    if nb is not None:
        return nb
    if k == "dual":
        d = core.category_data(table, spec)
        sers = [s for s in d["series"] if not s["hidden"]]
        left = [s for s in sers if not s.get("right")]
        right = [s for s in sers if s.get("right")]
        lu, ru = (left[0]["unit"] if left else ""), (right[0]["unit"] if right else "")
        shape = spec.get("dual_shape") or "bar_line"
        zero = spec.get("zero_base", True) is not False
        L, R, aligned = svg.dual_domains([v for s in left for v in s["values"]], [v for s in right for v in s["values"]], zero,
                                         left_zero=shape == "bar_line" or zero)
        return [dict(base, mode="chart", ctype="dual", cats=d["cats"], series=left + right, nleft=len(left), shape=shape,
                     unit=lu, unit2=ru, dom=L, dom2=R, aligned=aligned, zero=zero,
                     neg=any(v is not None and v < 0 for s in sers for v in s["values"]),
                     x_title=spec.get("x_title", spec.get("x") or ""),
                     y_title=spec.get("y_title") or svg.unit_title(left[0]["name"] if len(left) == 1 else "", lu),
                     y2_title=spec.get("y2_title") or (svg.unit_title(right[0]["name"], ru) if right else ""), note=d["note"])]
    if k in core.CAT_KINDS and k not in ("heatmap", "panels"):
        d = core.category_data(table, spec)
        sers = [s for s in d["series"] if not s["hidden"]]
        if k in ("lollipop", "dumbbell"):  # 막대사탕 1계열, 덤벨 2계열
            sers = [s for s in sers if not s.get("role")][:1 if k == "lollipop" else 2]
        if k in ("pie", "doughnut"):
            s = sers[0]
            items = [(c, v) for c, v in zip(d["cats"], s["values"]) if v and v > 0]
            return [dict(base, mode="chart", ctype=k, cats=[c for c, _ in items],
                         series=[{"name": s["name"], "values": [v for _, v in items], "idx": 0}], unit=s["unit"])]
        if k == "index":
            for s in sers:
                b0 = next((v for v in s["values"] if v), None)
                s["values"] = [None if v is None or not b0 else round(100 * v / b0, 4) for v in s["values"]]
        unit = "" if k == "index" else (sers[0]["unit"] if sers else "")
        ytitle = spec.get("y_title")
        if ytitle is None:
            ytitle = "지수 (첫 값=100)" if k == "index" else ("비율 (%)" if k == "col_stack100" else
                                                         svg.unit_title("" if len(sers) > 1 or spec.get("group") else (sers[0]["name"] if sers else ""), unit))
        return [dict(base, mode="chart", ctype=k, cats=d["cats"], series=sers, unit=unit, x_title=spec.get("x_title", spec.get("x") or ""),
                     y_title=ytitle, note=d["note"])]
    if k == "panels":
        d = core.category_data(table, spec)
        sers = [s for s in d["series"] if not s["hidden"]][:2]
        shape = "line" if (d["xtype"] == "date" and len(d["cats"]) > 6) else "col"
        return [dict(base, mode="chart", ctype=shape, cats=d["cats"], series=[s], unit=s["unit"], x_title=spec.get("x") or "",
                     y_title=svg.unit_title(s["name"], s["unit"]), title=f"{title} — {s['name']}", panel=i + 1)
                for i, s in enumerate(sers)]
    if k == "heatmap":
        d = core.category_data(table, spec)
        sers = [s for s in d["series"] if not s["hidden"]]
        return [dict(base, mode="heat", cats=d["cats"], series=sers, unit=sers[0]["unit"] if sers else "", x_title=spec.get("x") or "")]
    if k == "scatter":
        sers = [s for s in core.scatter_data(table, spec) if not s["hidden"]]
        return [dict(base, mode="chart", ctype="scatter", series=sers, trend=bool(spec.get("trend")),
                     x_title=spec.get("x_title") or svg.unit_title(spec["x"], core.unit_of(table, spec["x"])),
                     y_title=spec.get("y_title") or svg.unit_title(spec["ys"][0], core.unit_of(table, spec["ys"][0])))]
    if k == "hist":
        s = core.dist_data(table, spec)[0]
        bins = core.hist_bins(s["values"], int(spec.get("bins") or 0) or None)
        ew = core.decimals_of([b[0] for b in bins] + [bins[-1][1]] if bins else [])
        cats = [f"{core.fmt(a, ew)}~{core.fmt(b, ew)}" for a, b, _ in bins]
        u = core.unit_of(table, s["name"])
        return [dict(base, mode="chart", ctype="hist", cats=cats, bins=bins, src=s["name"],
                     series=[{"name": "빈도", "values": [c for _, _, c in bins], "idx": 0}], unit="",
                     x_title=spec.get("x_title") or svg.unit_title(s["name"], u), y_title=spec.get("y_title") or "빈도 (개)")]
    if k == "box":
        sers = [s for s in core.dist_data(table, spec) if not s["hidden"]]
        stats = [(s, core.box_stats(s["values"])) for s in sers]
        stats = [(s, st) for s, st in stats if st]
        u = core.unit_of(table, spec["ys"][0])
        return [dict(base, mode="chart", ctype="box", cats=[s["name"] for s, _ in stats], stats=[st for _, st in stats],
                     series=[s for s, _ in stats], unit=u, x_title=spec.get("group") or "",
                     y_title=spec.get("y_title") or svg.unit_title(spec["ys"][0], u))]
    raise ValueError(f"알 수 없는 그래프 종류: {k}")


def data_rows(block):
    """블록 → (머리글, 행들) — 문서에 함께 넣는 '편집 가능한 데이터 표'"""
    if block["mode"] in ("heat", "image", "spark") or (block["mode"] == "chart" and block["ctype"] not in (
            "scatter", "box", "hist", "bubble", "waterfall", "funnel")):
        sers = [s for s in block["series"] if s.get("role") != "base"]
        head = [block.get("x_title") or "항목"] + [s["name"] for s in sers]
        rows = [[c] + [s["values"][i] for s in sers] for i, c in enumerate(block["cats"])]
        if block.get("ctype") == "bullet" and len(sers) == 2:  # 달성률
            head.append("달성률(%)")
            for r in rows:
                r.append(round(r[1] / r[2] * 100, 1) if r[1] is not None and r[2] else None)
        return head, rows
    if block["ctype"] == "waterfall":
        return ([block.get("x_title") or "항목", block["src"], "막대 시작", "막대 끝", "구분"],
                [[st["cat"], st["value"], st["a"], st["b"], KIND_KO[st["kind"]]] for st in block["steps"]])
    if block["ctype"] == "funnel":
        vals = block["series"][1]["values"]
        return ([block.get("x_title") or "단계", block["src"], "이전 단계 대비(%)", "처음 대비(%)"],
                [[c, v, round(v / vals[i - 1] * 100, 1) if i and vals[i - 1] else None, round(v / vals[0] * 100, 1)]
                 for i, (c, v) in enumerate(zip(block["cats"], vals))])
    if block["ctype"] == "bubble":
        sp = block["spec"]
        multi = len(block["series"]) > 1
        head = [sp.get("label") or "이름"] + ([sp.get("group") or "계열"] if multi else []) + [sp["x"], sp["ys"][0], sp.get("size") or "크기"]
        rows = [[lab] + ([s["name"]] if multi else []) + [x, y, z] for s in block["series"] for x, y, z, lab in s["points"]]
        return head, rows
    if block["ctype"] == "scatter":
        sp = block["spec"]
        multi = len(block["series"]) > 1
        head = ([sp.get("group") or "계열"] if multi else []) + [sp["x"], sp["ys"][0]]
        rows = [([s["name"]] if multi else []) + [x, y] for s in block["series"] for x, y in s["points"]]
        return head, rows
    if block["ctype"] == "hist":
        return ["구간", "하한", "상한", "빈도"], [[c, a, b, n] for c, (a, b, n) in zip(block["cats"], block["bins"])]
    if block["ctype"] == "box":
        head = [block.get("x_title") or "계열", "n", "최소", "Q1", "중앙값", "Q3", "최대", "평균"]
        return head, [[c, st["n"], st["min"], st["q1"], st["median"], st["q3"], st["max"], round(st["mean"], 6)]
                      for c, st in zip(block["cats"], block["stats"])]
    return [], []


def cell_text(v, dec=None):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer() and abs(v) < 1e15 and dec in (None, 0):
            return f"{int(v):,}"
        return core.fmt(v, dec if dec is not None else core.decimals_of([v], 3)).replace("−", "-")
    return str(v)


# ── PPTX 네이티브 차트 (PPTX·DOCX 공통) ──────────────────────────────────────
def _pptx_chart(slide_shapes, block, x, y, cx, cy):
    from lxml import etree
    from pptx.chart.data import CategoryChartData, XyChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
    from pptx.oxml.ns import qn
    from pptx.util import Pt

    ct = block["ctype"]
    if ct in NEW_CT:
        return _pptx_chart_new(slide_shapes, block, x, y, cx, cy)
    spec = block["spec"]
    T = XL_CHART_TYPE
    types = {"col": T.COLUMN_CLUSTERED, "col_group": T.COLUMN_CLUSTERED, "col_stack": T.COLUMN_STACKED,
             "col_stack100": T.COLUMN_STACKED_100, "bar": T.BAR_CLUSTERED, "bar_group": T.BAR_CLUSTERED,
             "bar_stack": T.BAR_STACKED, "line": T.LINE_MARKERS, "index": T.LINE_MARKERS, "area": T.AREA,
             "area_stack": T.AREA_STACKED, "pie": T.PIE, "doughnut": T.DOUGHNUT, "hist": T.COLUMN_CLUSTERED,
             "box": T.COLUMN_STACKED, "scatter": T.XY_SCATTER,
             "dual": T.LINE_MARKERS if block.get("shape") == "line_line" else T.COLUMN_CLUSTERED}
    unit = block.get("unit", "")
    if ct == "scatter":
        cd = XyChartData()
        for s in block["series"]:
            ser = cd.add_series(s["name"])
            for px, py in s["points"]:
                ser.add_data_point(px, py)
        allv = [p[1] for s in block["series"] for p in s["points"]]
    elif ct == "box":
        cd = CategoryChartData(number_format="General")
        cd.categories = block["cats"]
        st = block["stats"]
        cd.add_series("Q1", [s["q1"] for s in st])
        cd.add_series("중앙값−Q1", [s["median"] - s["q1"] for s in st])
        cd.add_series("Q3−중앙값", [s["q3"] - s["median"] for s in st])
        allv = [s["max"] for s in st] + [s["min"] for s in st]
    elif ct == "dual":
        nl = block["nleft"]
        allv = [v for s in block["series"][:nl] for v in s["values"] if v is not None]
        cd = CategoryChartData(number_format=numfmt(allv, unit))
        cd.categories = block["cats"]
        for i, s in enumerate(block["series"]):
            # 오른쪽 축 계열은 이름에 표시 → 범례·내장 워크북 머리글에 그대로 보임
            cd.add_series(s["name"] + (" (오른쪽 축)" if i >= nl else ""), s["values"],
                          number_format=numfmt(s["values"], s.get("unit", "")))
    else:
        allv = [v for s in block["series"] for v in s["values"] if v is not None]
        cd = CategoryChartData(number_format=numfmt(allv, unit))
        cd.categories = block["cats"]
        for s in block["series"]:
            cd.add_series(s["name"], s["values"])
    gf = slide_shapes.add_chart(types[ct], x, y, cx, cy, cd)
    ch = gf.chart
    ch.font.size = Pt(10)
    ch.font.name = FONT
    ch.font.color.rgb = RGBColor.from_string("52514E")
    ch.has_title = True
    ch.chart_title.text_frame.text = block["title"]
    tp = ch.chart_title.text_frame.paragraphs[0]
    tp.font.size, tp.font.bold, tp.font.name = Pt(13), True, FONT
    tp.font.color.rgb = RGBColor.from_string("0B0B0B")
    nser = len(block["series"])
    want_legend = (nser >= 2 or ct in ("pie", "doughnut")) and block["legend"] != "none" and ct not in ("box", "hist")
    ch.has_legend = want_legend
    if want_legend:
        ch.legend.position = {"top": XL_LEGEND_POSITION.TOP, "bottom": XL_LEGEND_POSITION.BOTTOM,
                              "right": XL_LEGEND_POSITION.RIGHT}.get(block["legend"], XL_LEGEND_POSITION.TOP)
        if ct in ("pie", "doughnut"):
            ch.legend.position = XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(10)
    plot = ch.plots[0]
    if ct in ("pie", "doughnut"):
        plot.vary_by_categories = True
        ser = plot.series[0]
        for i in range(len(block["cats"])):
            pt = ser.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = RGBColor.from_string(hexcol(spec, i))
            pt.format.line.color.rgb = RGBColor.from_string("FFFFFF")
            pt.format.line.width = Pt(1.5)
        if block["labels"] != "none":
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_percentage, dl.show_value, dl.show_category_name = True, False, False
            dl.number_format, dl.number_format_is_linked = "0.0%", False
            dl.font.size, dl.font.bold = Pt(10), True
            dl.font.color.rgb = RGBColor.from_string("FFFFFF")
        if ct == "doughnut":
            hs = plot._element.find(qn("c:holeSize"))
            if hs is not None:
                hs.set("val", "55")
        return gf
    # 축
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor.from_string("E1E0D9")
    va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(9)
    va.tick_labels.font.color.rgb = RGBColor.from_string("898781")
    if ct not in ("col_stack100",):
        va.tick_labels.number_format = numfmt(allv, unit) if ct not in ("box",) else numfmt(allv)
        va.tick_labels.number_format_is_linked = False
    ca = ch.category_axis
    ca.format.line.color.rgb = RGBColor.from_string("C3C2B7")
    ca.tick_labels.font.size = Pt(9)
    ca.tick_labels.font.color.rgb = RGBColor.from_string("52514E")
    ca.has_major_gridlines = False
    if ct == "scatter":
        ca.has_major_gridlines = True
        ca.major_gridlines.format.line.color.rgb = RGBColor.from_string("E1E0D9")
    for ax, txt in ((va, block.get("y_title")), (ca, block.get("x_title"))):
        if txt:
            ax.has_title = True
            ax.axis_title.text_frame.text = txt
            f = ax.axis_title.text_frame.paragraphs[0].font
            f.size, f.bold, f.name = Pt(9.5), False, FONT
            f.color.rgb = RGBColor.from_string("52514E")
    if ct in ("bar", "bar_group", "bar_stack"):
        ca.reverse_order = True  # 첫 항목이 위에 오도록 (SVG 와 같은 순서)
    if ct == "index":
        va.minimum_scale = None
    # 계열 서식
    is_line = ct in ("line", "index")
    is_line = is_line or (ct == "dual" and block.get("shape") == "line_line")
    for i, ser in enumerate(plot.series):
        src = block["series"][i] if i < nser else None
        col = RGBColor.from_string(hexcol(spec, src["idx"] if src and "idx" in src else i))
        if ct == "box":
            continue
        if is_line or ct == "scatter":
            if is_line:
                ser.format.line.color.rgb = col
                ser.format.line.width = Pt(2.25)
                ser.smooth = False
            ser.marker.style = XL_MARKER_STYLE.CIRCLE
            ser.marker.size = 7
            ser.marker.format.fill.solid()
            ser.marker.format.fill.fore_color.rgb = col
            ser.marker.format.line.color.rgb = RGBColor.from_string("FFFFFF")
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = col
            if ct in ("col_stack", "col_stack100", "bar_stack", "area_stack", "hist"):
                ser.format.line.color.rgb = RGBColor.from_string("FFFFFF")
                ser.format.line.width = Pt(1)
    if ct.startswith("col") or ct.startswith("bar") or ct in ("hist", "box"):
        plot.gap_width = 6 if ct == "hist" else (60 if nser > 1 else 90)
        if ct in ("col_group", "bar_group", "col", "bar") and nser > 1:
            plot.overlap = -8
    # 값 라벨
    lab = block["labels"]
    if ct not in ("box", "scatter") and lab != "none" and (lab == "all" or (nser == 1 and len(block["cats"]) <= 16 and ct != "dual")
                                                          or (ct in ("col_stack",) and False)):
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(9)
        dl.font.color.rgb = RGBColor.from_string("0B0B0B")
        dl.number_format = numfmt(allv, unit) if ct != "col_stack100" else '0.0"%"'
        dl.number_format_is_linked = False
        dl.show_value = True
        if ct in ("col", "col_group", "bar", "bar_group", "hist"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
        elif is_line:
            dl.position = XL_LABEL_POSITION.ABOVE
        elif ct in ("col_stack", "bar_stack", "col_stack100"):
            dl.position = XL_LABEL_POSITION.CENTER
            dl.font.color.rgb = RGBColor.from_string("FFFFFF")
    C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    if ct == "dual":
        _pptx_dual_axis(ch, block, plot)
    if ct == "scatter" and block.get("trend"):
        for ser in plot.series:
            el = ser._element
            tl = etree.fromstring(f'<c:trendline xmlns:c="{C}" xmlns:a="{A}"><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="52514E"/></a:solidFill></a:ln></c:spPr>'
                                  f'<c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="0"/></c:trendline>')
            anchor = el.find(qn("c:errBars"))
            if anchor is None:
                anchor = el.find(qn("c:xVal"))
            anchor.addprevious(tl)
    if ct == "box":
        st = block["stats"]
        s0, s1, s2 = plot.series
        s0.format.fill.background()
        s0.format.line.fill.background()
        for s, a in ((s1, "B7D3F6"), (s2, "86B6EF")):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = RGBColor.from_string(a)
            s.format.line.color.rgb = RGBColor.from_string("2A78D6")
            s.format.line.width = Pt(1.5)

        def errbars(kind, vals):
            pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
            return etree.fromstring(
                f'<c:errBars xmlns:c="{C}" xmlns:a="{A}"><c:errBarType val="{kind}"/><c:errValType val="cust"/><c:noEndCap val="0"/>'
                f'<c:{kind}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>{pts}</c:numLit></c:{kind}>'
                f'<c:spPr><a:ln w="15875"><a:solidFill><a:srgbClr val="52514E"/></a:solidFill></a:ln></c:spPr></c:errBars>')
        s0._element.find(qn("c:cat")).addprevious(errbars("minus", [x["q1"] - x["min"] for x in st]))
        s2._element.find(qn("c:cat")).addprevious(errbars("plus", [x["max"] - x["q3"] for x in st]))
        ch.has_legend = False
    return gf


def _axis_title_xml(text, color, rot=True):
    rp = (f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/>')
    return (f'<c:title><c:tx><c:rich><a:bodyPr{" rot=\"-5400000\" vert=\"horz\"" if rot else ""}/><a:lstStyle/><a:p><a:pPr>'
            f'<a:defRPr sz="950" b="1">{rp}</a:defRPr></a:pPr><a:r><a:rPr lang="ko-KR" sz="950" b="1">{rp}</a:rPr>'
            f'<a:t>{svg.esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')


def _pptx_dual_axis(ch, block, plot):
    """python-pptx 에는 콤보·보조축 API 가 없다 → 차트 XML 을 직접 고친다.
    오른쪽 축 계열을 첫 차트 그룹에서 빼 새 c:lineChart(축 id 2개 새로)로 옮기고,
    숨긴 c:catAx + 오른쪽 c:valAx(crosses=max)를 추가. 두 값 축 범위는 SVG 와 같은 규칙(0 기준·0선 맞춤)."""
    from lxml import etree
    from pptx.oxml.ns import qn
    C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ns = f'xmlns:c="{C}" xmlns:a="{A}"'
    nl = block["nleft"]
    spec = block["spec"]
    grp = plot._element
    pa = grp.getparent()
    sers = grp.findall(qn("c:ser"))
    moved = sers[nl:]
    if not moved:
        return
    new_sers = []
    for k, sr in enumerate(moved):
        src = block["series"][nl + k]
        col = hexcol(spec, src.get("idx", nl + k))
        parts = [etree.tostring(sr.find(qn(t))).decode() for t in ("c:idx", "c:order", "c:tx")]
        cat, val = sr.find(qn("c:cat")), sr.find(qn("c:val"))
        new_sers.append(
            f'<c:ser {ns}>' + "".join(parts) +
            f'<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill><a:round/></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>'
            + etree.tostring(cat).decode() + etree.tostring(val).decode() + '<c:smooth val="0"/></c:ser>')
        grp.remove(sr)
    ax_cat, ax_val = "50050", "50060"
    line = etree.fromstring(f'<c:lineChart {ns}><c:grouping val="standard"/><c:varyColors val="0"/>' + "".join(new_sers) +
                            f'<c:marker val="1"/><c:axId val="{ax_cat}"/><c:axId val="{ax_val}"/></c:lineChart>')
    grp.addnext(line)
    # 왼쪽 값 축 범위·제목 색
    lva = pa.find(qn("c:valAx"))
    lo, hi, _, st = block["dom"]
    lo2, hi2, _, st2 = block["dom2"]
    fixed = block.get("aligned") and block.get("neg")  # 음수가 있으면 두 축 위아래를 고정해 0선을 맞춘다

    def scale(ax, lo_, hi_, st_, zero):
        sc = ax.find(qn("c:scaling"))
        for t in ("c:max", "c:min"):
            for e in sc.findall(qn(t)):
                sc.remove(e)
        if fixed:
            sc.append(etree.fromstring(f'<c:max {ns} val="{hi_}"/>'))
        if fixed or zero:
            sc.append(etree.fromstring(f'<c:min {ns} val="{lo_ if fixed else 0}"/>'))
        if fixed:
            if ax.find(qn("c:majorUnit")) is None:  # 스키마 순서: crossAx, crosses, crossBetween 다음
                anchor = [e for e in ax if e.tag in (qn("c:crossAx"), qn("c:crosses"), qn("c:crossesAt"), qn("c:crossBetween"))][-1]
                anchor.addnext(etree.fromstring(f'<c:majorUnit {ns} val="{st_}"/>'))
    scale(lva, lo, hi, st, True)
    left_col = hexcol(spec, block["series"][0].get("idx", 0)) if nl == 1 else "52514E"
    ttl = lva.find(qn("c:title"))
    if ttl is not None and block.get("y_title"):
        lva.replace(ttl, etree.fromstring(_axis_title_xml(block["y_title"], svg.darken("#" + left_col).lstrip("#").upper() if nl == 1 else left_col).replace("<c:title>", f"<c:title {ns}>", 1)))
    rcol = hexcol(spec, block["series"][nl].get("idx", nl))
    rnum = numfmt([v for s in block["series"][nl:] for v in s["values"]], block.get("unit2", ""))
    title = _axis_title_xml(block["y2_title"], svg.darken("#" + rcol).lstrip("#").upper()) if block.get("y2_title") else ""
    cat_ax = etree.fromstring(
        f'<c:catAx {ns}><c:axId val="{ax_cat}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>'
        f'<c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="{ax_val}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
        f'<c:noMultiLvlLbl val="0"/></c:catAx>')
    val_ax = etree.fromstring(
        f'<c:valAx {ns}><c:axId val="{ax_val}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="r"/>{title.replace("<c:title>", "<c:title>", 1)}<c:numFmt formatCode="{svg.esc(rnum)}" sourceLinked="0"/>'
        f'<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="898781"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>'
        f'<c:crossAx val="{ax_cat}"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>')
    scale(val_ax, lo2, hi2, st2, block.get("zero", True))
    last_ax = [e for e in pa if e.tag in (qn("c:catAx"), qn("c:valAx"), qn("c:dateAx"), qn("c:serAx"))][-1]
    last_ax.addnext(cat_ax)
    cat_ax.addnext(val_ax)


def _set_ea_font(run_or_style_rpr, name=FONT):
    from docx.oxml.ns import qn
    rf = run_or_style_rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = run_or_style_rpr.makeelement(qn("w:rFonts"), {})
        run_or_style_rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), name)


def heat_hex(v, lo, hi):
    f = (v - lo) / ((hi - lo) or 1)
    return svg.SEQ[min(len(svg.SEQ) - 1, int(round(f * (len(svg.SEQ) - 1))))].lstrip("#").upper()


# ── PPTX ─────────────────────────────────────────────────────────────────
def build_pptx(table, specs, path, doc_title="", include_table=True):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Emu, Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    report = []

    def textbox(slide, text, x, y, w, h, size, bold=False, color="0B0B0B"):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = text
        for p in tf.paragraphs:
            p.font.size, p.font.bold, p.font.name = Pt(size), bold, FONT
            p.font.color.rgb = RGBColor.from_string(color)
        return tb

    if doc_title:
        s = prs.slides.add_slide(blank)
        textbox(s, doc_title, Inches(0.8), Inches(2.8), Inches(11.7), Inches(1.2), 34, True)
        textbox(s, f"그림 {len(specs)}개 · chart local 에서 생성 — 차트를 오른쪽 클릭 → '데이터 편집'으로 값을 고칠 수 있습니다",
                Inches(0.8), Inches(4.0), Inches(11.7), Inches(0.6), 14, False, "52514E")
    for n, spec in enumerate(specs, 1):
        blocks = blocks_for(table, spec)
        slide = prs.slides.add_slide(blank)
        title = spec.get("title") or blocks[0]["title"]
        textbox(slide, f"그림 {n}. {title}", Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.6), 20, True)
        top, avail = Inches(1.0), Inches(5.9)
        modes = []
        if blocks[0]["mode"] == "heat":
            b = blocks[0]
            _pptx_heat_table(slide, b, Inches(0.6), top, Inches(12.1), min(avail, Inches(0.42) * (len(b["series"]) + 1)))
            modes.append("table")
        elif blocks[0]["mode"] == "spark":
            _pptx_spark_rows(slide, blocks[0], Inches(0.6), top, Inches(12.1), avail)
            modes.append("native")
        elif blocks[0]["mode"] == "image":
            b = blocks[0]
            wd = tempfile.mkdtemp(prefix="chart-img-")
            try:
                png = png_for(table, spec, wd, f"fig{n}.png")
                if png:
                    pic = slide.shapes.add_picture(png, Inches(0.6), top, height=avail)
                    if pic.width > Inches(12.1):  # 넓은 그림은 폭에 맞춤
                        r_ = Inches(12.1) / pic.width
                        pic.width, pic.height = Inches(12.1), int(pic.height * r_)
                    pic.left = int((prs.slide_width - pic.width) / 2)
                    modes.append("image")
                else:
                    head, rows = data_rows(b)
                    _pptx_data_table(slide, head, rows, Inches(0.6), top, Inches(12.1), min(avail, Inches(0.36) * (len(rows) + 1)))
                    modes.append("table")
            finally:
                shutil.rmtree(wd, ignore_errors=True)
        else:
            h_each = int(avail / len(blocks))
            for i, b in enumerate(blocks):
                _pptx_chart(slide.shapes, b, Inches(0.6), top + h_each * i, Inches(12.1), h_each - Inches(0.05))
                modes.append("native")
        note = _note(spec, blocks)
        if note:
            textbox(slide, note, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4), 10.5, False, "898781")
        report.append(_rep(n, title, "pptx", modes, blocks))
    prs.core_properties.author = "chart local"
    prs.core_properties.title = doc_title or "그래프"
    prs.save(path)
    return report


def _pptx_data_table(slide, head, rows, x, y, w, h):
    """그림을 못 만들 때(rsvg-convert 없음) 넣는 편집 가능한 데이터 표"""
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    rows = rows[:30]
    tbl = slide.shapes.add_table(len(rows) + 1, len(head), x, y, w, h).table
    decs = [core.decimals_of([r[j] for r in rows if isinstance(r[j], (int, float))], 3) for j in range(len(head))]
    for j, hd in enumerate(head):
        tbl.cell(0, j).text = str(hd)
    for i, r in enumerate(rows, 1):
        for j, v in enumerate(r):
            tbl.cell(i, j).text = cell_text(v, decs[j] if isinstance(v, (int, float)) else None)
    for row in tbl.rows:
        for c in row.cells:
            for p in c.text_frame.paragraphs:
                p.font.size, p.font.name = Pt(10.5), FONT


def _pptx_heat_table(slide, b, x, y, w, h):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    rows, cols = len(b["series"]) + 1, len(b["cats"]) + 1
    tbl = slide.shapes.add_table(rows, cols, x, y, w, h).table
    vals = [v for s in b["series"] for v in s["values"] if v is not None]
    lo, hi = min(vals), max(vals)
    dec = core.decimals_of(vals)

    def put(cell, text, fill=None, bold=False, color="0B0B0B"):
        cell.text = text
        p = cell.text_frame.paragraphs[0]
        p.font.size, p.font.bold, p.font.name = Pt(11), bold, FONT
        p.font.color.rgb = RGBColor.from_string(color)
        if fill:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(fill)
    put(tbl.cell(0, 0), b.get("x_title") or "", "F0EFEC", True)
    for j, c in enumerate(b["cats"]):
        put(tbl.cell(0, j + 1), str(c), "F0EFEC", True)
    for i, s in enumerate(b["series"]):
        put(tbl.cell(i + 1, 0), s["name"], "F0EFEC", True)
        for j, v in enumerate(s["values"]):
            if v is None:
                put(tbl.cell(i + 1, j + 1), "", "FFFFFF")
            else:
                f = heat_hex(v, lo, hi)
                put(tbl.cell(i + 1, j + 1), core.fmt(v, dec).replace("−", "-"), f, False,
                    "FFFFFF" if svg.lum("#" + f) < 0.35 else "0B0B0B")


def _note(spec, blocks, fmt=None):
    k = spec["kind"]
    parts = []
    office = fmt != "hwpx"  # 오피스 차트를 어떻게 꾸몄는지에 대한 안내는 HWPX(그림)에는 뺀다
    if k == "box":
        parts.append("상자: Q1~Q3, 가운데 경계=중앙값, 수염: 최소~최대 (문서 차트는 이상값을 따로 표시하지 않음)")
    if k == "panels":
        parts.append("단위가 다른 두 계열이라 이중축 대신 위·아래 두 그래프로 나눔")
    if k == "index":
        parts.append("각 계열의 첫 값을 100 으로 맞춘 지수")
    if k == "dual" and blocks and blocks[0].get("nleft") is not None:
        b = blocks[0]
        r = b["series"][b["nleft"]:]
        parts.append(f"이중축: 왼쪽 축 {'·'.join(x['name'] for x in b['series'][:b['nleft']])}, 오른쪽 축 {'·'.join(x['name'] for x in r)}"
                     " — 두 축 눈금이 독립이라 추세 비교 시 주의")
    if k == "hist":
        parts.append("구간: 하한 이상 ~ 상한 미만 (마지막 구간은 상한 포함)")
    if k == "heatmap":
        parts.append("히트맵은 색을 칠한 표로 넣음 — 값을 고치면 색은 직접 바꿔야 함 (XLSX 는 조건부 서식으로 자동)")
    if k == "waterfall" and office:
        parts.append("폭포: 투명한 '기준' 계열 위에 증가·감소·합계를 쌓은 누적 막대 — '기준' 계열을 지우거나 칠하지 마세요")
        if blocks and any(min(st["a"], st["b"]) < 0 for st in blocks[0].get("steps") or []):
            parts.append("누적이 0 아래로 내려가는 구간은 엑셀·파워포인트에서는 맞게 보이나 LibreOffice 는 음수 누적 막대를 위로 그려 어긋남")
    if k == "funnel" and office:
        parts.append("깔때기: 투명한 '여백' 계열로 가운데를 맞춘 누적 가로 막대 (전환율은 데이터 표)")
    if k == "candlestick":
        up, dn = svg.candle_colors(spec)
        parts.append(f"캔들: 몸통 {'빨강' if up == svg.UP_COL else '첫 색'}=상승(종가 ≥ 시가), {'파랑' if dn == svg.DOWN_COL else '둘째 색'}=하락, 심지=고가~저가"
                     + (" — 주식형 차트(시가·고가·저가·종가 순서)" if office else ""))
    if k == "lollipop" and office:
        parts.append("막대사탕: 선 없는 꺾은선 표식 + 아래로 100% 오차 막대(줄기)")
    if k == "dumbbell" and office:
        parts.append("덤벨: 두 계열 표식을 최고-최저 선으로 이음")
    if k == "bullet":
        parts.append("불릿: 실적=막대, 목표=가로 표식" + (" (그림의 '목표까지 연한 바탕'은 문서 차트에서 생략)" if office else ""))
    if k == "semi_doughnut" and office:
        parts.append("반원 도넛: 아래쪽 반은 합계와 같은 투명 조각 — 조각 라벨은 % 대신 값(비율은 데이터 표)")
    if k == "sparkline":
        parts.append("스파크라인: 계열마다 눈금이 따로라 높이로 계열끼리 크기를 비교하지 마세요")
    if k == "bubble":
        parts.append(f"거품: 원 넓이 ∝ {spec.get('size') or '크기'}")
    if k == "stack_line" and office:
        parts.append(f"누적 막대와 꺾은선({spec.get('y2', '')})이 같은 값 축을 씀")
    if k in IMAGE_KINDS and office:
        parts.append(f"{core.jo(core.KINDS[k].split(' (')[0], '은', '는')} 오피스 차트에 같은 형식이 없어 그림으로 넣음 — 값은 데이터 표에서 고쳐 그림을 다시 만드세요")
    for b in blocks:
        if b.get("note"):
            parts.append(b["note"])
            break
    return " · ".join(parts)


def _rep(n, title, fmt_, modes, blocks):
    label = {"native": "네이티브 차트 (데이터 편집 가능)", "table": "색 칠한 표 (편집 가능, 차트 개체 아님)",
             "image": "그림(PNG) + 데이터 표", "hwp-native": "한글 차트 개체 (kordoc 차트)"}
    uniq = []
    for m in modes:
        if m not in uniq:
            uniq.append(m)
    if blocks and not any(b.get("mode") == "heat" for b in blocks):  # 히트맵이 아닌데 표 = 그림 변환기(rsvg-convert)가 없어 데이터 표만
        label["table"] = "데이터 표만 (그림 변환기 없음 — 편집 가능, 차트 개체 아님)"
    return {"n": n, "title": title, "format": fmt_, "modes": uniq, "label": " + ".join(label[m] for m in uniq),
            "editable": all(m in ("native", "table", "hwp-native") for m in uniq)}


# ── DOCX: python-pptx 차트 파트를 word 문서에 이식 ─────────────────────────
def build_docx(table, specs, path, doc_title="", include_table=True):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
    from lxml import etree
    from pptx import Presentation
    from pptx.util import Emu

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.2)
    for st in ("Normal", "Heading 1", "Heading 2", "Caption", "Title"):
        try:
            sty = doc.styles[st]
        except KeyError:
            continue
        sty.font.name = FONT
        _set_ea_font(sty.element.get_or_add_rPr())
    doc.styles["Normal"].font.size = Pt(10.5)
    if doc_title:
        h = doc.add_heading(doc_title, level=1)
    report = []
    width_emu = int(Cm(16.6))
    prs = Presentation()  # 차트 XML·내장 워크북 생성용 임시 프레젠테이션
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    chart_no = 0
    pkg = doc.part.package
    for n, spec in enumerate(specs, 1):
        blocks = blocks_for(table, spec)
        title = spec.get("title") or blocks[0]["title"]
        cap = doc.add_paragraph()
        r = cap.add_run(f"그림 {n}. {title}")
        r.bold = True
        r.font.size = Pt(11.5)
        modes = []

        def embed(gf, w_emu, h_emu, run):
            """python-pptx 차트 파트(chartSpace + 내장 xlsx)를 word/charts 로 옮기고 run 에 인라인 그림으로 붙임"""
            nonlocal chart_no
            cpart = gf.chart_part
            chart_no += 1
            xml = etree.tostring(cpart._element, xml_declaration=True, encoding="UTF-8", standalone=True)
            wb_blob = cpart.chart_workbook.xlsx_part.blob
            new_chart = Part(PackURI(f"/word/charts/chart{chart_no}.xml"),
                             "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", xml, pkg)
            wb_part = Part(PackURI(f"/word/embeddings/Microsoft_Excel_Worksheet{chart_no}.xlsx"),
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", wb_blob, pkg)
            wb_rid = new_chart.relate_to(wb_part, RT.PACKAGE)
            # 차트 XML 의 externalData r:id 를 새 관계 id 로
            new_chart._blob = re.sub(rb'(<c:externalData r:id=")[^"]+(")', lambda m: m.group(1) + wb_rid.encode() + m.group(2), xml)
            rid = doc.part.relate_to(new_chart, RT.CHART)
            inline = parse_xml(
                f'<wp:inline distT="0" distB="0" distL="0" distR="0" '
                f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                f'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'<wp:extent cx="{w_emu}" cy="{h_emu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
                f'<wp:docPr id="{1000 + chart_no}" name="차트 {chart_no}"/><wp:cNvGraphicFramePr/>'
                f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                f'<c:chart r:id="{rid}"/></a:graphicData></a:graphic></wp:inline>')
            drawing = run._r.makeelement(qn("w:drawing"), {})
            drawing.append(inline)
            run._r.append(drawing)

        for b in blocks:
            if b["mode"] == "heat":
                _docx_heat_table(doc, b)
                modes.append("table")
                continue
            if b["mode"] == "image":
                wd = tempfile.mkdtemp(prefix="chart-img-")
                try:
                    png = png_for(table, spec, wd, f"fig{n}.png")
                    if png:
                        doc.add_picture(png, width=Cm(16.6))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        modes.append("image")
                    else:
                        modes.append("table")  # 그림 없이 아래 데이터 표만
                finally:
                    shutil.rmtree(wd, ignore_errors=True)
                continue
            if b["mode"] == "spark":  # 계열마다 한 줄: 이름 | 작은 꺾은선(네이티브 차트) | 마지막 값·처음 대비
                t = doc.add_table(rows=len(b["series"]), cols=3)
                t.style = "Table Grid"
                sw, sh_ = int(Cm(10.2)), int(Cm(1.15))
                for j, sr in enumerate(b["series"]):
                    t.cell(j, 0).text = svg.unit_title(sr["name"], sr.get("unit", ""))
                    from pptx.chart.data import CategoryChartData
                    from pptx.enum.chart import XL_CHART_TYPE
                    cd = CategoryChartData(number_format=numfmt(sr["values"], sr.get("unit", "")))
                    cd.categories = b["cats"]
                    cd.add_series(sr["name"], sr["values"])
                    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Emu(0), Emu(0), Emu(sw), Emu(sh_), cd)
                    _spark_chart_style(gf.chart, hexcol(spec, sr.get("idx", j)))
                    embed(gf, sw, sh_, t.cell(j, 1).paragraphs[0].add_run())
                    pv = [v for v in sr["values"] if v is not None]
                    if pv:
                        ch_ = f" {'▲' if pv[-1] > pv[0] else ('▼' if pv[-1] < pv[0] else '–')} {abs((pv[-1] - pv[0]) / pv[0] * 100):.1f}%" if pv[0] else ""
                        t.cell(j, 2).text = core.fmt(pv[-1], core.decimals_of(pv), sr.get("unit", "")).replace("−", "-") + ch_
                for row in t.rows:
                    row.cells[0].width, row.cells[1].width, row.cells[2].width = Cm(3.6), Cm(10.6), Cm(2.6)
                cap2 = doc.add_paragraph()
                r2 = cap2.add_run(f"{b['cats'][0]} → {b['cats'][-1]} · 계열마다 눈금이 따로입니다" if b["cats"] else "")
                r2.font.size = Pt(9)
                modes.append("native")
                continue
            hgt = int(Cm(9.0 if len(blocks) == 1 else 6.4))
            gf = _pptx_chart(slide.shapes, b, Emu(0), Emu(0), Emu(width_emu), Emu(hgt))
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            embed(gf, width_emu, hgt, p.add_run())
            modes.append("native")
        note = _note(spec, blocks)
        if include_table or "image" in modes or (blocks[0]["mode"] == "image" and "table" in modes):
            for b in blocks:
                if b["mode"] == "heat":
                    continue
                head, rows = data_rows(b)
                if rows:
                    _docx_table(doc, head, rows, f"표 {n}{'-' + str(b['panel']) if b.get('panel') else ''}. {b['title']} — 데이터")
        if note:
            pn = doc.add_paragraph()
            rn = pn.add_run("※ " + note)
            rn.font.size = Pt(9)
            rn.font.color.rgb = RGBColor(0x89, 0x87, 0x81)
        report.append(_rep(n, title, "docx", modes, blocks))
    doc.core_properties.author = "chart local"
    doc.core_properties.title = doc_title or "그래프"
    doc.save(path)
    return report


def _docx_table(doc, head, rows, caption):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    cp = doc.add_paragraph()
    rc = cp.add_run(caption)
    rc.font.size = Pt(9.5)
    rc.bold = True
    rows = rows[:400]
    t = doc.add_table(rows=len(rows) + 1, cols=len(head))
    t.style = "Table Grid"
    decs = [core.decimals_of([r[j] for r in rows if isinstance(r[j], (int, float))], 3) for j in range(len(head))]
    for j, h in enumerate(head):
        c = t.cell(0, j)
        c.text = str(h)
        _shade(c, "F0EFEC")
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(9.5)
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.text = cell_text(v, decs[j] if isinstance(v, (int, float)) else None)
            for p in c.paragraphs:
                if isinstance(v, (int, float)):
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.font.size = Pt(9.5)
    doc.add_paragraph()


def _shade(cell, hex6):
    from docx.oxml import parse_xml
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="clear" w:color="auto" w:fill="{hex6}"/>'))


def _docx_heat_table(doc, b):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor
    vals = [v for s in b["series"] for v in s["values"] if v is not None]
    lo, hi = min(vals), max(vals)
    dec = core.decimals_of(vals)
    t = doc.add_table(rows=len(b["series"]) + 1, cols=len(b["cats"]) + 1)
    t.style = "Table Grid"
    t.cell(0, 0).text = b.get("x_title") or ""
    _shade(t.cell(0, 0), "F0EFEC")
    for j, c in enumerate(b["cats"]):
        t.cell(0, j + 1).text = str(c)
        _shade(t.cell(0, j + 1), "F0EFEC")
    for i, s in enumerate(b["series"]):
        t.cell(i + 1, 0).text = s["name"]
        _shade(t.cell(i + 1, 0), "F0EFEC")
        for j, v in enumerate(s["values"]):
            c = t.cell(i + 1, j + 1)
            if v is None:
                continue
            f = heat_hex(v, lo, hi)
            c.text = core.fmt(v, dec).replace("−", "-")
            _shade(c, f)
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.size = Pt(9.5)
                    if svg.lum("#" + f) < 0.35:
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()


# ── XLSX: 데이터 시트 + 그림별 시트(수식으로 데이터 시트 참조) + 엑셀 네이티브 차트 ─────
def _q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def build_xlsx(table, specs, path, doc_title="", include_table=True):
    import openpyxl
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "데이터"
    cols = table["columns"]
    head_fill = PatternFill("solid", fgColor="F0EFEC")
    for j, c in enumerate(cols, 1):
        cell = ws.cell(1, j, c["name"])
        cell.font = Font(bold=True, name=FONT)
        cell.fill = head_fill
        if c.get("unit"):
            ws.cell(1, j).comment = None
    nrow = len(table["rows"])
    for i, r in enumerate(table["rows"], 2):
        for j, c in enumerate(cols, 1):
            v = r[j - 1] if j - 1 < len(r) else ""
            ws.cell(i, j, _xl_val(v, c))
    for j, c in enumerate(cols, 1):
        if c["type"] == "num":
            vs = [num(r[j - 1]) for r in table["rows"] if j - 1 < len(r)]
            fmt_ = numfmt(vs, c.get("unit", ""))
            for i in range(2, nrow + 2):
                ws.cell(i, j).number_format = fmt_
        ws.column_dimensions[L(j)].width = max(10, min(28, svg.tw(c["name"], 1.1) + 4))
    # 단위는 머리글 아래 메모 대신 2행 위에 두지 않고, 열 머리글 옆 '단위' 표로 따로 적는다(수식 범위 단순화)
    if any(c.get("unit") for c in cols):
        ws.cell(nrow + 3, 1, "단위").font = Font(bold=True, name=FONT)
        for j, c in enumerate(cols, 1):
            if c.get("unit"):
                ws.cell(nrow + 4, j, c["unit"])
    ws.freeze_panes = "A2"
    last = nrow + 1
    colref = lambda name: f"{_q('데이터')}!${L(core.col_index(table, name) + 1)}$2:${L(core.col_index(table, name) + 1)}${last}"
    rowref = lambda name, i: f"={_q('데이터')}!${L(core.col_index(table, name) + 1)}${i + 2}"
    report = []
    sheet_sparks = {}
    img_dir = tempfile.mkdtemp(prefix="chart-xlimg-")
    for n, spec in enumerate(specs, 1):
        blocks = blocks_for(table, spec)
        title = spec.get("title") or blocks[0]["title"]
        sh = wb.create_sheet(f"그림{n}")
        sh["A1"] = f"그림 {n}. {title}"
        sh["A1"].font = Font(bold=True, size=13, name=FONT)
        note = _note(spec, blocks)
        sh["A2"] = ("※ " + note + " · " if note else "※ ") + "아래 표는 '데이터' 시트를 가리키는 수식입니다 — 데이터 시트 값을 고치면 차트가 바뀝니다"
        sh["A2"].font = Font(size=9, color="898781", name=FONT)
        row0 = 4
        modes = []
        anchor_row = row0
        spark_cols = []
        for bi, b in enumerate(blocks):
            used, chart, sparks = _xlsx_block(sh, table, spec, b, row0, colref, rowref, head_fill)
            if b["mode"] == "spark":
                sheet_sparks.setdefault(wb.worksheets.index(sh), []).extend(sparks)
                spark_cols.append(used["ncols"] - 1)
                modes.append("native")
            elif b["mode"] == "image":
                png = png_for(table, spec, img_dir, f"fig{n}.png", 900, 520)
                if png:
                    from openpyxl.drawing.image import Image as XLImage
                    im = XLImage(png)
                    im.width, im.height = 720, 416
                    im.anchor = f"{L(used['ncols'] + 2)}{anchor_row}"
                    sh.add_image(im)
                    modes.append("image")
                    anchor_row += 22
                else:
                    modes.append("table")
            elif chart is not None:
                chart.anchor = f"{L(used['ncols'] + 2)}{anchor_row}"
                sh.add_chart(chart)
                modes.append("native")
                anchor_row += 20
            else:
                modes.append("table")
            row0 = max(row0 + used["nrows"] + 2, row0 + 1)
            if b["mode"] == "heat":
                rng = used["range"]
                sh.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="CDE2FB", mid_type="percentile",
                                                                  mid_value=50, mid_color="3987E5", end_type="max", end_color="0D366B"))
        sh.column_dimensions["A"].width = 16
        for j in range(2, 12):
            sh.column_dimensions[L(j)].width = 12
        for j in spark_cols:
            sh.column_dimensions[L(j)].width = 30
        report.append(_rep(n, title, "xlsx", modes, blocks))
    wb.calculation.fullCalcOnLoad = True
    wb.properties.creator = "chart local"
    wb.properties.title = doc_title or "그래프"
    wb.save(path)
    shutil.rmtree(img_dir, ignore_errors=True)
    _xlsx_add_sparklines(path, sheet_sparks)
    return report


def num(v):
    return core.num(v)


def _xl_val(v, c):
    if c["type"] == "num":
        x = core.num(v)
        return x if x is not None else (None if str(v).strip() in core._EMPTY else v)
    if c["type"] == "date" and re.fullmatch(r"(19|20)\d{2}", str(v).strip()):
        return int(v)
    return v if v != "" else None


def _xlsx_block(sh, table, spec, b, r0, colref, rowref, head_fill):
    """그림 시트에 차트용 표(수식)를 쓰고 openpyxl 차트를 만든다. 반환: (크기 정보, 차트|None, 스파크라인 목록)"""
    if b["mode"] == "spark" or b.get("ctype") in ("waterfall", "funnel", "bubble"):
        x, grp = spec.get("x"), spec.get("group")
        xs = core.col_values(table, x) if x else None
        agg = "AVERAGEIFS" if spec.get("agg") == "mean" else "SUMIFS"

        def cell_ref(s, c, j):  # 스파크라인 칸: 데이터 시트를 가리키는 수식(없으면 값)
            if xs is None:
                return s["values"][j]
            if grp:
                lvl = str(s["name"]).replace('"', '""')
                return f'={agg}({colref(spec["ys"][0])},{colref(x)},"{str(c).replace(chr(34), chr(34) * 2)}",{colref(grp)},"{lvl}")'
            if xs.count(c) == 1:
                return rowref(s["name"], xs.index(c))
            if c in xs:
                return f'={agg}({colref(s["name"])},{colref(x)},"{str(c).replace(chr(34), chr(34) * 2)}")'
            return s["values"][j]
        return _xlsx_new(sh, table, spec, b, r0, colref, rowref, head_fill, cell_ref)
    used, chart = _xlsx_block_cat(sh, table, spec, b, r0, colref, rowref, head_fill)
    return used, chart, []


def _xlsx_block_cat(sh, table, spec, b, r0, colref, rowref, head_fill):
    from openpyxl.chart import AreaChart, BarChart, DoughnutChart, LineChart, PieChart, Reference, ScatterChart, Series
    from openpyxl.chart.data_source import NumDataSource, NumRef
    from openpyxl.chart.error_bar import ErrorBars
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.layout import Layout, ManualLayout
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.trendline import Trendline
    from openpyxl.drawing.line import LineProperties
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter as L

    title_font = Font(bold=True, name=FONT)
    ct = b.get("ctype") or "heat"
    x = spec.get("x")
    grp = spec.get("group")
    agg = "AVERAGEIFS" if spec.get("agg") == "mean" else "SUMIFS"
    xcol = None
    if x:
        xcol = table["columns"][core.col_index(table, x)]
    if b["mode"] == "heat" or ct not in ("scatter", "hist", "box"):
        series = b["series"]
        cats = b["cats"]
        sh.cell(r0, 1, b.get("x_title") or x or "항목").font = title_font
        for j, s in enumerate(series, 2):
            sh.cell(r0, j, s["name"]).font = title_font
            sh.cell(r0, j).fill = head_fill
        sh.cell(r0, 1).fill = head_fill
        xs = core.col_values(table, x) if x else None
        for i, c in enumerate(cats, 1):
            sh.cell(r0 + i, 1, _xl_val(c, xcol) if xcol else c)
            for j, s in enumerate(series, 2):
                ref = None
                if xs is not None:
                    if grp:
                        ref = (f"={agg}({colref(spec['ys'][0])},{colref(x)},$A{r0 + i},{colref(grp)},{L(j)}${r0})")
                    elif xs.count(c) == 1:
                        ref = rowref(s["name"], xs.index(c))
                    elif c in xs:
                        ref = f"={agg}({colref(s['name'])},{colref(x)},$A{r0 + i})"
                v = s["values"][i - 1]
                if ct == "index":  # 원값은 오른쪽 보조 열에, 지수 = 원값 / 첫 원값 × 100
                    raw = len(series) + 1 + j
                    sh.cell(r0, raw, f"{s['name']} (원값)").font = title_font
                    sh.cell(r0 + i, raw, ref if ref else None)
                    sh.cell(r0 + i, j, f'=IFERROR({L(raw)}{r0 + i}/{L(raw)}${r0 + 1}*100,"")')
                    sh.cell(r0 + i, j).number_format = "0.0"
                    continue
                sh.cell(r0 + i, j, ref if ref else v)
                sh.cell(r0 + i, j).number_format = numfmt([v], s.get("unit", b.get("unit", "")))
        used = {"nrows": len(cats) + 1, "ncols": len(series) * (2 if ct == "index" else 1) + 1,
                "range": f"B{r0 + 1}:{L(len(series) + 1)}{r0 + len(cats)}"}
        if b["mode"] in ("heat", "image"):
            return used, None
        if ct in NEW_CT:
            return used, _xlsx_new_cat(sh, spec, b, r0, len(cats), len(series))
        if ct == "dual":
            nl = b["nleft"]
            for j in range(nl + 2, len(series) + 2):
                sh.cell(r0, j, f"{series[j - 2]['name']} (오른쪽 축)")
            return used, _xl_dual(sh, spec, b, r0, len(cats), nl, len(series))
        if ct in ("pie", "doughnut"):
            ch = DoughnutChart(holeSize=55) if ct == "doughnut" else PieChart()
        elif ct in ("line", "index"):
            ch = LineChart()
        elif ct in ("area", "area_stack"):
            ch = AreaChart()
            ch.grouping = "stacked" if ct == "area_stack" else "standard"
        else:
            ch = BarChart()
            ch.type = "bar" if ct.startswith("bar") else "col"
            ch.grouping = {"col_stack": "stacked", "bar_stack": "stacked", "col_stack100": "percentStacked"}.get(ct, "clustered")
            if ch.grouping != "clustered":
                ch.overlap = 100
            ch.gapWidth = 60 if len(series) > 1 else 90
        data = Reference(sh, min_col=2, max_col=len(series) + 1, min_row=r0, max_row=r0 + len(cats))
        catref = Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + len(cats))
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(catref)
        for i, s in enumerate(ch.series):
            col = hexcol(b["spec"], series[i].get("idx", i))
            if ct in ("pie", "doughnut"):
                from openpyxl.chart.series import DataPoint
                for p in range(len(cats)):
                    pt = DataPoint(idx=p)
                    pt.graphicalProperties.solidFill = hexcol(spec, p)
                    pt.graphicalProperties.line.solidFill = "FFFFFF"
                    s.dPt.append(pt)
            elif ct in ("line", "index"):
                s.graphicalProperties.line.solidFill = col
                s.graphicalProperties.line.width = 28575
                s.marker = Marker(symbol="circle", size=7)
                s.marker.graphicalProperties = GraphicalProperties(solidFill=col)
                s.marker.graphicalProperties.line.solidFill = "FFFFFF"
                s.smooth = False
            else:
                s.graphicalProperties.solidFill = col
                s.graphicalProperties.line.solidFill = "FFFFFF" if "stack" in ct else col
        if ct in ("bar", "bar_group", "bar_stack"):
            ch.x_axis.scaling.orientation = "maxMin"
        _xl_common(ch, b, len(series), ct)
        return used, ch
    if ct == "hist":
        src = b["src"]
        heads = ["구간", "하한", "상한", "빈도"]
        for j, h in enumerate(heads, 1):
            sh.cell(r0, j, h).font = title_font
            sh.cell(r0, j).fill = head_fill
        nb = len(b["bins"])
        for i, (a, bb, cnt) in enumerate(b["bins"], 1):
            r = r0 + i
            sh.cell(r, 2, a)
            sh.cell(r, 3, bb)
            sh.cell(r, 1, f'=TEXT(B{r},"#,##0.##")&"~"&TEXT(C{r},"#,##0.##")')
            op = "<=" if i == nb else "<"
            sh.cell(r, 4, f'=COUNTIFS({colref(src)},">="&B{r},{colref(src)},"{op}"&C{r})')
        ch = BarChart()
        ch.type, ch.gapWidth = "col", 6
        ch.add_data(Reference(sh, min_col=4, min_row=r0, max_row=r0 + nb), titles_from_data=True)
        ch.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + nb))
        ch.series[0].graphicalProperties.solidFill = hexcol(spec, 0)
        ch.series[0].graphicalProperties.line.solidFill = "FFFFFF"
        _xl_common(ch, b, 1, ct)
        return {"nrows": nb + 1, "ncols": 4}, ch
    if ct == "box":
        heads = [b.get("x_title") or "계열", "최소", "Q1", "중앙값", "Q3", "최대", "", "막대:Q1", "막대:중앙값−Q1", "막대:Q3−중앙값", "수염 아래", "수염 위"]
        for j, h in enumerate(heads, 1):
            if h:
                sh.cell(r0, j, h).font = title_font
                sh.cell(r0, j).fill = head_fill
        ys = spec["ys"]
        for i, name in enumerate(b["cats"], 1):
            r = r0 + i
            sh.cell(r, 1, name)
            if grp:
                rng = f'IF({colref(grp)}=$A{r},{colref(ys[0])},"")'
                fs = [f"=MIN({rng})", f"=QUARTILE.INC({rng},1)", f"=MEDIAN({rng})", f"=QUARTILE.INC({rng},3)", f"=MAX({rng})"]
            else:
                rng = colref(name)
                fs = [f"=MIN({rng})", f"=QUARTILE.INC({rng},1)", f"=MEDIAN({rng})", f"=QUARTILE.INC({rng},3)", f"=MAX({rng})"]
            for j, f in enumerate(fs, 2):
                sh.cell(r, j, f)
            sh.cell(r, 8, f"=C{r}")
            sh.cell(r, 9, f"=D{r}-C{r}")
            sh.cell(r, 10, f"=E{r}-D{r}")
            sh.cell(r, 11, f"=C{r}-B{r}")
            sh.cell(r, 12, f"=F{r}-E{r}")
        if grp:  # 배열 수식(IF 범위) — 엑셀 365 는 그대로, 이전 버전은 CSE 로 저장
            from openpyxl.worksheet.formula import ArrayFormula
            for i in range(1, len(b["cats"]) + 1):
                r = r0 + i
                for j in range(2, 7):
                    f = sh.cell(r, j).value
                    sh.cell(r, j).value = ArrayFormula(f"{L(j)}{r}", f)
        nb = len(b["cats"])
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap, ch.gapWidth = "col", "stacked", 100, 90
        ch.add_data(Reference(sh, min_col=8, max_col=10, min_row=r0, max_row=r0 + nb), titles_from_data=True)
        ch.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + nb))
        s0, s1, s2 = ch.series
        s0.graphicalProperties.noFill = True
        s0.graphicalProperties.line.noFill = True
        for s, f in ((s1, "B7D3F6"), (s2, "86B6EF")):
            s.graphicalProperties.solidFill = f
            s.graphicalProperties.line.solidFill = "2A78D6"
        ref = lambda c: NumDataSource(numRef=NumRef(f=f"{_q(sh.title)}!${c}${r0 + 1}:${c}${r0 + nb}"))
        s0.errBars = ErrorBars(errBarType="minus", errValType="cust", noEndCap=False, minus=ref("K"),
                               spPr=GraphicalProperties(ln=LineProperties(solidFill="52514E", w=15875)))
        s2.errBars = ErrorBars(errBarType="plus", errValType="cust", noEndCap=False, plus=ref("L"),
                               spPr=GraphicalProperties(ln=LineProperties(solidFill="52514E", w=15875)))
        _xl_common(ch, b, 1, ct)
        ch.legend = None
        return {"nrows": nb + 1, "ncols": 12}, ch
    # scatter: 계열마다 (X, Y) 열 쌍 — 데이터 시트 행을 가리키는 수식
    ch = ScatterChart()
    ch.style = 13
    xname, yname = spec["x"], spec["ys"][0]
    gv = core.col_values(table, grp) if grp else None
    xs_all = core.col_nums(table, xname)
    col = 1
    maxn = 0
    for si, s in enumerate(b["series"]):
        sh.cell(r0, col, f"{s['name']} · {xname}" if grp else xname).font = title_font
        sh.cell(r0, col + 1, s["name"] if grp else yname).font = title_font
        rows = [i for i, xv in enumerate(xs_all) if xv is not None and (gv is None or gv[i] == s["name"])
                and core.col_nums(table, yname)[i] is not None] if True else []
        for k, i in enumerate(rows, 1):
            sh.cell(r0 + k, col, rowref(xname, i))
            sh.cell(r0 + k, col + 1, rowref(yname, i))
        n = len(rows)
        maxn = max(maxn, n)
        xr = Reference(sh, min_col=col, min_row=r0 + 1, max_row=r0 + n)
        yr = Reference(sh, min_col=col + 1, min_row=r0, max_row=r0 + n)
        se = Series(yr, xr, title_from_data=True)
        c = hexcol(spec, s.get("idx", si))
        se.marker = Marker(symbol="circle", size=7)
        se.marker.graphicalProperties = GraphicalProperties(solidFill=c)
        se.marker.graphicalProperties.line.solidFill = "FFFFFF"
        se.graphicalProperties.line.noFill = True
        if b.get("trend"):
            se.trendline = Trendline(trendlineType="linear", dispRSqr=True, dispEq=False)
        ch.series.append(se)
        col += 3
    _xl_common(ch, b, len(b["series"]), ct)
    return {"nrows": maxn + 1, "ncols": col - 1}, ch


def _xl_title(text, hex6):
    from openpyxl.chart.text import RichText, Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
    cp = CharacterProperties(sz=1000, b=True, solidFill=hex6)
    return Title(tx=Text(rich=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])])),
                 overlay=False)


def _xl_dual(sh, spec, b, r0, ncat, nl, nser):
    """엑셀 보조축 차트: 왼쪽 축 막대(또는 선) + 오른쪽 축 꺾은선(y_axis.axId=200, crosses=max)"""
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    shape = b.get("shape") or "bar_line"
    c1 = LineChart() if shape == "line_line" else BarChart()
    if shape != "line_line":
        c1.type, c1.grouping, c1.gapWidth = "col", "clustered", 60 if nl > 1 else 90
    c1.add_data(Reference(sh, min_col=2, max_col=nl + 1, min_row=r0, max_row=r0 + ncat), titles_from_data=True)
    c1.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + ncat))
    c2 = LineChart()
    c2.add_data(Reference(sh, min_col=nl + 2, max_col=nser + 1, min_row=r0, max_row=r0 + ncat), titles_from_data=True)
    c2.set_categories(Reference(sh, min_col=1, min_row=r0 + 1, max_row=r0 + ncat))

    def style_line(s, col):
        s.graphicalProperties.line.solidFill = col
        s.graphicalProperties.line.width = 28575
        s.marker = Marker(symbol="circle", size=7)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=col)
        s.marker.graphicalProperties.line.solidFill = "FFFFFF"
        s.smooth = False
    for i, s in enumerate(c1.series):
        col = hexcol(spec, b["series"][i].get("idx", i))
        if shape == "line_line":
            style_line(s, col)
        else:
            s.graphicalProperties.solidFill = col
            s.graphicalProperties.line.solidFill = col
    for i, s in enumerate(c2.series):
        style_line(s, hexcol(spec, b["series"][nl + i].get("idx", nl + i)))
    c1.title = b["title"]
    c1.width, c1.height = 18, 10
    c1.x_axis.delete = False
    c1.y_axis.delete = False
    c2.y_axis.delete = False
    c2.y_axis.axId = 200
    c2.y_axis.crosses = "max"  # 오른쪽에 붙임
    c2.y_axis.axPos = "r"
    c2.y_axis.majorGridlines = None
    lcol = hexcol(spec, b["series"][0].get("idx", 0))
    rcol = hexcol(spec, b["series"][nl].get("idx", nl))
    if b.get("y_title"):
        c1.y_axis.title = _xl_title(b["y_title"], svg.darken("#" + lcol).lstrip("#").upper() if nl == 1 else "52514E")
    if b.get("y2_title"):
        c2.y_axis.title = _xl_title(b["y2_title"], svg.darken("#" + rcol).lstrip("#").upper())
    if b.get("x_title"):
        c1.x_axis.title = b["x_title"]
    c1.y_axis.number_format = numfmt([v for s in b["series"][:nl] for v in s["values"]], b.get("unit", ""))
    c2.y_axis.number_format = numfmt([v for s in b["series"][nl:] for v in s["values"]], b.get("unit2", ""))
    fixed = b.get("aligned") and b.get("neg")
    for ax, (lo, hi, _, st), zero in ((c1.y_axis, b["dom"], True), (c2.y_axis, b["dom2"], b.get("zero", True))):
        if fixed:
            ax.scaling.min, ax.scaling.max, ax.majorUnit = lo, hi, st
        elif zero:
            ax.scaling.min = 0
    if b["legend"] == "none":
        c1.legend = None
    else:
        c1.legend.position = {"top": "t", "bottom": "b", "right": "r"}.get(b["legend"], "t")
    c1 += c2
    return c1


def _xl_common(ch, b, nser, ct):
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
    ch.title = b["title"]
    ch.width, ch.height = 18, 10 if not b.get("panel") else 8
    if ct not in ("pie", "doughnut"):
        try:
            ch.x_axis.delete = False
            ch.y_axis.delete = False
        except AttributeError:
            pass
        if b.get("y_title"):
            ch.y_axis.title = b["y_title"]
        if b.get("x_title"):
            ch.x_axis.title = b["x_title"]
        ch.y_axis.majorGridlines = ch.y_axis.majorGridlines or __import__("openpyxl.chart.axis", fromlist=["ChartLines"]).ChartLines()
        unit = b.get("unit", "")
        if ct not in ("col_stack100", "box", "hist", "scatter"):
            vals = [v for s in b["series"] for v in (s.get("values") or []) if v is not None]
            ch.y_axis.number_format = numfmt(vals, unit)
            ch.y_axis.numFmt.sourceLinked = False if ch.y_axis.numFmt is not None else None
    want_legend = (nser >= 2 or ct in ("pie", "doughnut")) and b["legend"] != "none"
    if not want_legend:
        ch.legend = None
    else:
        ch.legend.position = {"top": "t", "bottom": "b", "right": "r"}.get(b["legend"], "t") if ct not in ("pie", "doughnut") else "r"
    lab = b["labels"]
    if ct in ("pie", "doughnut") and lab != "none":
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showPercent = True
        ch.dataLabels.showVal = False
    elif ct not in ("box", "scatter", "area_stack", "bubble") and lab != "none" and (lab == "all" or (nser == 1 and len(b.get("cats") or []) <= 16)):
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showVal = True
        for k in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
            setattr(ch.dataLabels, k, False)


# ── HWPX (kordoc) ───────────────────────────────────────────────────────────
def kordoc_cli():
    for p in (os.environ.get("KORDOC_CLI"), os.path.join(HERE, "node_modules", "kordoc", "dist", "cli.js"),
              os.path.join(HERE, "..", "kordoc-local", "node_modules", "kordoc", "dist", "cli.js")):
        if p and os.path.exists(p):
            return os.path.abspath(p)
    return None


def hwpx_available():
    return bool(kordoc_cli() and shutil.which("node"))


HWP_TYPES = {"col": "column", "col_group": "column", "col_stack": "column_stacked", "col_stack100": "column_stacked",
             "bar": "bar", "bar_group": "bar", "bar_stack": "bar_stacked", "line": "line", "index": "line",
             "area": "area", "area_stack": "area_stacked", "pie": "pie", "doughnut": "doughnut", "scatter": "scatter",
             "hist": "column", "radar": "radar"}


def _fence_txt(s):
    return re.sub(r"[,:：\n]", lambda m: {",": "·", ":": "∶", "：": "∶", "\n": " "}[m.group()], str(s)).strip() or "-"


def svg_to_png(svg_text, path, scale=2):
    exe = shutil.which("rsvg-convert")
    if not exe:
        return False
    with tempfile.NamedTemporaryFile("w", suffix=".svg", delete=False, encoding="utf-8") as f:
        f.write(svg_text)
        tmp = f.name
    try:
        r = subprocess.run([exe, "-z", str(scale), "-o", path, tmp], capture_output=True, timeout=60)
        return r.returncode == 0 and os.path.exists(path)
    finally:
        os.unlink(tmp)


def _md_table(head, rows, decs=None):
    esc_ = lambda s: str(s).replace("|", "∣").replace("\n", " ")
    decs = decs or [core.decimals_of([r[j] for r in rows if isinstance(r[j], (int, float))], 3) for j in range(len(head))]
    out = ["| " + " | ".join(esc_(h) for h in head) + " |", "|" + "|".join("---:" if decs[j] is not None and any(isinstance(r[j], (int, float)) for r in rows) else "---" for j in range(len(head))) + "|"]
    for r in rows[:300]:
        out.append("| " + " | ".join(esc_(cell_text(v, decs[j] if isinstance(v, (int, float)) else None)) for j, v in enumerate(r)) + " |")
    return "\n".join(out)


def build_hwpx(table, specs, path, doc_title="", include_table=True, workdir=None):
    cli = kordoc_cli()
    if not cli or not shutil.which("node"):
        raise RuntimeError("HWPX 내보내기에는 Node 와 kordoc 이 필요합니다 (kordoc-local 을 설치하거나 KORDOC_CLI 지정)")
    wd = workdir or tempfile.mkdtemp(prefix="chart-hwpx-")
    md = [f"# {doc_title or '그래프'}", ""]
    report = []
    for n, spec in enumerate(specs, 1):
        blocks = blocks_for(table, spec)
        title = spec.get("title") or blocks[0]["title"]
        md += [f"**그림 {n}. {title}**", ""]
        modes = []
        if spec["kind"] in HWP_IMAGE_KINDS:  # 한글 차트(kordoc)에 없는 형식 → 그림 한 장 + 편집용 데이터 표
            png = os.path.join(wd, f"fig{n}.png")
            if svg_to_png(svg.render(table, spec, 720, 440), png, 2):
                with open(png, "rb") as f:
                    b64 = __import__("base64").b64encode(f.read()).decode()
                md += [f"![그림 {n}](data:image/png;base64,{b64})", ""]
                modes.append("image")
            else:
                modes.append("table")
            for b in blocks:
                head, rows = data_rows(b)
                if rows:
                    md += [f"표 {n}{'-' + str(b['panel']) if b.get('panel') else ''}. {b['title']} — 데이터", "", _md_table(head, rows), ""]
            blocks_iter = []
        else:
            blocks_iter = blocks
        for b in blocks_iter:
            ct = b.get("ctype")
            if b["mode"] == "chart" and ct in HWP_TYPES:
                md += ["```chart", f"type: {HWP_TYPES[ct]}", f"title: {_fence_txt(b['title'])}"]
                if ct == "scatter":
                    s0 = b["series"][0]  # kordoc 산점도는 x 하나를 공유 → 계열이 여럿이면 첫 계열만 차트, 나머지는 표
                    md.append("cat: " + ", ".join(repr(round(x, 6)).rstrip("0").rstrip(".") if isinstance(x, float) else str(x) for x, _ in s0["points"]))
                    md.append(f"{_fence_txt(s0['name'] if len(b['series']) > 1 else spec['ys'][0])}: " + ", ".join(_numtxt(y) for _, y in s0["points"]))
                    md.append(f"colors: #{hexcol(spec, s0.get('idx', 0))}")
                else:
                    sers = b["series"]
                    vals_by = [s["values"] for s in sers]
                    if ct == "col_stack100":
                        tots = [sum(abs(v or 0) for v in col) for col in zip(*vals_by)]
                        vals_by = [[(100 * (v or 0) / t if t else 0) for v, t in zip(vs, tots)] for vs in vals_by]
                    md.append("cat: " + ", ".join(_fence_txt(c) for c in b["cats"]))
                    for s, vs in zip(sers, vals_by):
                        md.append(f"{_fence_txt(s['name'])}: " + ", ".join(_numtxt(v) for v in vs))
                    if ct in ("pie", "doughnut"):
                        md.append("colors: " + ", ".join("#" + hexcol(spec, i) for i in range(len(b["cats"]))))
                    else:
                        md.append("colors: " + ", ".join("#" + hexcol(spec, s.get("idx", i)) for i, s in enumerate(sers)))
                md += ["size: 160x95" if not b.get("panel") else "size: 160x70", "```", ""]
                modes.append("hwp-native")
            else:
                png = os.path.join(wd, f"fig{n}{'_' + str(b.get('panel')) if b.get('panel') else ''}.png")
                if svg_to_png(svg.render(table, spec, 720, 440), png, 2):
                    # kordoc CLI 는 파일 경로 그림을 자리표시자로 넣는다 → data: URI 로 실제 바이트를 넘김
                    with open(png, "rb") as f:
                        b64 = __import__("base64").b64encode(f.read()).decode()
                    md += [f"![그림 {n}](data:image/png;base64,{b64})", ""]
                    modes.append("image")
                elif b["mode"] != "heat":
                    modes.append("table")
            if include_table or b["mode"] == "heat" or "image" in modes or "table" in modes:
                head, rows = data_rows(b)
                if rows:
                    md += [f"표 {n}{'-' + str(b['panel']) if b.get('panel') else ''}. {b['title']} — 데이터", "", _md_table(head, rows), ""]
                    if b["mode"] == "heat" and "image" not in modes:
                        modes.append("table")
        note = _note(spec, blocks, "hwpx")
        if spec["kind"] == "dual":
            note = (note + " · " if note else "") + "한글 차트(kordoc)는 보조 축을 지원하지 않아 그림으로 넣음 — 값은 아래 표에서 고쳐 그림을 다시 만드세요"
        if spec["kind"] in HWP_IMAGE_KINDS:
            note = (note + " · " if note else "") + f"한글 차트(kordoc)에 {core.KINDS[spec['kind']].split(' (')[0]} 형식이 없어 그림으로 넣음 — 값은 아래 표에서 고쳐 그림을 다시 만드세요"
        if spec["kind"] == "scatter" and (len(blocks[0]["series"]) > 1 or spec.get("trend")):
            note = (note + " · " if note else "") + ("한글 차트에는 첫 계열만 넣음 — 나머지는 표 참고" if len(blocks[0]["series"]) > 1 else "") + \
                   (" 추세선은 한글 차트에서 '차트 → 추세선'으로 추가" if spec.get("trend") else "")
        if note:
            md += [f"※ {note.strip(' ·')}", ""]
        report.append(_rep(n, title, "hwpx", modes, blocks))
    path = os.path.abspath(path)
    mdp = os.path.join(wd, "doc.md")
    with open(mdp, "w", encoding="utf-8") as f:  # 줄표는 kordoc 표기법 검사에 걸린다 → 가운뎃점
        f.write("\n".join(md).replace(" — ", " · ").replace("—", "·"))
    r = subprocess.run(["node", cli, "generate", mdp, "-o", path, "--preset", "보고서", "--no-cover", "--no-toc",
                        "--no-page-numbers", "--summary", "그래프와 원자료를 정리함"], capture_output=True, text=True, cwd=wd, timeout=300)
    if r.returncode != 0 or not os.path.exists(path):
        raise RuntimeError("kordoc generate 실패: " + (r.stdout + r.stderr)[-600:])
    v = subprocess.run(["node", cli, "validate", path], capture_output=True, text=True, cwd=wd, timeout=120)
    for x in report:
        x["validate"] = v.returncode == 0
    if not workdir:
        shutil.rmtree(wd, ignore_errors=True)
    return report


def _numtxt(v):
    if v is None:
        return "0"
    if float(v).is_integer():
        return str(int(v))
    return repr(round(float(v), 6))


BUILDERS = {"xlsx": build_xlsx, "pptx": build_pptx, "docx": build_docx, "hwpx": build_hwpx}


# ── 검증: 만든 파일을 다시 열어 차트·데이터가 살아 있는지 확인 ─────────────────
def verify(path):
    """형식별 구조 검사. 반환: {ok, charts, tables, images, checks:[문자열], errors:[문자열]}"""
    import zipfile
    from lxml import etree
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    res = {"format": ext, "charts": 0, "tables": 0, "images": 0, "checks": [], "errors": []}
    C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    try:
        z = zipfile.ZipFile(path)
        bad = z.testzip()
        if bad:
            res["errors"].append(f"zip 손상: {bad}")
        names = set(z.namelist())
        ctypes = z.read("[Content_Types].xml").decode() if "[Content_Types].xml" in names else ""
        res["images"] = len([n for n in names if re.match(r"(ppt|word|xl)/media/.+\.(png|jpe?g|gif|emf)$", n)])
        if ext in ("pptx", "docx"):
            import openpyxl
            prefix = "ppt/charts/" if ext == "pptx" else "word/charts/"
            charts = sorted(n for n in names if n.startswith(prefix) and n.endswith(".xml") and "/_rels/" not in n)
            for cn in charts:
                if f'PartName="/{cn}"' not in ctypes:
                    res["errors"].append(f"{cn}: [Content_Types] 등록 없음")
                xml = z.read(cn)
                root_ = etree.fromstring(xml)
                pa_ = root_.find(f"{C}chart").find(f"{C}plotArea")
                plot_els = [e for e in pa_ if e.tag.endswith("Chart")]  # 주식형(stockChart)은 python-pptx 가 못 읽어 XML 로 직접
                sers = [sr for pe in plot_els for sr in pe.findall(f"{C}ser")]
                vals = [[float(v.text) for tag in ("val", "yVal", "bubbleSize") for v in sr.findall(f"{C}{tag}//{C}pt/{C}v")] for sr in sers]
                nvax = len(pa_.findall(f"{C}valAx"))
                ax_ids = {a.find(f"{C}axId").get("val") for a in pa_ if a.tag.endswith("Ax") and a.find(f"{C}axId") is not None}
                for pe in plot_els:
                    if not {e.get("val") for e in pe.findall(f"{C}axId")} <= ax_ids:
                        res["errors"].append(f"{cn}: {pe.tag.split('}')[1]} 이(가) 없는 축을 가리킴")
                plot_names = [pe.tag.split("}")[1].replace("Chart", "").capitalize() for pe in plot_els]
                relp = cn.replace("charts/", "charts/_rels/") + ".rels"
                rels = etree.fromstring(z.read(relp)) if relp in names else None
                ext_el = etree.fromstring(xml).find(f"{C}externalData")
                rid = ext_el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") if ext_el is not None else None
                tgt = None
                if rels is not None and rid:
                    for r in rels:
                        if r.get("Id") == rid:
                            tgt = os.path.normpath(os.path.join(os.path.dirname(cn), r.get("Target"))).replace("\\", "/")
                if not tgt or tgt not in names:
                    res["errors"].append(f"{cn}: 내장 워크북 연결 없음 (데이터 편집 불가)")
                else:
                    wb = openpyxl.load_workbook(io.BytesIO(z.read(tgt)), data_only=True)
                    ws = wb.worksheets[0]
                    nums = [c for row in ws.iter_rows(values_only=True) for c in row if isinstance(c, (int, float))]
                    flat = [v for vs in vals for v in vs if v is not None]
                    if flat and not all(any(abs(v - n) < 1e-6 * max(1, abs(v)) for n in nums) for v in flat[:50]):
                        res["errors"].append(f"{cn}: 내장 워크북 값이 차트 캐시와 다름")
                res["checks"].append(f"{os.path.basename(cn)}: {'+'.join(plot_names)} · 값 축 {nvax} · 계열 {len(sers)} · 값 {sum(len(v) for v in vals)}개 · 내장 xlsx {'OK' if tgt in names else '없음'}")
            res["charts"] = len(charts)
            if ext == "docx":
                doc = z.read("word/document.xml").decode()
                res["tables"] = doc.count("<w:tbl>")
                drefs = len(re.findall(r'<c:chart [^>]*r:id=', doc))
                if drefs != len(charts):
                    res["errors"].append(f"본문 차트 참조 {drefs}개 ≠ 차트 파트 {len(charts)}개")
            else:
                res["tables"] = sum(z.read(n).decode().count("<a:tbl>") for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
        elif ext == "xlsx":
            import openpyxl
            charts = sorted(n for n in names if n.startswith("xl/charts/chart") and n.endswith(".xml"))
            wb = openpyxl.load_workbook(path)
            sheets = set(wb.sheetnames)
            for cn in charts:
                root = etree.fromstring(z.read(cn))
                fs = [f.text for f in root.iter(f"{C}f")]
                for f in fs:
                    m = re.match(r"^'?(.*?)'?!", f or "")
                    if not m or m.group(1).replace("''", "'") not in sheets:
                        res["errors"].append(f"{cn}: 참조 시트 없음 {f}")
                pa_ = root.find(f"{C}chart").find(f"{C}plotArea")
                kinds = [el.tag.split("}")[1] for el in pa_ if el.tag.endswith("Chart")]
                nvax = len(root.findall(f".//{C}valAx"))
                ax_ids = {a.find(f"{C}axId").get("val") for a in pa_ if a.tag.endswith("Ax") and a.find(f"{C}axId") is not None}
                for el in pa_:
                    if el.tag.endswith("Chart") and not {e.get("val") for e in el.findall(f"{C}axId")} <= ax_ids:
                        res["errors"].append(f"{cn}: {el.tag.split('}')[1]} 이(가) 없는 축을 가리킴 (값 축 {nvax}개)")
                res["checks"].append(f"{os.path.basename(cn)}: {'+'.join(kinds)} · 값 축 {nvax} · 범위 참조 {len(fs)}개 → {', '.join(sorted({re.match(r'^.?(.*?).?!', f).group(1) for f in fs if f}))}")
            res["charts"] = len(charts)
            # 수식이 가리키는 데이터 시트 셀이 실제 값인지 (직접 참조만)
            data = wb["데이터"] if "데이터" in sheets else None
            nref = 0
            for ws in wb.worksheets[1:]:
                for row in ws.iter_rows():
                    for c in row:
                        v = c.value
                        if isinstance(v, str) and (m := re.fullmatch(r"='데이터'!\$([A-Z]+)\$(\d+)", v)):
                            nref += 1
                            if data[f"{m.group(1)}{m.group(2)}"].value is None:
                                res["errors"].append(f"{ws.title}!{c.coordinate}: 빈 셀 참조 {v}")
            res["checks"].append(f"데이터 시트 직접 참조 수식 {nref}개 확인, 조건부 서식 {sum(len(ws.conditional_formatting) for ws in wb.worksheets)}개")
            nspark = 0
            for n_ in names:
                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n_):
                    sx = z.read(n_).decode("utf-8")
                    for f in re.findall(r"<x14:sparkline><xm:f>(.*?)</xm:f><xm:sqref>(.*?)</xm:sqref>", sx):
                        nspark += 1
                        m = re.match(r"^'?(.*?)'?!", f[0].replace("&amp;", "&"))
                        if not m or m.group(1).replace("''", "'") not in sheets:
                            res["errors"].append(f"{n_}: 스파크라인 참조 시트 없음 {f[0]}")
            if nspark:
                res["checks"].append(f"엑셀 스파크라인 {nspark}개 (x14:sparklineGroups) — 범위 시트 확인")
            res["sparklines"] = nspark
            res["tables"] = 1
        elif ext == "hwpx":
            hpf = z.read("Contents/content.hpf").decode()
            charts = sorted(n for n in names if n.startswith("Chart/") and n.endswith(".xml"))
            sec = "".join(z.read(n).decode() for n in names if n.startswith("Contents/section"))
            for cn in charts:
                if f'href="{cn}"' not in hpf:
                    res["errors"].append(f"{cn}: content.hpf 등록 없음")
                if f'chartIDRef="{cn}"' not in sec:
                    res["errors"].append(f"{cn}: 본문에서 참조 안 함")
                root = etree.fromstring(z.read(cn))
                nser = len(list(root.iter(f"{C}ser")))
                npt = len(list(root.iter(f"{C}pt")))
                res["checks"].append(f"{cn}: OOXML chartSpace · 계열 {nser} · 캐시 점 {npt}")
            res["charts"] = len(charts)
            res["tables"] = sec.count("<hp:tbl ")
            res["images"] = len([n for n in names if n.startswith("BinData/") and not n.endswith("/")])
            cli = kordoc_cli()
            if cli and shutil.which("node"):
                v = subprocess.run(["node", cli, "validate", os.path.abspath(path)], capture_output=True, text=True, timeout=120)
                (res["checks"] if v.returncode == 0 else res["errors"]).append("kordoc validate: " + (v.stdout + v.stderr).strip().splitlines()[-1][:160])
    except Exception as e:  # 열 수 없는 파일
        res["errors"].append(f"{type(e).__name__}: {e}")
    res["ok"] = not res["errors"]
    return res
