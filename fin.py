"""chart local — 주가 차트 지표 (표준 라이브러리만)

캔들스틱·OHLC 막대 명세의 선택 항목 → 지표 값. 계산 규칙(문서·화면·내보내기 모두 같음):
- 이동평균(SMA): 창 w 의 단순 평균. 앞 w−1 개는 빈 값. 창 안에 빈 값이 있으면 그 점도 빈 값. 창이 데이터보다 길면 뺀다
- 볼린저 밴드: 중심 = SMA(n), 위·아래 = 중심 ± k × 표준편차. 표준편차는 모집단 표준편차(n 으로 나눔, 엑셀 STDEV.P) —
  흔한 차트 프로그램(TradingView·TA-Lib 기본)과 같음
- EMA: α = 2/(n+1). 첫 값은 앞 n 개의 단순 평균(SMA 로 시작, TA-Lib 과 같은 방식), 이후 EMA = 이전 + α × (값 − 이전)
- MACD(12, 26, 9): MACD 선 = EMA12 − EMA26 (26번째 값부터), 시그널 = MACD 선의 EMA9 (MACD 선 9개가 모인 뒤부터),
  히스토그램 = MACD 선 − 시그널
- RSI(14): 와일더(Wilder) 평활. 첫 평균 상승·하락폭 = 앞 14개 변화의 단순 평균, 이후 평균 = (이전 × 13 + 이번) / 14.
  RSI = 100 − 100 / (1 + 평균 상승 / 평균 하락). 평균 하락이 0 이면 100 (둘 다 0 이면 50)
- 매물대(거래량 프로파일): 전체 저가~고가를 같은 폭 구간(기본 24개)으로 나누고, 기간마다 거래량을 그 기간의
  저가~고가 범위에 고르게 퍼뜨려 겹치는 길이만큼 구간에 나눈다(고가 = 저가인 기간은 그 가격 구간에 전부).
  POC = 거래량이 가장 많은 구간. 가치 영역(기본 70%) = POC 에서 시작해 위·아래 이웃 중 거래량이 많은 쪽을 하나씩 붙여
  전체의 70% 이상이 될 때까지
"""
import math

MA_DEFAULT = [5, 20, 60, 120]
MA_MAX = 250
PANELS = ("volume", "macd", "rsi")
PANEL_KO = {"volume": "거래량", "macd": "MACD", "rsi": "RSI"}
FIN_KINDS = ("candlestick", "ohlc")
FIN_KEYS = ("ma", "bb", "panels", "vprofile", "colors_updown", "macd", "rsi")
UPDOWN = ("kr", "intl")


# ── 지표 ────────────────────────────────────────────────────────────────
def sma(vals, n):
    out = [None] * len(vals)
    if n < 1:
        return out
    for i in range(n - 1, len(vals)):
        w = vals[i - n + 1:i + 1]
        if any(v is None for v in w):
            continue
        out[i] = sum(w) / n
    return out


def rolling_std(vals, n):
    """모집단 표준편차(n 으로 나눔)"""
    out = [None] * len(vals)
    for i in range(n - 1, len(vals)):
        w = vals[i - n + 1:i + 1]
        if n < 1 or any(v is None for v in w):
            continue
        m = sum(w) / n
        out[i] = math.sqrt(sum((v - m) ** 2 for v in w) / n)
    return out


def bollinger(close, n=20, k=2.0):
    mid, sd = sma(close, n), rolling_std(close, n)
    up = [None if m is None else m + k * s for m, s in zip(mid, sd)]
    lo = [None if m is None else m - k * s for m, s in zip(mid, sd)]
    return mid, up, lo


def ema(vals, n):
    """SMA 로 시작하는 EMA. 앞쪽 빈 값은 건너뛰고 처음으로 n 개가 연달아 있는 곳에서 시작. 시작 뒤 빈 값은 빈 값(이전 EMA 유지)"""
    out = [None] * len(vals)
    a = 2 / (n + 1)
    run, prev = 0, None
    for i, v in enumerate(vals):
        if prev is None:
            run = run + 1 if v is not None else 0
            if run >= n:
                prev = sum(vals[i - n + 1:i + 1]) / n
                out[i] = prev
            continue
        if v is None:
            continue
        prev = prev + a * (v - prev)
        out[i] = prev
    return out


def macd(close, fast=12, slow=26, signal=9):
    ef, es = ema(close, fast), ema(close, slow)
    line = [None if a is None or b is None else a - b for a, b in zip(ef, es)]
    sig = ema(line, signal)
    hist = [None if a is None or b is None else a - b for a, b in zip(line, sig)]
    return line, sig, hist


def rsi(close, n=14):
    out = [None] * len(close)
    if len(close) <= n or any(v is None for v in close[:n + 1]):
        return out
    ch = [None] + [None if a is None or b is None else b - a for a, b in zip(close, close[1:])]
    ag = sum(max(c, 0) for c in ch[1:n + 1]) / n
    al = sum(max(-c, 0) for c in ch[1:n + 1]) / n

    def val(g, l_):
        if l_ == 0:
            return 100.0 if g > 0 else 50.0
        return 100 - 100 / (1 + g / l_)
    out[n] = val(ag, al)
    for i in range(n + 1, len(close)):
        c = ch[i]
        if c is None:
            continue
        ag = (ag * (n - 1) + max(c, 0)) / n
        al = (al * (n - 1) + max(-c, 0)) / n
        out[i] = val(ag, al)
    return out


def volume_profile(highs, lows, vols, bins=24, area=0.70):
    """반환: {edges:[구간 경계 bins+1개], vols:[구간 거래량], poc: 번호, va:(아래 번호, 위 번호), total} 또는 None"""
    pts = [(h, l, v) for h, l, v in zip(highs, lows, vols) if h is not None and l is not None and v is not None and v >= 0]
    if not pts:
        return None
    lo = min(min(h, l) for h, l, _ in pts)
    hi = max(max(h, l) for h, l, _ in pts)
    bins = max(4, min(100, int(bins or 24)))
    if hi <= lo:
        hi = lo + 1
    wd = (hi - lo) / bins
    edges = [lo + wd * i for i in range(bins)] + [hi]
    acc = [0.0] * bins
    for h, l, v in pts:
        a, b = min(h, l), max(h, l)
        if b - a <= 0:
            acc[min(bins - 1, int((a - lo) / wd))] += v
            continue
        i0, i1 = max(0, int((a - lo) / wd)), min(bins - 1, int((b - lo) / wd))
        for i in range(i0, i1 + 1):
            ov = min(b, edges[i + 1]) - max(a, edges[i])
            if ov > 0:
                acc[i] += v * ov / (b - a)
    total = sum(acc)
    poc = max(range(bins), key=lambda i: (acc[i], -abs(i - bins / 2)))
    d, u, got = poc, poc, acc[poc]
    while total > 0 and got < area * total - 1e-9 and (d > 0 or u < bins - 1):
        below = acc[d - 1] if d > 0 else -1
        above = acc[u + 1] if u < bins - 1 else -1
        if above >= below:
            u += 1
            got += acc[u]
        else:
            d -= 1
            got += acc[d]
    return {"edges": edges, "vols": acc, "poc": poc, "va": (d, u), "total": total, "area": area}


# ── 명세 선택 항목 ────────────────────────────────────────────────────────
def _int(v, lo, hi, dflt):
    try:
        x = int(float(v))
    except (TypeError, ValueError):
        return dflt
    return max(lo, min(hi, x))


def _num(v, lo, hi, dflt):
    try:
        x = float(v)
    except (TypeError, ValueError):
        return dflt
    return dflt if math.isnan(x) else max(lo, min(hi, x))


def has_fin(spec):
    """기존 캔들(거래량만) 말고 새 선택 항목이 있는지"""
    return any(spec.get(k) for k in ("ma", "bb", "vprofile", "macd", "rsi")) or "panels" in spec or spec.get("kind") == "ohlc"


def options(spec, n, has_vol):
    """명세 → 정리된 선택 항목 + 뺀 것 안내. n = 기간 수"""
    notes = []
    ma_in = spec.get("ma") or []
    if not isinstance(ma_in, list):
        ma_in = []
    ws = sorted({_int(w, 1, MA_MAX, 0) for w in ma_in} - {0})
    ma = [w for w in ws if 2 <= w <= n]
    if len(ma) < len(ws):
        notes.append(f"이동평균 {'·'.join(str(w) for w in ws if w not in ma)}일은 기간 {n}개보다 길어 뺌")
    bb = spec.get("bb")
    if bb:
        bb = bb if isinstance(bb, dict) else {}
        bb = {"period": _int(bb.get("period"), 2, 200, 20), "k": round(_num(bb.get("k"), 0.5, 4, 2.0), 2)}
        if bb["period"] > n:
            notes.append(f"볼린저 밴드({bb['period']}일)는 기간 {n}개보다 길어 뺌")
            bb = None
    else:
        bb = None
    m = spec.get("macd") if isinstance(spec.get("macd"), dict) else {}
    mp = (_int(m.get("fast"), 2, 100, 12), _int(m.get("slow"), 3, 200, 26), _int(m.get("signal"), 2, 50, 9))
    if mp[0] >= mp[1]:
        mp = (12, 26, 9)
    rp = _int((spec.get("rsi") or {}).get("period") if isinstance(spec.get("rsi"), dict) else None, 2, 100, 14)
    if "panels" in spec:
        raw = spec.get("panels") if isinstance(spec.get("panels"), list) else []
        panels = [p for p in PANELS if p in raw]
    else:  # 예전 캔들: 거래량 열이 있으면 거래량 막대
        panels = ["volume"] if has_vol else []
    if "volume" in panels and not has_vol:
        panels.remove("volume")
        notes.append("거래량 열이 없어 거래량 칸을 뺌")
    if "macd" in panels and n < mp[1] + mp[2] - 1:
        panels.remove("macd")
        notes.append(f"MACD 는 기간이 {mp[1] + mp[2] - 1}개 이상이어야 해 뺌")
    if "rsi" in panels and n < rp + 1:
        panels.remove("rsi")
        notes.append(f"RSI 는 기간이 {rp + 1}개 이상이어야 해 뺌")
    vp = spec.get("vprofile")
    if vp:
        vp = vp if isinstance(vp, dict) else {}
        vp = {"bins": _int(vp.get("bins"), 4, 100, 24), "va": vp.get("va", True) is not False}
        if not has_vol:
            notes.append("거래량 열이 없어 매물대를 뺌")
            vp = None
    else:
        vp = None
    ud = spec.get("colors_updown") if spec.get("colors_updown") in UPDOWN else "kr"
    return {"ma": ma, "bb": bb, "panels": panels, "vprofile": vp, "updown": ud, "macd": mp, "rsi": rp, "notes": notes}


def compute(table, spec, d=None):
    """category_data → 가격·거래량 + 지표. 반환 dict 또는 None(시가·고가·저가·종가 4열이 없음)"""
    import core
    d = d or core.category_data(table, spec)
    px = [s for s in d["series"] if not s.get("role")]
    if len(px) < 4 or not d["cats"]:
        return None
    vol = next((s for s in d["series"] if s.get("role") == "volume" and not s["hidden"]), None)
    O, H, L, C = (list(s["values"]) for s in px[:4])
    n = len(d["cats"])
    op = options(spec, n, vol is not None)
    out = {"cats": d["cats"], "O": O, "H": H, "L": L, "C": C, "px": px[:4], "unit": px[3].get("unit", ""),
           "vol": vol, "V": list(vol["values"]) if vol else None, "opt": op, "note": d.get("note") or ""}
    out["ma"] = [(w, sma(C, w)) for w in op["ma"]]
    out["bb"] = bollinger(C, op["bb"]["period"], op["bb"]["k"]) if op["bb"] else None
    out["macd"] = macd(C, *op["macd"]) if "macd" in op["panels"] else None
    out["rsi"] = rsi(C, op["rsi"]) if "rsi" in op["panels"] else None
    out["vp"] = volume_profile(H, L, out["V"], op["vprofile"]["bins"]) if op["vprofile"] and vol else None
    return out


def describe(spec):
    """제목·이유에 붙이는 짧은 설명 (예: '이동평균 5·20 · 볼린저 20,2σ')"""
    parts = []
    ma = [w for w in (spec.get("ma") or []) if isinstance(w, (int, float))]
    if ma:
        parts.append("이동평균 " + "·".join(str(int(w)) for w in ma))
    if spec.get("bb"):
        b = spec["bb"] if isinstance(spec["bb"], dict) else {}
        parts.append(f"볼린저 {_int(b.get('period'), 2, 200, 20)},{_num(b.get('k'), 0.5, 4, 2.0):g}σ")
    pn = [PANEL_KO[p] for p in PANELS if p in (spec.get("panels") or [])]
    if pn and pn != ["거래량"]:
        parts.append("·".join(pn))
    if spec.get("vprofile"):
        parts.append("매물대")
    return " · ".join(parts)
