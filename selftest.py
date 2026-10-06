#!/usr/bin/env python3
"""LLM 없이 검증: 숫자·단위·머리글·단위 행·전치 → 열 타입 추론 → 후보 추천 규칙(부적합 차트 배제) → SVG 렌더(15종)
→ 4개 형식 내보내기 + 파일을 다시 열어 차트·내장 워크북·수식 참조 구조 검사 → LLM 숫자 거르기 → 이력 → HTTP.
WORKSPACE 는 임시 폴더로 바꿔 실데이터 폴더에 흔적을 남기지 않는다.
python3 selftest.py            # SELFTEST_OUT=폴더 를 주면 SVG·PNG·문서를 거기 저장(눈으로 확인용)"""
import base64, io, json, os, re, shutil, sys, tempfile, threading, urllib.request, zipfile

TMP = tempfile.mkdtemp(prefix="chart-selftest-")
os.environ["WORKSPACE"] = TMP
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app, core, export, samples, svg  # noqa: E402
from xml.dom import minidom  # noqa: E402

OUT = os.environ.get("SELFTEST_OUT")
ok = 0


def check(cond, msg):
    global ok
    assert cond, msg
    ok += 1


try:
    check(app.WS == TMP, "WORKSPACE 가 임시 폴더가 아님")
    # 1) 숫자·단위
    for s, v, u in [("1,234", 1234, ""), ("12.5%", 12.5, "%"), ("(300)", -300, ""), ("△15", -15, ""), ("−3.2", -3.2, ""),
                    ("₩1,000", 1000, ""), ("3.5억원", 3.5, "억원"), ("12 kg", 12, "kg"), ("1.2e3", 1200, "")]:
        got = core.split_num(s)
        check(got[0] is not None and abs(got[0] - v) < 1e-9 and got[1] == u, f"split_num {s} → {got}")
    for s in ["-", "", "n/a", "가나다", "A-12", "2023-01-05x"]:
        check(core.num(s) is None, f"숫자 아님: {s}")
    check(all(core.is_date(s) for s in ["2024", "2024년", "2024-03", "2024.03.05", "2024년 3월", "3월", "1분기", "2024 Q2", "상반기"]), "날짜 인식")

    # 2) 정규화: 제목·단위 줄, 머리글 단위, 연도=날짜
    t, notes = core.normalize(core.read_text(samples.BUDGET))
    cols = {c["name"]: c for c in t["columns"]}
    check(t["title"] == "연도별 연구개발 예산", t["title"])
    check(cols["연도"]["type"] == "date" and cols["정부출연금"]["type"] == "num" and cols["정부출연금"]["unit"] == "억원", cols)
    check(cols["인력"]["unit"] == "명", "머리글 안 단위")
    # 머리글 아래 단위 행 + 쉼표 천단위 + 퍼센트 열
    t2, n2 = core.normalize(core.read_text("부서\t예산\t비율\n\t(백만원)\t(%)\nA부\t1,200\t30.5\nB부\t2,500\t45\nC부\t900\t24.5"))
    c2 = {c["name"]: c for c in t2["columns"]}
    check(c2["예산"]["unit"] == "백만원" and c2["비율"]["unit"] == "%" and len(t2["rows"]) == 3, (c2, n2))
    # 머리글 없음
    t3, n3 = core.normalize([["1", "2"], ["3", "4"], ["5", "6"]])
    check([c["name"] for c in t3["columns"]] == ["열1", "열2"] and len(t3["rows"]) == 3, t3)
    # 연도가 가로 머리글 → 자동 전치
    t4, n4 = core.normalize(core.read_text("구분\t2021\t2022\t2023\t2024\n예산\t10\t12\t15\t18\n인건비\t5\t6\t6\t7"))
    check(t4["columns"][0]["name"] == "구분" and t4["columns"][0]["type"] == "date" and len(t4["rows"]) == 4 and any("바꿈" in n for n in n4), (t4, n4))
    # CSV (따옴표 안 쉼표) · cp949
    t5, _ = core.normalize(core.read_csv_bytes('지역,값\n"서울, 강남","1,200"\n부산,800\n'.encode("cp949")))
    check(t5["rows"][0] == ["서울, 강남", "1,200"] and t5["columns"][1]["type"] == "num", t5)
    # XLSX 업로드 (시트 2개)
    import openpyxl
    wb = openpyxl.Workbook(); wb.active.title = "요약"; wb.active.append(["x"])
    ws = wb.create_sheet("측정"); ws.append(["시료", "값(mm)"]); [ws.append([f"S{i}", i * 1.5]) for i in range(1, 6)]
    b = io.BytesIO(); wb.save(b)
    r = app.do_parse({"file_b64": base64.b64encode(b.getvalue()).decode(), "filename": "t.xlsx", "sheet": "측정"})
    check(r["sheets"] == ["요약", "측정"] and r["table"]["columns"][1]["unit"] == "mm" and len(r["table"]["rows"]) == 5, r)

    # 3) 추천 규칙
    def kinds(text):
        tb, _ = core.normalize(core.read_text(text))
        return tb, core.recommend(tb)
    tb_b, cb = kinds(samples.BUDGET)
    kb = [c["kind"] for c in cb]
    check(kb[0] == "line" and "panels" in kb and "col_stack" in kb and "pie" not in kb and 6 <= len(cb) <= 12, kb)
    dual = next(c for c in cb if c["kind"] == "dual")
    check(dual["y2"] == "인력" and dual["ys"] == ["정부출연금", "수탁과제"] and "독립" in dual["why"] and "index" in kb, dual)
    tb_m, cm = kinds(samples.MEASURE)
    km = [c["kind"] for c in cm]
    check(km[0] == "scatter" and cm[0].get("trend") and "hist" in km and "box" in km and "pie" not in km, km)
    tb_s, cs = kinds(samples.SHARE)
    ks = [c["kind"] for c in cs]
    check("pie" in ks and "doughnut" in ks and "bar" in ks, ks)
    many = "부서\t값\n" + "\n".join(f"부서{i}\t{i + 3}" for i in range(30))
    _, cmany = kinds(many)
    check("pie" not in [c["kind"] for c in cmany] and "doughnut" not in [c["kind"] for c in cmany], "범주 30개 원그래프 금지")
    neg = "항목\t손익\nA\t10\nB\t-5\nC\t7"
    _, cneg = kinds(neg)
    check("pie" not in [c["kind"] for c in cneg], "음수 원그래프 금지")
    longf = "연도\t부서\t예산\n" + "\n".join(f"{y}\t{d}\t{v}" for y in (2022, 2023, 2024) for d, v in (("A", 10), ("B", 20), ("C", 5)))
    tb_l, cl = kinds(longf)
    check(any(c.get("group") == "부서" for c in cl), "긴 형식 → 피벗(그룹)")
    d = core.category_data(tb_l, {"kind": "col_group", "x": "연도", "ys": ["예산"], "group": "부서"})
    check(d["cats"] == ["2022", "2023", "2024"] and [s["name"] for s in d["series"]] == ["A", "B", "C"] and d["series"][1]["values"] == [20, 20, 20], d)
    dup = core.category_data(core.normalize(core.read_text("팀\t값\nA\t1\nA\t2\nB\t5"))[0], {"kind": "col", "x": "팀", "ys": ["값"]})
    check(dup["series"][0]["values"] == [3, 5] and dup["note"], "중복 x 합계")
    srt = core.category_data(tb_s, {"kind": "bar", "x": "부서", "ys": ["예산 비중"], "sort": "desc"})
    check(srt["series"][0]["values"] == sorted(srt["series"][0]["values"], reverse=True), "정렬")
    check(core.quantile([1, 2, 3, 4], .25) == 1.75, "QUARTILE.INC 와 같은 사분위")
    hb = core.hist_bins([1, 2, 2, 3, 9, 10])
    check(sum(c for _, _, c in hb) == 6, hb)

    # 3-1) 이중축 범위: 0 아래·위 칸 수를 맞춰 0선 공유, 음수면 두 축 0 위치가 같음
    L, R, al = svg.dual_domains([120, -40, 210], [8.5, -2.1, 12.4])
    check(al and abs(-L[0] / (L[1] - L[0]) - (-R[0] / (R[1] - R[0]))) < 1e-9 and len(L[2]) == len(R[2]), (L, R))
    L, R, al = svg.dual_domains([1850, 2310], [1520, 1655])
    check(L[0] == 0 and R[0] == 0 and len(L[2]) == len(R[2]), "둘 다 0부터 + 눈금 수 같음")
    L, R, al = svg.dual_domains([1850, 2310], [1520, 1655], zero=False, left_zero=True)
    check(not al and R[0] > 0, "오른쪽 축 데이터 범위")
    neg_t = core.normalize(core.read_text("연도\t영업이익(억원)\t이익률(%)\n2020\t120\t8.5\n2021\t-40\t-2.1\n2022\t60\t3.2\n2023\t150\t9.8"))[0]
    neg_dual = {"kind": "dual", "x": "연도", "ys": ["영업이익"], "y2": "이익률", "title": "영업이익과 이익률"}

    # 4) SVG: 모든 종류 · XML 유효 · 팔레트 고정 순서 · 숨긴 계열 색 유지 · 외부 참조 없음
    all_specs = []
    for tb, cands in ((tb_b, cb), (tb_m, cm), (tb_s, cs), (tb_l, cl)):
        for c in cands:
            all_specs.append((tb, c))
    extra = [(tb_b, {"kind": k, "x": "연도", "ys": ["정부출연금", "수탁과제"], "title": k}) for k in
             ("bar_group", "bar_stack", "col_stack100", "area", "index")]
    extra += [(neg_t, neg_dual), (tb_b, dict(dual, dual_shape="line_line", ys=["정부출연금"]))]
    all_specs += extra
    seen = set()
    for tb, sp in all_specs:
        s = svg.render(tb, sp)
        minidom.parseString(s.encode())
        check("그릴 수 없음" not in s and "http://" not in s.replace("http://www.w3.org/2000/svg", ""), (sp["kind"], s[:200]))
        seen.add(sp["kind"])
        if OUT:
            os.makedirs(OUT, exist_ok=True)
            open(os.path.join(OUT, f"{sp['kind']}-{len(seen)}.svg"), "w").write(s)
    check(len(seen) >= 16 and "dual" in seen, f"렌더한 종류 {sorted(seen)}")
    sd = svg.render(tb_b, dual)
    check("인력 (오른쪽 축)" in sd and "인력 (명)" in sd and "(억원)" in sd or "단위: 억원" in sd, "이중축 범례·축 제목")
    check(svg.darken(svg.PAL[2]) in sd, "오른쪽 축 제목은 계열색(읽히게 어둡게)")
    s_all = svg.render(tb_b, {"kind": "line", "x": "연도", "ys": ["정부출연금", "수탁과제", "자체사업"]})
    s_hid = svg.render(tb_b, {"kind": "line", "x": "연도", "ys": ["정부출연금", "수탁과제", "자체사업"], "hidden": ["수탁과제"]})
    check(svg.PAL[2] in s_hid and svg.PAL[1] not in s_hid and all(c in s_all for c in svg.PAL[:3]), "숨겨도 색이 계열을 따라감")
    check(svg.PAL == ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"], "팔레트 순서")
    s_t = svg.render(tb_b, {"kind": "col", "x": "연도", "ys": ["정부출연금"], "title": "제목<&>", "labels": "all"})
    check("제목&lt;&amp;&gt;" in s_t and "2,310" in s_t, "이스케이프·천단위 라벨")

    # 5) 내보내기 4형식 + 구조 검사
    outdir = OUT or os.path.join(TMP, "out")
    os.makedirs(outdir, exist_ok=True)
    pick = {"budget": (tb_b, cb), "measure": (tb_m, cm), "share": (tb_s, cs), "long": (tb_l, cl)}
    for name, (tb, cands) in pick.items():
        for fmt in ("xlsx", "pptx", "docx") + (("hwpx",) if export.hwpx_available() else ()):
            p = os.path.join(outdir, f"{name}.{fmt}")
            rep = export.BUILDERS[fmt](tb, cands, p, doc_title=f"{name} 시험")
            v = export.verify(p)
            check(v["ok"], f"{name}.{fmt} 검사 실패: {v['errors']}")
            nat = sum(1 for r in rep if "native" in r["modes"] or "hwp-native" in r["modes"])
            check(nat >= len(cands) - 2 and v["charts"] >= nat, f"{name}.{fmt} 네이티브 {nat}/{len(cands)} 차트 {v['charts']}")
            check(all(r["label"] for r in rep), "형식 보고")
    # 이중축: PPTX·DOCX 는 차트 XML 을 고친 콤보(막대+선, 값 축 2개, 오른쪽 crosses=max), XLSX 는 openpyxl 보조축, HWPX 는 그림+표
    from lxml import etree
    from pptx.chart.chart import Chart
    from pptx.oxml import parse_xml as pparse
    CNS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    for fmt in ("pptx", "docx", "xlsx") + (("hwpx",) if export.hwpx_available() else ()):
        p = os.path.join(outdir, f"dual.{fmt}")
        export.BUILDERS[fmt](tb_b, [dual, dict(dual, dual_shape="line_line", ys=["정부출연금"])], os.path.join(outdir, f"dual-b.{fmt}"))
        vb = export.verify(os.path.join(outdir, f"dual-b.{fmt}"))
        check(vb["ok"] and (fmt == "hwpx" or vb["charts"] == 2), (fmt, vb["errors"]))
        rep = export.BUILDERS[fmt](neg_t, [neg_dual], p)
        v = export.verify(p)
        check(v["ok"], (fmt, v["errors"]))
        if fmt == "hwpx":
            check(rep[0]["modes"] == ["image"] and v["images"] == 1 and v["charts"] == 0, (rep, v))
            continue
        check(rep[0]["modes"] == ["native"], rep)
        zz = zipfile.ZipFile(p)
        cn = next(n for n in zz.namelist() if re.search(r"charts/chart1\.xml$", n))
        root = etree.fromstring(zz.read(cn))
        pa = root.find(f"{CNS}chart").find(f"{CNS}plotArea")
        kinds_ = [e.tag.split("}")[1] for e in pa if e.tag.endswith("Chart")]
        vaxes = pa.findall(f"{CNS}valAx")
        check(kinds_ == ["barChart", "lineChart"] and len(vaxes) == 2, (fmt, kinds_, len(vaxes)))
        check(any(a.find(f"{CNS}crosses") is not None and a.find(f"{CNS}crosses").get("val") == "max" for a in vaxes), "오른쪽 축 crosses=max")
        mins = [float(a.find(f"{CNS}scaling").find(f"{CNS}min").get("val")) for a in vaxes]
        maxs = [float(a.find(f"{CNS}scaling").find(f"{CNS}max").get("val")) for a in vaxes]
        check(abs(mins[0] / (maxs[0] - mins[0]) - mins[1] / (maxs[1] - mins[1])) < 1e-9, ("0선 맞춤", mins, maxs))
        line_ids = {e.get("val") for e in pa.find(f"{CNS}lineChart").findall(f"{CNS}axId")}
        bar_ids = {e.get("val") for e in pa.find(f"{CNS}barChart").findall(f"{CNS}axId")}
        check(len(line_ids & {a.find(f"{CNS}axId").get("val") for a in vaxes}) == 1 and line_ids != bar_ids, "선은 다른 값 축")
        if fmt != "xlsx":
            ch2 = Chart(pparse(zz.read(cn)), None)
            names = [s.name for pl in ch2.plots for s in pl.series]
            check([pl.__class__.__name__ for pl in ch2.plots] == ["BarPlot", "LinePlot"] and names == ["영업이익", "이익률 (오른쪽 축)"], names)
            check(list(ch2.plots[1].series[0].values) == [8.5, -2.1, 3.2, 9.8], "오른쪽 축 값")

    # PPTX: 차트 값이 원본과 같고 '데이터 편집'용 내장 워크북 있음
    from pptx import Presentation
    prs = Presentation(os.path.join(outdir, "budget.pptx"))
    charts = [sh.chart for sl in prs.slides for sh in sl.shapes if sh.has_chart]
    first = charts[0]
    check(list(first.plots[0].categories) == ["2019", "2020", "2021", "2022", "2023", "2024", "2025"], list(first.plots[0].categories))
    check(list(first.plots[0].series[0].values)[-1] == 2310, list(first.plots[0].series[0].values))
    check(first.part.chart_workbook.xlsx_part is not None, "PPTX 내장 워크북")
    # DOCX: 차트 파트 = 같은 데이터
    z = zipfile.ZipFile(os.path.join(outdir, "budget.docx"))
    check("word/charts/chart1.xml" in z.namelist() and any(n.startswith("word/embeddings/") for n in z.namelist()), "DOCX 차트 파트")
    check(b"2310" in z.read("word/charts/chart1.xml"), "DOCX 차트 캐시 값")
    # XLSX: 그림 시트 값이 데이터 시트를 가리키는 수식, 히스토그램 COUNTIFS, 상자그림 QUARTILE
    wbx = openpyxl.load_workbook(os.path.join(outdir, "budget.xlsx"))
    check(wbx["그림1"]["B5"].value == "='데이터'!$B$2", wbx["그림1"]["B5"].value)
    wbm = openpyxl.load_workbook(os.path.join(outdir, "measure.xlsx"))
    fs = " ".join(str(getattr(c.value, "text", c.value)) for ws in wbm.worksheets for row in ws.iter_rows() for c in row if c.value is not None)
    check("COUNTIFS(" in fs and "QUARTILE.INC" in fs, "히스토그램·상자그림 수식")
    if export.hwpx_available():
        zh = zipfile.ZipFile(os.path.join(outdir, "budget.hwpx"))
        check(any(n.startswith("Chart/chart") for n in zh.namelist()), "HWPX 한글 차트 파트")
        sec = zh.read("Contents/section0.xml").decode()
        check("chartIDRef=" in sec and "<hp:tbl" in sec, "HWPX 차트+표")

    # 6) LLM 숫자 거르기 (가짜 LLM)
    orig = app.llm
    app.llm = lambda system, user, model=None, temperature=0.3: json.dumps({
        "c1": {"title": "연도별 연구개발 예산 추이", "insight": "2019년부터 2025년까지 세 항목의 흐름을 비교"},
        "c2": {"title": "예산 25% 증가", "insight": "정부출연금이 꾸준히 늘어남을 보여 줌"},
        "zz": {"title": "없는 후보"}}, ensure_ascii=False)
    r = app.describe(tb_b, cb)
    app.llm = orig
    check(r["items"]["c1"]["title"] == "연도별 연구개발 예산 추이" and "2019년" in r["items"]["c1"]["insight"], r)
    check("title" not in r["items"]["c2"] and r["items"]["c2"]["insight"] and r["dropped"] == 1 and "zz" not in r["items"], r)

    # 7) 이력 + HTTP
    srv = app.ThreadingHTTPServer(("127.0.0.1", 0), app.H)
    port = srv.server_address[1]
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{port}"

    def post(path, body):
        req = urllib.request.Request(base + path, json.dumps(body).encode(), {"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=120) as resp:
            return json.load(resp)
    with urllib.request.urlopen(base + "/") as resp:
        html = resp.read().decode()
        check(resp.headers.get("X-Author") and "data-sig" in html and 'name="author"' in html, "저작자 표기")
        check(not re.search(r'<(script|link)[^>]+(src|href)="https?://', html), "외부 CDN 없음")
    p = post("/api/parse", {"text": samples.SHARE})
    sg = post("/api/suggest", {"table": p["table"], "source": "selftest"})
    check(sg["id"] and sg["candidates"] and sg["candidates"][0]["svg"].startswith("<svg"), "suggest")
    rr = post("/api/render", {"table": p["table"], "spec": dict(sg["candidates"][0], title="바꾼 제목", legend="none")})
    check("바꾼 제목" in rr["svg"], "편집 반영")
    ex = post("/api/export", {"id": sg["id"], "table": p["table"], "specs": sg["candidates"][:2], "format": "pptx", "title": "시험 문서"})
    check(ex["verify"]["ok"] and ex["file"].endswith(".pptx"), ex)
    with urllib.request.urlopen(base + "/" + ex["url"]) as resp:
        check(resp.read(2) == b"PK", "파일 내려받기")
    h = json.load(urllib.request.urlopen(base + "/api/history"))
    check(h and h[0]["id"] == sg["id"] and h[0]["exports"], h)
    post("/api/history/delete", {"id": sg["id"]})
    check(not json.load(urllib.request.urlopen(base + "/api/history")), "이력 삭제")
    try:
        post("/api/export", {"table": p["table"], "specs": [], "format": "pptx"})
        check(False, "빈 선택 거부")
    except urllib.error.HTTPError as e:
        check(e.code == 400, "빈 선택 400")
    srv.shutdown()
    print(f"selftest OK — {ok}개 확인 (hwpx: {'kordoc' if export.hwpx_available() else '생략 — kordoc 없음'})")
finally:
    shutil.rmtree(TMP, ignore_errors=True)
