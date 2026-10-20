"""샘플 데이터 13종 (가상의 값 — 화면 시연·selftest 용)"""
import random

BUDGET = """연도별 연구개발 예산 (단위: 억원)
연도\t정부출연금\t수탁과제\t자체사업\t인력(명)
2019\t1,850\t620\t210\t1,520
2020\t1,920\t655\t230\t1,548
2021\t2,010\t702\t245\t1,561
2022\t2,105\t730\t260\t1,590
2023\t2,180\t790\t270\t1,602
2024\t2,240\t845\t290\t1,630
2025\t2,310\t880\t305\t1,655"""

SHARE = """부서\t예산 비중(%)\t인원(명)
원자력안전연구부\t32.0\t210
방사선과학연구부\t18.5\t132
핵주기기술개발부\t22.0\t158
인공지능응용연구실\t9.5\t64
기획·지원부서\t18.0\t121"""


def _measure():
    rnd = random.Random(7)
    rows = ["시편\t재질\t온도(°C)\t인장강도(MPa)\t연신율(%)"]
    for i in range(36):
        mat = "A" if i % 2 else "B"
        t = 25 + i * 16 + rnd.uniform(-6, 6)
        s = (640 if mat == "A" else 600) - 0.36 * t + rnd.gauss(0, 16)
        e = 17 + 0.021 * t + rnd.gauss(0, 1.4)
        rows.append(f"S-{i + 1:02d}\t{mat}\t{t:.1f}\t{s:.1f}\t{e:.1f}")
    return "\n".join(rows)


MEASURE = _measure()

def _stock():
    rnd = random.Random(11)
    rows = ["일자\t시가\t고가\t저가\t종가\t거래량(주)"]
    p = 71200
    import datetime
    d = datetime.date(2025, 9, 1)
    k = 0
    while k < 22:
        if d.weekday() < 5:
            o = p + rnd.randint(-6, 6) * 100
            c = o + rnd.randint(-14, 15) * 100
            hi = max(o, c) + rnd.randint(1, 9) * 100
            lo = min(o, c) - rnd.randint(1, 9) * 100
            v = rnd.randint(80, 260) * 1000 + (120000 if abs(c - o) > 900 else 0)
            rows.append(f"{d:%Y-%m-%d}\t{o:,}\t{hi:,}\t{lo:,}\t{c:,}\t{v:,}")
            p = c
            k += 1
        d += datetime.timedelta(days=1)
    return "\n".join(rows)


STOCK = _stock()

SKILLS = """직무 역량 평가 (5점 만점)
역량\t김연구\t이선임\t팀 평균
문제 해결\t4.5\t3.8\t3.9
전문 지식\t4.8\t4.1\t4.0
의사소통\t3.2\t4.4\t3.7
협업\t3.6\t4.6\t3.9
기획력\t4.1\t3.5\t3.6
자기 개발\t4.4\t3.9\t3.5"""

PNL = """월별 손익 증감 (단위: 백만원)
항목\t금액
기초 잔액\t1,200
1월\t180
2월\t-95
3월\t240
4월\t-160
5월\t310
6월\t-70
기말 잔액\t1,605"""

FUNNEL = """채용 단계별 지원자 수
단계\t인원(명)
서류 접수\t1,240
서류 합격\t430
필기 합격\t182
면접 합격\t64
최종 합격\t28"""

TARGET = """부서별 연간 목표 대비 실적 (단위: 억원)
부서\t목표\t실적
원자력안전연구부\t120\t131
방사선과학연구부\t85\t78
핵주기기술개발부\t96\t102
인공지능응용연구실\t40\t33
기획·지원부서\t30\t29"""

RATE = """기준금리 추이 (단위: %)
월\t기준금리
2024-01\t3.50
2024-02\t3.50
2024-03\t3.50
2024-04\t3.50
2024-05\t3.50
2024-06\t3.50
2024-07\t3.50
2024-08\t3.50
2024-09\t3.50
2024-10\t3.25
2024-11\t3.00
2024-12\t3.00
2025-01\t3.00
2025-02\t2.75
2025-03\t2.75
2025-04\t2.75
2025-05\t2.50
2025-06\t2.50"""

REGION = """시도\t인구(만명)\t1인당 소득(백만원)\t연구기관 수(곳)
서울\t938\t26.4\t412
부산\t329\t21.0\t58
대구\t237\t20.1\t44
인천\t300\t21.3\t61
광주\t142\t21.1\t39
대전\t144\t22.6\t118
울산\t110\t25.9\t21
세종\t39\t23.5\t27
경기\t1,363\t22.8\t356
강원\t152\t20.4\t30
충북\t159\t21.9\t41
충남\t213\t22.1\t52
전북\t175\t20.0\t36
전남\t180\t20.6\t29
경북\t255\t20.8\t47
경남\t325\t21.2\t55"""

KPI = """월별 운영 지표
월\t방문자(명)\t신규 가입(명)\t문의(건)\t처리 시간(시간)\t만족도(점)\t오류(건)
2025-01\t12,400\t820\t310\t18.5\t4.1\t42
2025-02\t11,900\t760\t295\t19.2\t4.0\t45
2025-03\t13,800\t910\t340\t17.1\t4.2\t38
2025-04\t14,600\t990\t362\t16.4\t4.2\t33
2025-05\t15,100\t1,040\t371\t15.8\t4.3\t31
2025-06\t14,200\t950\t355\t16.9\t4.2\t36
2025-07\t16,800\t1,180\t402\t15.2\t4.4\t29
2025-08\t17,300\t1,220\t415\t14.6\t4.4\t27
2025-09\t16,500\t1,150\t398\t14.9\t4.5\t25"""

RAIN = """월별 평균 강수량 (단위: mm)
월\t강수량
1월\t17
2월\t28
3월\t47
4월\t72
5월\t104
6월\t149
7월\t395
8월\t364
9월\t169
10월\t52
11월\t53
12월\t22"""

def _stock_daily(days=160):
    """가상 데이터 — 일봉 160일 (고정 씨앗 무작위 걸음: 추세가 바뀌는 구간 + 변동성 변화, 큰 움직임 날엔 거래량 증가)"""
    import datetime
    rnd = random.Random(20251020)
    rows = ["가상 주가 일봉 (가상 데이터)", "일자\t시가\t고가\t저가\t종가\t거래량(주)"]
    p, d, k = 52000.0, datetime.date(2025, 3, 3), 0
    while k < days:
        if d.weekday() < 5:
            drift = 0.0016 if k < 60 else (-0.0022 if k < 105 else 0.0028)
            vol = 0.014 if k < 90 else 0.022
            o = p * (1 + rnd.gauss(0, vol * 0.3))
            c = o * (1 + drift + rnd.gauss(0, vol))
            hi = max(o, c) * (1 + abs(rnd.gauss(0, vol * 0.5)))
            lo = min(o, c) * (1 - abs(rnd.gauss(0, vol * 0.5)))
            r50 = lambda x: int(round(x / 50.0)) * 50
            o, c, hi, lo = r50(o), r50(c), r50(hi), r50(lo)
            hi, lo = max(hi, o, c), min(lo, o, c)
            v = int(rnd.uniform(0.7, 1.3) * 180000 * (1 + 12 * abs(c - o) / o)) // 100 * 100
            rows.append(f"{d:%Y-%m-%d}\t{o:,}\t{hi:,}\t{lo:,}\t{c:,}\t{v:,}")
            p = c
            k += 1
        d += datetime.timedelta(days=1)
    return "\n".join(rows)


STOCK_DAILY = _stock_daily()

SAMPLES = [
    {"id": "budget", "label": "연도별 예산", "text": BUDGET},
    {"id": "measure", "label": "실험 측정값(산점)", "text": MEASURE},
    {"id": "share", "label": "부서별 비율", "text": SHARE},
    {"id": "stock", "label": "주가 (시가·고가·저가·종가)", "text": STOCK},
    {"id": "stock_daily", "label": "주가 일봉 160일 (가상 데이터 · 이동평균·지표)", "text": STOCK_DAILY},
    {"id": "skills", "label": "역량 평가 (방사형)", "text": SKILLS},
    {"id": "pnl", "label": "월별 손익 (폭포)", "text": PNL},
    {"id": "funnel", "label": "채용 단계 (깔때기)", "text": FUNNEL},
    {"id": "target", "label": "목표 대비 실적", "text": TARGET},
    {"id": "rate", "label": "기준금리 (계단선)", "text": RATE},
    {"id": "region", "label": "시도별 지표 (거품·트리맵)", "text": REGION},
    {"id": "kpi", "label": "월별 운영 지표 (스파크라인)", "text": KPI},
    {"id": "rain", "label": "월별 강수량 (로즈)", "text": RAIN},
]
