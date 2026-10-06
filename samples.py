"""샘플 데이터 3종 (가상의 값 — 화면 시연·selftest 용)"""
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

SAMPLES = [
    {"id": "budget", "label": "연도별 예산", "text": BUDGET},
    {"id": "measure", "label": "실험 측정값(산점)", "text": MEASURE},
    {"id": "share", "label": "부서별 비율", "text": SHARE},
]
