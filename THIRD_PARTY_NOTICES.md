# 제3자 구성요소 (Third-party notices) — chart-local

조사일 2026-10-12. 실제 파일(`requirements.txt`, `setup.sh`, `*.py` import, `export.py` 의 외부 실행, `ui.html`) 기준.
버전은 포털 서버에서 실제로 쓰는 Python 환경에서 `importlib.metadata` 로 확인한 값(`requirements.txt` 는 하한만 고정). 라이선스는 패키지 메타데이터·GitHub LICENSE 에서 확인했다.
이 표는 고지용 정리이며 법률 자문이 아니다. 상업 이용 전에는 원문을 다시 확인할 것.

## 동봉 (저장소에 들어 있음)

제3자 코드·글꼴·아이콘 동봉 없음. SVG 차트 렌더러(`svg.py`)는 이 저장소에서 새로 작성했고 외부 차트 라이브러리·CDN 을 쓰지 않는다. 샘플 데이터(`samples.py`)는 가상의 값.
(`ui.html` 의 `"D2Coding"` 은 글꼴 이름만 지정 — 글꼴 파일은 들어 있지 않음.)

## 설치 (`setup.sh` → pip, 동봉 안 함)

| 이름 | 버전 | 쓰임 | 라이선스 | 링크 | 상업 이용 | 비고 |
|---|---|---|---|---|---|---|
| openpyxl | 3.1.5 (`>=3.1`) | XLSX 내보내기·차트 | MIT | https://foss.heptapod.net/openpyxl/openpyxl | 가능 | |
| et_xmlfile | 2.0.0 | openpyxl 의존성 | MIT | https://foss.heptapod.net/openpyxl/et_xmlfile | 가능 | |
| python-pptx | 1.0.2 (`>=0.6.22`) | PPTX 차트·내장 워크북 (DOCX 차트 파트도 이것으로 생성) | MIT | https://github.com/scanny/python-pptx | 가능 | |
| python-docx | 1.2.0 (`>=1.1`) | DOCX 내보내기 | MIT | https://github.com/python-openxml/python-docx | 가능 | |
| lxml | 6.1.1 | python-pptx/docx 의존성, XML 조작 | BSD-3-Clause | https://github.com/lxml/lxml | 가능 | 휠에 libxml2·libxslt (MIT) 포함 |
| XlsxWriter | 3.2.9 | python-pptx 의존성 | BSD-2-Clause | https://github.com/jmcnamara/XlsxWriter | 가능 | |
| Pillow | 12.1.1 | python-pptx 의존성 | MIT-CMU (HPND) | https://github.com/python-pillow/Pillow | 가능 | |
| typing_extensions | 4.15.0 | 의존성 | PSF-2.0 | https://github.com/python/typing_extensions | 가능 | |

## 선택 사용 (있으면 씀, 동봉 안 함)

| 이름 | 버전 | 쓰임 | 라이선스 | 링크 | 상업 이용 | 비고 |
|---|---|---|---|---|---|---|
| kordoc | 4.17.2 (`../kordoc-local/node_modules`) | HWPX 내보내기 — `node` 로 CLI 실행 (`export.py`) | MIT (Copyright (c) 2026 chrisryugj) | https://github.com/chrisryugj/kordoc | 가능 | 별도 프로세스로 실행, 코드 복사 없음 |
| Node.js | 시스템 | kordoc 실행 | MIT | https://nodejs.org | 가능 | |
| rsvg-convert (librsvg) | 시스템 | HWPX 의 상자그림·히트맵을 PNG 로 변환 | LGPL-2.1+ | https://gitlab.gnome.org/GNOME/librsvg | 조건부 | 외부 명령으로 호출만 함(링크·동봉 없음). 함께 재배포할 때만 LGPL 고지 |
| Ollama | 설치 시점 최신 | LLM 실행기(제목·설명 제안). 없으면 그 기능만 꺼짐 | MIT | https://github.com/ollama/ollama | 가능 | |

## 모델 가중치 (저장소에 없음)

| 이름 | 쓰임 | 라이선스 | 링크 | 상업 이용 | 비고 |
|---|---|---|---|---|---|
| Gemma 4 31B (`gemma4:31b`) | 포털 배포 기본 모델 | Apache-2.0 | https://ai.google.dev/gemma/docs/gemma_4_license | 가능 | `setup.sh` 는 모델을 내려받지 않고 서버의 첫 모델을 씀 |

## 확인 필요

없음 (2026-10-12 조사 기준).

## 이 저장소의 원저작물

`app.py`, `core.py`, `svg.py`, `export.py`, `samples.py`, `ui.html`, `selftest.py`, `setup.sh`, `README.md` — DongJu Kim (gggg8657), MIT (`LICENSE`).
