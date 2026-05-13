# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

사내 PPTX 보고서 자가점검 도구. 작성자가 제출 전 `/report-reviewer <pptx>` 한 번으로 오타·데이터 정합성·결론 논리 등 6개 카테고리를 검토받는 Claude Code Skill.

**핵심 제약**: Windows + MS PowerPoint 설치 필수 (슬라이드 → 이미지 변환에 COM 사용).

## 개발 환경 설정

모든 소스 코드와 테스트는 `skill_src/` 아래에 있다. pytest도 이 디렉토리 안에서 실행한다.

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
python -m venv .venv
.venv/Scripts/pip install -e ".[dev]"
```

## 주요 명령어

모든 명령은 `skill_src/` 기준 (절대 경로 권장).

```bash
# 단위 테스트만 (CI 가능, COM 불필요)
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
.venv/Scripts/pytest tests/unit/ -v

# 전체 테스트 (real_com 마커 자동 제외)
.venv/Scripts/pytest -v

# 단일 테스트
.venv/Scripts/pytest tests/unit/test_extractor.py::test_extract_table_shape -v

# COM 통합 테스트 (Windows + PowerPoint 필요, 수동 실행)
.venv/Scripts/pytest -m real_com -v

# CLI: PPTX 추출 + 이미지 변환 (인자: pptx_path, --out)
.venv/Scripts/python extract_and_render.py "C:/path/to/report.pptx" --out "C:/path/to/review_ws"

# CLI: 리포트 생성 (인자: work_dir, --out)
.venv/Scripts/python render_report.py "C:/path/to/review_ws" --out "C:/path/to/review_output"

# 글로벌 배포 (PowerShell)
./deploy.ps1
```

**워크플로 옵션 vs CLI 옵션 구분**:
- `/report-reviewer ... --resume|--rerun-stage2|--rerender`는 메인 Claude가 SKILL.md에 따라 산출물 파일 존재 여부로 분기하는 **워크플로 옵션**임
- `extract_and_render.py` / `render_report.py`는 이런 옵션이 없고, 매번 처음부터 실행됨 (개발자가 단계별 디버깅 시 그냥 다시 돌리면 됨)

골든 파일 재생성: `tests/fixtures/_make_fixtures.py` 실행. 이후 `tests/fixtures/golden/` 결과 수동 검토 필요.

## 아키텍처

### 3단 파이프라인

```
PPTX → [Python: 추출+이미지] → [Claude SA: 분석+검증] → [Python: 리포트]
```

1. **추출 단계** (`extract_and_render.py`): python-pptx로 텍스트·표·도형·위치 추출 → PowerPoint COM으로 슬라이드 JPG 변환 → `extracted.json` + `slide_inputs/slide_NNN.json` 생성
2. **분석 단계** (SKILL.md 워크플로): 메인 Claude가 subagent를 세 단계로 dispatch (총 1+7+1=9개 SA):
   - 1단계: `report-reviewer-slide-analyzer` 1개 — 슬라이드별 multimodal 분석 (이미지 + JSON 동시 입력)
   - 2단계: 7개 SA 동시 dispatch
     - 카테고리 6개(typo/terminology/data/conclusion/improvement/logic) — 슬라이드 단위 미시 검토, 결과는 raw `findings[]`로 concat → `findings_raw.json` 저장
     - `report-reviewer-document` 1개 — 문서 전체 거시 검토, `document_review` 객체로 저장
   - 2.5단계: `report-reviewer-merger` 1개 — 같은 원문·같은 대상의 multi-category 중복 지적을 단일 finding으로 종합 재작성. 출력은 `categories[]`·`source_finding_ids[]` 포함된 통합 `findings[]`
3. **렌더 단계** (`render_report.py`): 통합 findings + extracted → `review.md` + `review.html` (맨 앞 "문서 전체 평가" → 발견된 이슈(통합) → 슬라이드별 → 카테고리별 — 통합 finding은 자신의 모든 카테고리 그룹에 중복 표시)

`review_ws/` 산출물 (디버깅 시 확인 위치):
- `extracted.json` — Step 1 결과 (전체 메타데이터)
- `slides/slide_NNN.jpg`, `thumbnails/slide_NNN.jpg` — Step 1 이미지
- `slide_inputs/slide_NNN.json` — Step 2 SA 입력 (슬라이드별)
- `slide_summaries.json` — Step 2 결과 (1단계 SA 누적)
- `findings_raw.json` — Step 3 결과 (6개 카테고리 SA concat, 통합 전)
- `findings.json` — Step 3.5 결과 (Merger SA 통합 + document_review)

### 주요 데이터 흐름

`extracted.json` 스키마:
```
metadata: {file_path, title, author, slide_count, ...}
slides[]: {index, title, shapes[], notes, image_path, thumbnail_path, has_embedded}
  shapes[]: {shape_id("s{i}_sh{j}"), type, position_emu, position_pct(0~1), text, table, embedded_progid}
```

`findings.json` 스키마 (Merger SA 통합 후):
```
summary: {total_issues, by_severity, by_category}
findings[]: {id(F+숫자), categories[]("typo"|"terminology"|...),
             severity, slide_index, shape_id,
             position_pct, position_hint, quoted_text, issue, suggestion, evidence,
             source_finding_ids[]}
document_review: {위와 동일}
```

`findings_raw.json` 스키마 (Merger SA 이전 raw):
```
findings[]: {id(prefix+숫자), category(단수), severity, slide_index, shape_id, ...}
```

Finding ID prefix:
- raw (6개 카테고리 SA): `T`(typo), `TM`(terminology), `D`(data), `C`(conclusion), `I`(improvement), `L`(logic)
- 통합 (Merger): `F` 1종, `categories[]`·`source_finding_ids[]`로 원본 추적
- 거시 SA(document)는 ID 없음 — `document_review` 객체로 별도 저장 (미러링 안 함)

reporter는 단수 `category`와 복수 `categories[]` 둘 다 처리 (legacy 호환). 통합 finding은 자신의 모든 카테고리 그룹에 중복 표시 (A안).

### 파일별 책임

| 파일 | 역할 |
|---|---|
| `src/extractor.py` | PPTX → extracted dict (python-pptx, 결정론적) |
| `src/slide_renderer.py` | 슬라이드 → JPG + 썸네일 (PowerPoint COM, `pythoncom.CoInitialize` 쌍 보장) |
| `src/locator.py` | `position_pct` → 한국어 위치 hint (순수 함수) |
| `src/reporter_md.py` | findings → 마크다운 리포트 (표준 라이브러리만) |
| `src/reporter_html.py` | findings → HTML 리포트 (Jinja2, 썸네일+위치박스 오버레이) |
| `src/config.py` | 환경변수 → frozen dataclass |
| `templates/report.html.j2` | HTML 리포트 Jinja2 템플릿 (position-box CSS 이미 포함) |
| `SKILL.md` | 메인 Claude가 실행할 워크플로 가이드 (Step 0~5) |
| `agents_src/*.md` | 1+7+1=9개 subagent 시스템 프롬프트 (slide-analyzer + 6 카테고리 + document 거시 + merger 통합, `tools: Read`만) |

### slide_renderer의 두 모드

`render()`는 `extracted["slides"][i]["has_embedded"]` 플래그로 자동 선택:
- **모드 1** (`convert_pptx_to_images_without_embedding`): 단순 `slide.Export()` → `dict[int, Path]`
- **모드 2** (`convert_pptx_with_embedded_to_images`): 임베디드 OLE 순회. PDF만 DoVerb(3)+pdf2image로 별도 추출. PowerPoint/Excel/Word 임베디드는 슬라이드 캡처에 의존 (nested COM Dispatch 회피).

두 모드 모두 `try/finally`로 `pres.Close()` + `app.Quit()` + `pythoncom.CoUninitialize()` 보장.

## 테스트 구조

- `tests/unit/` — COM mock 사용, CI 가능 (자동 실행)
- `tests/integration/` — 두 종류:
  - `test_cli_render_report.py` — COM 불필요, 자동 실행
  - `test_cli_extract_and_render.py`, `test_slide_renderer_real.py` — `real_com` 마커, 수동 실행
- `tests/fixtures/` — `sample_text_only/table/image.pptx` (생성 스크립트: `_make_fixtures.py`), `sample_with_embedded.pptx`는 수동 생성 필요

COM mock은 `conftest.py`의 `mock_powerpoint_com` fixture: `monkeypatch.setitem(sys.modules, "win32com", ...)` + Export 호출 시 PIL 더미 이미지 실제 생성.

골든 파일 위치: `tests/fixtures/golden/` (변경 시 수동 재생성 후 체크리스트 검토 필요).

## 환경변수

`Config.from_env(os.environ)`으로 읽음:

| 변수 | 기본값 | 설명 |
|---|---|---|
| `REPORT_REVIEWER_BATCH_SIZE` | 5 | SA 배치 크기 |
| `REPORT_REVIEWER_IMAGE_MAX_DIM` | 1928 | 슬라이드 이미지 가로 px |
| `REPORT_REVIEWER_PDF_DPI` | 150 | 임베디드 PDF → PNG DPI |
| `REPORT_REVIEWER_THUMBNAIL_MAX_DIM` | 480 | 썸네일 최대 px |

## 배포

```powershell
# skill_src/ → ~/.claude/skills/report-reviewer/ (robocopy, .venv 제외)
# agents_src/*.md → ~/.claude/agents/
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
./deploy.ps1

# 배포 후 글로벌 위치에 venv 생성
cd "$env:USERPROFILE/.claude/skills/report-reviewer"
python -m venv .venv
.venv\Scripts\pip install -e .
```

배포 후 Claude Code 재시작 필요.

## 주의사항

- `tests/fixtures/sample_with_embedded.pptx` 미존재 → `test_extract_detects_embedded` 자동 skip (정상)
- `skill_src/` 외부에서 pytest 실행 시 상위 디렉토리의 `pyproject.toml`을 탐색할 수 있어 설정 혼동 가능 — 반드시 `skill_src/` 안에서 실행
- Pillow 저장 포맷은 `"JPEG"` (대문자, `"JPG"` 아님)
- reporter_html의 `_TEMPLATES_DIR`은 `src/` 기준 `../templates/` — 배포 후 경로 유지됨
