---
name: report-reviewer
description: 사내 PPTX 보고서를 작성자가 자가점검하는 도구. 오타·용어 통일·데이터·결론·개선·논리 6개 카테고리로 검토. 사용법 `/report-reviewer <pptx_path>` 또는 `/report-reviewer <pptx_path> --resume|--rerun-stage2|--rerender`.
---

# report-reviewer Skill

## 사용 시점
사용자가 다음 형식으로 호출:
- `/report-reviewer C:/path/to/report.pptx` — 처음부터
- `/report-reviewer ... --resume` — 중간 산출물 보존되어 있으면 이어서
- `/report-reviewer ... --rerun-stage2` — 1단계 결과 재사용, 2단계만
- `/report-reviewer ... --rerender` — findings.json 재사용, 리포트만

## Claude(메인)의 워크플로

### Step 0: 환경 확인
- Skill 폴더 위치: `~/.claude/skills/report-reviewer/`
- 가상환경: `~/.claude/skills/report-reviewer/.venv/Scripts/python` (없으면 사용자에게 설치 안내)
- 작업 디렉토리: 입력 PPTX와 같은 폴더에 `review_ws/` 생성

### Step 1: 추출 + 이미지 변환
```bash
"<skill>/.venv/Scripts/python" "<skill>/extract_and_render.py" "<pptx_path>" --out "<pptx_dir>/review_ws"
```
- `--resume`이면 `review_ws/extracted.json` 이미 존재 시 이 단계 건너뜀
- 산출물: `review_ws/extracted.json`, `review_ws/slides/slide_NNN.jpg`, `review_ws/thumbnails/slide_NNN.jpg`, `review_ws/slide_inputs/slide_NNN.json`

### Step 2: 1단계 SA 배치 dispatch (슬라이드별 분석)
- `extracted.json`의 `slide_count = N`
- 환경변수 `REPORT_REVIEWER_BATCH_SIZE` (기본 5)로 한 번에 dispatch할 SA 수 결정
- 각 슬라이드 i에 대해:
  - Task 도구로 `report-reviewer-slide-analyzer` subagent dispatch
  - prompt에 슬라이드별 입력 파일 경로 + 이미지 경로 포함
- 각 SA 응답에서 JSON 코드 블록 추출 → 누적
- 모든 결과를 `review_ws/slide_summaries.json`에 저장

### Step 3: 2단계 SA 동시 dispatch (6개 카테고리 + 거시 1개 = 7개)
- 7개 SA를 동시 dispatch:
  - **카테고리 SA (6개, 슬라이드 단위 미시 검토)** — 결과는 배열 → `findings[]`로 concat
    - `report-reviewer-typo` (ID prefix `T`)
    - `report-reviewer-terminology` (ID prefix `TM`)
    - `report-reviewer-data` (ID prefix `D`)
    - `report-reviewer-conclusion` (ID prefix `C`)
    - `report-reviewer-improvement` (ID prefix `I`)
    - `report-reviewer-logic` (ID prefix `L`)
  - **거시 SA (1개, 문서 전체 검토)** — 결과는 객체 → `document_review`로 저장
    - `report-reviewer-document` (출력: `thesis_question`, `overall_grade`, 5개 항목 평가, `cross_slide_concerns[]` 등)
- 6개 카테고리 SA 응답을 concat하여 `findings[]` 생성
- document SA 응답(객체 1개)을 `document_review`에 저장
- 심각도·카테고리 통계 집계 후 `review_ws/findings.json` 저장 (스키마: `summary`, `findings[]`, `document_review`)

### Step 4: 리포트 렌더
```bash
"<skill>/.venv/Scripts/python" "<skill>/render_report.py" "<pptx_dir>/review_ws" --out "<pptx_dir>/review_output"
```
- 산출물: `review.md`, `review.html`, `assets/`

### Step 5: 사용자에게 결과 안내
- review.md 위치 안내
- Critical 이슈 위주 핵심 요약 (3~5건)
- HTML 리포트 열기 가이드 (`start review_output/review.html` 등)

## 에러 처리
- **PowerPoint COM 미설치 / Windows 아님**: Step 1 실패 시 친절 메시지로 종료
- **이미지 토큰 한계 초과**: SA dispatch 시 한계 에러 catch → 슬라이드 PNG를 1280px → 960px로 다운스케일 재시도
- **Subagent 실패**: 1회 재시도, 그래도 실패 시 placeholder 결과로 채우고 리포트에 명시 (카테고리 SA 실패 시 빈 배열 `[]`, document SA 실패 시 `document_review` 키 생략)

## 옵션 처리
- `--resume`: extracted.json 있으면 Step 1 건너뜀, slide_summaries.json 있으면 Step 2도 건너뜀
- `--rerun-stage2`: extracted.json + slide_summaries.json 사용, Step 3부터
- `--rerender`: findings.json 사용, Step 4만
