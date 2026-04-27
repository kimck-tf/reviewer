# report-reviewer — 사내 보고서 자가점검 도구 설계

작성일: 2026-04-27
상태: Draft (사용자 검토 대기)

## 1. 개요

### 1.1 목적
사내 보고서(주로 차량 CAE 해석 결과 보고서, 기타 기술용역·개발계획 보고서)를 작성자 본인이 작성 중·제출 전에 자가점검할 수 있는 Claude Code 기반 도구.

### 1.2 1차 사용자·시나리오
- **사용자**: 보고서 작성자 본인 (주로 차량 CAE 해석 엔지니어)
- **사용 시점**: 작성 중·제출 직전
- **사용 방식**: Claude Code CLI에서 `/report-reviewer <pptx_path>` 호출
- **강조점**: 즉시 피드백, 위치 추적이 명확한 개선 제안

### 1.3 입력·출력
- **입력**: PPTX 파일 1개
- **출력**:
  - `<work_dir>/extracted.json` — 추출 결과 중간 산출물
  - `<work_dir>/slides/*.jpg`, `<work_dir>/thumbnails/*.jpg` — 슬라이드 이미지
  - `<work_dir>/slide_summaries.json`, `<work_dir>/findings.json` — Subagent 분석 결과
  - `<output_dir>/review.md` — 마크다운 리포트 (Claude Code 콘솔에서 열람)
  - `<output_dir>/review.html` — HTML 리포트 (썸네일 + 위치 박스 오버레이 + 심각도 색상)

## 2. 요구사항

### 2.1 1차 MVP 기능 (LLM이 모두 검증)

| # | 카테고리 | 검증 내용 |
|---|---|---|
| 1 | 오타 (typo) | 오타·맞춤법·표기 오류 |
| 2 | 용어·맥락 통일성 (terminology) | 용어·약어·표기 일관성, 한·영 혼용 일관성 |
| 3 | 데이터 검토 (data) | 표·본문 수치 정합성, 단위 누락, 인용 일치성 |
| 4 | 결론 검증 (conclusion) | 결론이 데이터·자료로 뒷받침되는지 |
| 5 | 개선 제안 (improvement) | 정보 전달·결론 뒷받침을 위한 개선 제안 |
| 6 | 결론 강도·논리 (logic) | overclaim 검출, 일반화 오류, 가정·한계 명시 |

### 2.2 1차 MVP 제외 (2차 이후)
- CAE 도메인 특화 검증(단위계, 좌표계, 모델 메타정보, Pass/Fail 판정 등) — 기술용역·개발계획 보고서에도 적용 가능한 범용성을 1차에서 우선
- PPTX 코멘트 직접 패치, 자동 수정 PR 생성 등

### 2.3 비기능 요구사항
- **PPT 수정 위치 추적성 필수** — 슬라이드 번호 + 슬라이드 제목 + 도형 식별자 + 위치 hint + 텍스트 인용
- **재실행·resume 가능** — 단계별 중간 산출물 보존
- **재사용성** — 사용자 글로벌 배포 (`~/.claude/skills/`, `~/.claude/agents/`)

## 3. 아키텍처

### 3.1 분업 원칙
- **Python (결정론적)**: 추출 + 슬라이드 PNG 변환 + 리포트 렌더만 담당
- **Claude (메인 + Subagents)**: 모든 추론·판단·이미지 인식·종합 담당
- **Gemini 등 외부 LLM 사용 안 함** — Claude Sonnet 4.6 (1M)의 multimodal 능력 활용

### 3.2 패키징 형태
- **Skill + Claude Code 표준 Subagent 혼합**
- Skill: `~/.claude/skills/report-reviewer/`
- Subagent: `~/.claude/agents/report-reviewer-*.md`

### 3.3 데이터 흐름

```
사용자: /report-reviewer <pptx_path>
  │
  ▼
Claude(메인) — SKILL.md 가이드 따름
  │
  ▼ Bash 호출
extract_and_render.py <pptx> --out <work_dir>
  ├─ python-pptx로 슬라이드별 텍스트·표·도형·위치·노트·메타 추출
  └─ 슬라이드 → 이미지 변환 (자동 모드 선택: 임베디드 OLE 감지 시 모드 2)
  → <work_dir>/extracted.json
  → <work_dir>/slides/*.jpg
  → <work_dir>/thumbnails/*.jpg
  │
  ▼
Claude(메인) — 1단계 SA 배치 dispatch (Task 도구)
  │ report-reviewer-slide-analyzer × N (배치 5개씩)
  │   각 SA 입력: 슬라이드 i의 extracted 데이터 + slide_i.jpg (Read multimodal)
  │   각 SA 출력: 구조화된 슬라이드 분석 결과
  │ → <work_dir>/slide_summaries.json
  │
  ▼
Claude(메인) — 2단계 SA 동시 dispatch (6개)
  │ report-reviewer-{typo,terminology,data,conclusion,improvement,logic}
  │   각 SA 입력: slide_summaries + extracted.json 참조 권한
  │   각 SA 출력: findings[] (자기 카테고리)
  │ → <work_dir>/findings.json (집계)
  │
  ▼ Bash 호출
render_report.py <work_dir> --out <output_dir>
  → <output_dir>/review.md
  → <output_dir>/review.html (썸네일 + 위치 박스 + 심각도 색상)
  → <output_dir>/assets/
  │
  ▼
Claude(메인) — 사용자에게 위치 안내 + 핵심 요약 (Critical 우선)
```

## 4. 디렉토리 구조

### 4.1 Skill 본체

```
~/.claude/skills/report-reviewer/
├── SKILL.md                          # Claude(메인)의 진입점·워크플로 가이드
├── README.md                         # 사용자용 설치·사용법
├── pyproject.toml                    # 의존성 명세
├── extract_and_render.py             # CLI: 추출 + 이미지 변환
├── render_report.py                  # CLI: findings.json → MD/HTML
├── src/
│   ├── __init__.py
│   ├── config.py
│   ├── extractor.py
│   ├── slide_renderer.py
│   ├── locator.py
│   ├── reporter_md.py
│   └── reporter_html.py
├── templates/
│   ├── report.html.j2
│   └── style.css
└── tests/
    ├── conftest.py
    ├── fixtures/
    ├── unit/
    ├── integration/
    └── llm/
```

### 4.2 Subagent

```
~/.claude/agents/
├── report-reviewer-slide-analyzer.md       # 1단계
├── report-reviewer-typo.md                 # 2단계 × 6
├── report-reviewer-terminology.md
├── report-reviewer-data.md
├── report-reviewer-conclusion.md
├── report-reviewer-improvement.md
└── report-reviewer-logic.md
```

## 5. 컴포넌트 상세

### 5.1 Python 모듈

| 모듈 | 책임 | 핵심 인터페이스 |
|---|---|---|
| `config.py` | 환경변수, 경로, 기본값 dataclass | `Config` |
| `extractor.py` | python-pptx 추출, 임베디드 OLE 감지 | `extract(pptx_path) -> dict` |
| `slide_renderer.py` | 슬라이드 → 이미지 변환 (자동 모드 선택) | `render(pptx_path, out_dir) -> Dict[int, List[Path]]` |
| `locator.py` | EMU 좌표 → 한국어 hint 변환 | `to_hint(position) -> str` |
| `reporter_md.py` | findings.json → Markdown | `render(findings, extracted, out_path) -> Path` |
| `reporter_html.py` | findings.json → HTML (Jinja2) | `render(findings, extracted, out_dir) -> Path` |

### 5.2 슬라이드 → 이미지 변환 (slide_renderer.py)

사용자가 검증한 PowerPoint COM 기반 방법 채택 (참고: `8. Reviewer/reference/pptx추출.png`).

| 모드 | 함수 | 트리거 |
|---|---|---|
| 모드 1 | `convert_pptx_to_images_without_embedding` | 임베디드 OLE 없는 슬라이드 (기본) |
| 모드 2 | `convert_pptx_with_embedded_to_images` | 임베디드 OLE 감지 시 (자동 fallback) |

- 해상도: 1928×1080 (검증된 값)
- 포맷: JPG (모드 1), PNG (모드 2의 PDF 경유)
- 자동 모드 선택은 `slide_renderer.py` 내부에 캡슐화 (SKILL은 신경 쓰지 않음)
- 임베디드 PDF/PPT/Excel은 별도 이미지로 추출하여 1단계 SA에 함께 전달

### 5.3 CLI 진입점

```bash
# Step 1+2: 추출 + 이미지 변환
python extract_and_render.py <pptx_path> --out <work_dir>

# Step 4: 리포트 렌더
python render_report.py <work_dir> --out <output_dir>
```

옵션 (SKILL이 처리):
```bash
/report-reviewer <pptx>                    # 처음부터
/report-reviewer <pptx> --resume           # work_dir 있으면 이어서
/report-reviewer <pptx> --rerun-stage2     # 1단계 재사용, 2단계만
/report-reviewer <pptx> --rerender         # findings.json 재사용, 리포트만
```

### 5.4 Subagent 정의

#### 5.4.1 1단계: report-reviewer-slide-analyzer

```yaml
---
name: report-reviewer-slide-analyzer
description: PPTX 슬라이드 1장의 텍스트·표·이미지를 multimodal 분석하여 핵심 내용·주장·데이터 포인트를 구조화. report-reviewer skill에서만 사용.
tools: Read
---
```

**역할**: 슬라이드의 (1) 추출된 텍스트·표 데이터(JSON) + (2) 슬라이드 이미지를 함께 보고 다음을 구조화 출력:
1. 슬라이드 핵심 메시지 (1~2문장)
2. 핵심 주장(claims) 리스트
3. 핵심 데이터 포인트 [(값, 단위, 맥락)]
4. 그림·차트·컨투어 등 시각 정보 관찰 (이미지에서만 보이는 정보)
5. 발표자 노트 요약 (있는 경우)
6. 다른 슬라이드와의 잠재적 연결점 힌트

**출력 형식**: `slide_summary` JSON 스키마 (5.5.2 참조).
**원칙**: 원문 텍스트(JSON)와 이미지 OCR 결과가 다르면 원문을 우선 신뢰.

#### 5.4.2 2단계: 카테고리별 6개 Subagent

| Subagent | 역할 |
|---|---|
| `report-reviewer-typo` | 오타·맞춤법·표기 오류 (원문 텍스트 기반) |
| `report-reviewer-terminology` | 용어·약어·표기 통일성, 한·영 혼용 일관성 |
| `report-reviewer-data` | 표·본문 수치 정합성, 단위 누락, 인용 일치성 |
| `report-reviewer-conclusion` | 결론이 데이터·자료로 뒷받침되는지 |
| `report-reviewer-improvement` | 정보 전달·결론 뒷받침을 위한 개선 제안 |
| `report-reviewer-logic` | overclaim, 일반화 오류, 가정·한계 명시 |

공통 입출력:
- **Input**: `slide_summaries[]` + extracted.json 참조 권한 (Read tool)
- **Output**: `findings[]` JSON 배열
- **Tools**: `Read`

### 5.5 데이터 구조

#### 5.5.1 extracted.json (Python → SA 입력)

```json
{
  "metadata": {
    "file_path": "...",
    "title": "...",
    "author": "...",
    "created": "...",
    "modified": "...",
    "slide_count": 30
  },
  "slides": [
    {
      "index": 1,
      "title": "표지",
      "layout": "Title Slide",
      "shapes": [
        {
          "shape_id": "s1_sh1",
          "type": "TextBox|Title|Table|Picture|Chart|Group|EmbeddedOLE",
          "position_emu": {"left": 914400, "top": 685800, "width": ..., "height": ...},
          "position_pct": {"left": 0.10, "top": 0.07, "width": 0.80, "height": 0.10},
          "z_order": 1,
          "text": "텍스트 박스/제목 안 텍스트",
          "table": {"rows": 4, "cols": 5, "cells": [[...]]},
          "image_ref": "media/image3.png",
          "embedded_progid": "PowerPoint.Show.12"
        }
      ],
      "notes": "발표자 노트 본문",
      "image_path": "<work_dir>/slides/slide_001.jpg",
      "thumbnail_path": "<work_dir>/thumbnails/slide_001.jpg",
      "embedded_image_paths": ["<work_dir>/slides/embedded/slide_005_emb01.png"],
      "has_embedded": false
    }
  ]
}
```

#### 5.5.2 slide_summaries.json (1단계 SA → 2단계 SA 입력)

```json
{
  "slides": [
    {
      "index": 1,
      "key_message": "프론트 서스펜션 강도 해석 결과 요약",
      "claims": ["스트럿 마운트 안전계수 1.5 이상 만족"],
      "data_points": [{"value": 250, "unit": "MPa", "context": "최대 응력"}],
      "vision_observations": ["컨투어 이미지 좌측에 hot spot 위치"],
      "notes_summary": "...",
      "cross_slide_hints": ["슬라이드 8의 평가 기준과 비교 필요"]
    }
  ]
}
```

#### 5.5.3 findings.json (2단계 SA → 리포트)

**SSoT 정책**: `slide_summaries.json`이 단일 진실 소스(SSoT). `findings.json`에는 `slide_summaries`를 중복 포함하지 않음 (참조만). 리포트 렌더는 두 파일 모두 로드.

```json
{
  "summary": {
    "total_issues": 42,
    "by_severity": {"critical": 5, "warning": 20, "info": 17},
    "by_category": {"typo": 8, "terminology": 5, "data": 12, "conclusion": 6, "improvement": 7, "logic": 4}
  },
  "findings": [
    {
      "id": "F001",
      "category": "data",
      "severity": "critical",
      "slide_index": 5,
      "shape_id": "s5_sh3",
      "position_hint": "슬라이드 5 우측 상단 텍스트 박스 (좌측 70%, 상단 15%)",
      "position_pct": {"left": 0.70, "top": 0.15, "width": 0.25, "height": 0.10},
      "quoted_text": "최대 응력 250 MPa",
      "issue": "표 5의 셀에는 240 MPa로 기재되어 있어 본문 인용과 불일치",
      "suggestion": "본문을 240 MPa로 수정 또는 표 데이터 재확인",
      "evidence": "slide_5 표(s5_sh4) 행 3·열 2 = '240'"
    }
  ]
}
```

**`position_pct` 채움 책임**: 2단계 카테고리 SA는 `shape_id`로 `extracted.json`의 해당 도형을 조회하여 그 `position_pct`를 finding에 복사한다. HTML 리포트의 슬라이드 썸네일 위에 위치 박스를 그릴 때 사용. 도형을 특정할 수 없는 finding(예: 슬라이드 전체 결론)은 `position_pct`를 생략 가능 — 리포트는 박스 없이 카드만 표시.

## 6. 의존성·환경 요구사항

### 6.1 Python 의존성

```toml
[project]
dependencies = [
  "python-pptx>=0.6.23",
  "Pillow>=10.0",
  "Jinja2>=3.1",
  "pywin32>=306",        # Windows + PowerPoint COM
  "pdf2image>=1.17",     # 모드 2의 PDF 임베디드 처리 시
]
```

### 6.2 환경 요구사항

| 항목 | 요구 | 비고 |
|---|---|---|
| OS (런타임) | Windows | PowerPoint COM 의존 |
| OS (단위 테스트 L1) | Windows / Linux 모두 | COM은 mock 처리, 10.4절 참조 |
| MS PowerPoint | 설치 필수 (런타임) | win32com.client 사용 |
| Poppler for Windows | 모드 2 사용 시 필요 | pdf2image 백엔드, PATH 등록 |
| Claude Code CLI | 필수 | Skill·Subagent 동작 환경 |

## 7. 에러 처리

| 에러 유형 | 처리 |
|---|---|
| PowerPoint COM 미설치 | 시작 시 명확한 에러 메시지로 종료 |
| Poppler 미설치 | 임베디드 PDF 발견 시에만 경고, 해당 객체 스킵 |
| PPTX 손상·암호 | extractor 단계에서 잡고 종료 |
| 슬라이드 변환 실패 1장 | 텍스트 데이터로만 분석, 리포트에 명시 |
| Subagent 1단계 실패 | 1회 재시도 후 placeholder summary 사용 |
| Subagent 2단계 실패 | 1회 재시도 후 해당 카테고리만 누락, 리포트에 명시 |
| 이미지 토큰 한계 초과 | 슬라이드를 1928→1280→960 단계로 다운스케일 재시도 |
| 사용자 중단 | 중간 산출물 보존, `--resume` 옵션으로 재개 가능 |

## 8. 비용·성능 최적화

| 최적화 | 방법 |
|---|---|
| 1단계 결과 캐싱 | `slide_summaries.json` 저장, 재실행 시 재사용 |
| 2단계 입력 압축 | 2단계 SA에는 슬라이드 이미지 전달 안 함, 1단계 텍스트 요약만 |
| 2단계 Read 권한 | 필요 시 원문 인용을 위해 extracted.json 일부 Read 가능 |
| 썸네일 분리 | Subagent에는 원본 이미지, 썸네일은 HTML 리포트 전용 |
| 동시 실행 | 1단계 배치 5개씩 (환경변수 `REPORT_REVIEWER_BATCH_SIZE`로 조절), 2단계 6개 동시 1라운드 |

## 9. 재실행·resume

| 옵션 | 동작 |
|---|---|
| (없음) | 처음부터 |
| `--resume` | work_dir 존재 시 마지막 완료 단계부터 |
| `--rerun-stage2` | extracted.json + slide_summaries.json 재사용, 2단계만 |
| `--rerender` | findings.json 재사용, 리포트만 |

## 10. 테스트 전략

### 10.1 4계층

| 계층 | 대상 | 방법 |
|---|---|---|
| L1 단위 | Python 모듈 | pytest + mock (COM mock) |
| L2 통합 | CLI end-to-end | fixture PPTX 4종 실제 실행 |
| L3 SA Smoke + 골든 | 각 Subagent 출력 형식·정확도 | LLM 실호출, 골든 파일 비교 |
| L4 사용자 acceptance | 실제 사내 보고서 | 사용자 직접 |

### 10.2 Fixture PPTX

| Fixture | 검증 목적 |
|---|---|
| `sample_text_only.pptx` | 가장 단순 |
| `sample_with_table.pptx` | 표 추출 정확성 |
| `sample_with_image.pptx` | 슬라이드 PNG 변환 정확성 |
| `sample_with_embedded.pptx` | 모드 2 자동 fallback 트리거 |

### 10.3 정확도 메트릭 (L3)

| 메트릭 | 목표 |
|---|---|
| Precision | ≥ 0.7 |
| Recall | ≥ 0.6 |
| False positive rate | ≤ 0.3 |

### 10.4 CI 제약

- L1만 자동화 (Linux runner, COM mock)
- L2는 사용자 로컬 (Windows + MS Office 필요)
- L3는 수동·정기 (LLM 비용)

## 11. 향후 확장 (1차 MVP 외)

- CAE 도메인 특화 검증 (단위계, 좌표계, 모델 메타정보, Pass/Fail 판정 등)
- 표·그림 무결성 (캡션 누락, 본문 인용 정합성)
- 보안·민감정보 검출
- 이전 버전 보고서 자동 diff
- 동일 차종 과거 보고서 결과값 이상치 탐지
- PPTX 코멘트로 검토 결과 직접 패치
- 사내 표준 양식 준수 자동 점검표

## 12. 미결정·확인 필요 사항 (plan 단계에서 결정)

| 항목 | 현 상태 | 결정 시점 |
|---|---|---|
| 샘플 PPTX 제공 여부 | 추후 추가 (옵션 C) | 사용자 추가 시 fixture로 편입 |
| 1단계 배치 크기 기본값 | 5 (조절 가능) | 실측 후 튜닝 |
| Subagent 시스템 프롬프트 세부 문구 | 구현 단계에서 | 골든 파일로 회귀 검증 |
| 1단계 SA 입력 전달 방식 | (미정) | plan 단계 — Task prompt 인라인 JSON vs 임시 파일 경로 전달 |
| 2단계 SA의 extracted.json Read 범위 | "필요 시 일부" | plan 단계 — 슬라이드 단위 분할 파일 vs 전체 파일 |
| 6개 카테고리 SA 결과 병합 로직 | 단순 concat 가정 | plan 단계 — ID 부여 규칙·충돌 해결 명시 |
| 이미지 토큰 한계 감지 방법 | (미정) | plan 단계 — API 에러 코드 vs 사전 측정 |
| 정확도 메트릭 측정 방법론 | (미정) | L3 테스트 작성 시 — 라벨링 방법·측정 단위 |
