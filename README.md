# 8. Reviewer

사내 PPTX 보고서를 작성자가 제출 전 자가점검하는 Claude Code Skill.

`/report-reviewer <pptx>` 한 번으로 슬라이드 단위 미시 검토(오타·용어·데이터·결론·개선·논리)와 문서 전체 거시 검토(핵심 질문 답변·스토리라인·결정 정보·청중 적합성·슬라이드 간 모순)를 동시에 수행하고, 위치 박스 오버레이가 포함된 HTML 리포트로 결과를 돌려준다.

## 구성

```
PPTX → [Python: 추출+이미지] → [Claude SA: 8개 동시 분석] → [Python: 리포트]
       extract_and_render.py     slide-analyzer 1개            render_report.py
                                 + 카테고리 SA 6개
                                 + document SA 1개 (거시)
```

### 검토 카테고리 (1 + 7개 subagent)

**1단계 — multimodal 슬라이드 분석 (슬라이드 N장 × 1개 SA)**

| SA | 역할 |
|---|---|
| `report-reviewer-slide-analyzer` | 슬라이드 1장의 이미지·텍스트·표를 함께 보고 핵심 메시지·주장·데이터 포인트로 구조화 |

**2단계 — 7개 SA 동시 dispatch**

| SA | 시야 | 결과 |
|---|---|---|
| `report-reviewer-typo` | 슬라이드 단위 | 오타·맞춤법 (ID `T`) |
| `report-reviewer-terminology` | 슬라이드 단위 | 용어 통일 (ID `TM`) |
| `report-reviewer-data` | 슬라이드 단위 | 수치 정합성·단위·자릿수 (ID `D`) |
| `report-reviewer-conclusion` | 슬라이드 단위 | 결론 뒷받침 근거 (ID `C`) |
| `report-reviewer-improvement` | 슬라이드 단위 | 정보 전달·명확성 개선 (ID `I`) |
| `report-reviewer-logic` | 슬라이드 단위 | overclaim·논리 비약 (ID `L`) |
| `report-reviewer-document` | **문서 전체** | 핵심 질문 답변·구성 균형·결정 정보·청중 적합성·슬라이드 간 모순, 종합 등급 |

## 환경 요구사항

- Windows + MS PowerPoint 설치 (슬라이드 → 이미지 변환에 COM 사용)
- Python 3.10+
- Claude Code CLI
- (선택) Poppler for Windows — 임베디드 PDF 처리 시

## 빠른 시작

### 개발용 설치

```powershell
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
python -m venv .venv
.venv\Scripts\pip install -e ".[dev]"
.venv\Scripts\pytest tests/unit/ -v
```

### 글로벌 배포 (Skill로 등록)

```powershell
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
./deploy.ps1

cd "$env:USERPROFILE/.claude/skills/report-reviewer"
python -m venv .venv
.venv\Scripts\pip install -e .
```

배포 후 Claude Code 재시작.

### 사용

Claude Code에서:

```
/report-reviewer C:/path/to/report.pptx
```

옵션:
- `--resume`: 중간 산출물 보존되어 있으면 이어서
- `--rerun-stage2`: 1단계 SA 결과 재사용, 2단계만 재실행
- `--rerender`: `findings.json` 재사용, 리포트만 재생성

## 산출물

입력 PPTX와 같은 폴더에 다음이 생성된다.

### `review_ws/` (중간 산출물)

| 파일 | 단계 | 내용 |
|---|---|---|
| `extracted.json` | Step 1 | 슬라이드 메타데이터·도형·텍스트·표 |
| `slides/slide_NNN.jpg` | Step 1 | 슬라이드 원본 이미지 |
| `thumbnails/slide_NNN.jpg` | Step 1 | 썸네일 (HTML 리포트용) |
| `slide_inputs/slide_NNN.json` | Step 1 | 1단계 SA 입력 (슬라이드별) |
| `slide_summaries.json` | Step 2 | 1단계 SA 누적 결과 |
| `findings.json` | Step 3 | 2단계 7개 SA 결과 (`findings[]` + `document_review`) |

### `review_output/` (최종 리포트)

- `review.md` — 마크다운 리포트 (맨 앞에 "문서 전체 평가" 섹션 → 슬라이드별·카테고리별 이슈)
- `review.html` — HTML 리포트 (종합 등급 배지, 핵심 질문 답변, 5개 축 평가, 슬라이드 간 모순, 슬라이드별 썸네일에 위치 박스 오버레이)
- `assets/` — CSS·썸네일

## 디렉토리 구조

```
8. Reviewer/
├── CLAUDE.md            # 이 저장소에서 작업하는 Claude Code에 대한 가이드
├── agents_src/          # 8개 subagent 시스템 프롬프트
├── skill_src/           # Skill 본체 (Python 코드·SKILL.md·테스트)
│   ├── SKILL.md         # 메인 Claude의 워크플로 (Step 0~5)
│   ├── src/             # 추출·렌더·리포터 모듈
│   ├── templates/       # HTML 리포트 Jinja2 + CSS
│   ├── tests/           # 단위·통합 테스트
│   └── deploy.ps1       # 글로벌 배포 스크립트
├── docs/                # 설계·계획 문서
└── reference/           # 참고 자료
```

## 추가 문서

- [CLAUDE.md](./CLAUDE.md) — 아키텍처·테스트·환경변수·주의사항 등 개발자용 상세 가이드
- [skill_src/SKILL.md](./skill_src/SKILL.md) — 메인 Claude가 `/report-reviewer` 호출 시 따르는 단계별 워크플로
- [skill_src/README.md](./skill_src/README.md) — Skill 패키지 자체 설치·테스트 가이드
