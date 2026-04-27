# report-reviewer Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 사내 PPTX 보고서를 작성자가 자가점검할 수 있는 Claude Code Skill + 7개 Subagent를 구현한다. PPTX 추출·이미지 변환·리포트 렌더는 Python으로, 모든 검증·판단·이미지 인식은 Claude (메인 + Subagents)가 수행한다.

**Architecture:** 3단 파이프라인 — ① Python(추출 + PNG 변환) → ② Claude(1단계 SA: 슬라이드별 multimodal 분석, 2단계 SA: 6개 카테고리 검증) → ③ Python(MD/HTML 리포트). Skill은 Claude(메인)의 워크플로 가이드, Subagent는 Task 도구로 dispatch.

**Tech Stack:** Python 3.10+, python-pptx, Pillow, Jinja2, pywin32 (PowerPoint COM), pdf2image + Poppler (모드 2 임베디드 PDF 처리), pytest. Windows + MS PowerPoint 런타임 필수.

**Spec:** `C:\_PYTHON\0_CK_Project\8. Reviewer\docs\superpowers\specs\2026-04-27-report-reviewer-design.md`

## 실행 환경 가정

- 모든 Bash 명령은 **절대 경로 또는 명시적 `cd` 후 실행**. 셸 cwd는 매번 reset 가능성 있음.
- 작업 루트: `C:/_PYTHON/0_CK_Project/8. Reviewer/`
- Python 진입은 가상환경의 `.venv/Scripts/python` 절대 경로 사용 (Windows).
- `git add .` 금지 — 매 commit 시 수정 파일을 개별 명시. 첫 commit도 동일 원칙 (예외 없음).

## Spec 미결정 사항 ↔ Chunk 매핑 (forward reference)

| Spec §12 항목 | 다뤄지는 Chunk |
|---|---|
| 1단계 SA 입력 전달 방식 (인라인 JSON vs 임시 파일) | **Chunk 4** (SKILL.md 워크플로) |
| 2단계 SA의 extracted.json Read 범위 | **Chunk 4** (Subagent 정의) |
| 6개 카테고리 SA 결과 병합 로직 | **Chunk 4** (SKILL.md) |
| 이미지 토큰 한계 감지 방법 | **Chunk 4** (SKILL.md 에러 처리) |
| 정확도 메트릭 측정 방법론 | **Chunk 5 (선택)** — 사용자 샘플 추가 후 |

## Chunk 분할 (총 4 + 선택 1)

| Chunk | 범위 | 예상 파일 수 |
|---|---|---|
| **Chunk 1** | Phase 0~2: 셋업 + extractor + locator | 6 |
| **Chunk 2** | Phase 3: slide_renderer (PowerPoint COM) | 2 |
| **Chunk 3** | Phase 4~5: reporter_md + reporter_html | 4 |
| **Chunk 4** | Phase 6~7: CLI + SKILL.md + 7 Subagent + 글로벌 배포 | 11 |
| **Chunk 5 (선택)** | 골든 파일 정확도 메트릭 (사용자 샘플 추가 후) | 2 |

---

## File Structure

### 개발 디렉토리 (작업 위치, git 추적)

```
C:\_PYTHON\0_CK_Project\8. Reviewer\
├── .git/                                    # git init 후
├── .gitignore                               # __pycache__, *.pyc, review_ws/, dist/
├── pyproject.toml                           # 패키지 메타·의존성
├── README.md                                # 프로젝트 설명
├── docs/                                    # 기존 spec/plan
│   └── superpowers/
│       ├── specs/...
│       └── plans/...
├── reference/                               # 기존 참고 자료
│   └── pptx추출.png
├── skill_src/                               # 배포 전 Skill 본체 (글로벌로 복사 대상)
│   ├── SKILL.md
│   ├── README.md
│   ├── pyproject.toml
│   ├── extract_and_render.py                # CLI 1
│   ├── render_report.py                     # CLI 2
│   ├── src/
│   │   ├── __init__.py
│   │   ├── config.py
│   │   ├── extractor.py
│   │   ├── slide_renderer.py
│   │   ├── locator.py
│   │   ├── reporter_md.py
│   │   └── reporter_html.py
│   ├── templates/
│   │   ├── report.html.j2
│   │   └── style.css
│   └── tests/
│       ├── conftest.py
│       ├── fixtures/
│       │   ├── sample_text_only.pptx
│       │   ├── sample_with_table.pptx
│       │   ├── sample_with_image.pptx
│       │   ├── sample_with_embedded.pptx
│       │   └── golden/
│       │       ├── extractor_outputs/
│       │       └── locator_outputs.json
│       ├── unit/
│       │   ├── test_config.py
│       │   ├── test_extractor.py
│       │   ├── test_locator.py
│       │   ├── test_slide_renderer.py
│       │   ├── test_reporter_md.py
│       │   └── test_reporter_html.py
│       └── integration/
│           ├── test_cli_extract_and_render.py
│           └── test_cli_render_report.py
└── agents_src/                              # 배포 전 Subagent 정의 (글로벌로 복사 대상)
    ├── report-reviewer-slide-analyzer.md
    ├── report-reviewer-typo.md
    ├── report-reviewer-terminology.md
    ├── report-reviewer-data.md
    ├── report-reviewer-conclusion.md
    ├── report-reviewer-improvement.md
    └── report-reviewer-logic.md
```

### 배포 후 (글로벌 위치, 사용자 사용)

```
C:\Users\zenit\.claude\skills\report-reviewer\           # skill_src/ 복사본
C:\Users\zenit\.claude\agents\report-reviewer-*.md       # agents_src/*.md 복사본
```

### 파일별 책임

| 파일 | 책임 |
|---|---|
| `pyproject.toml` (skill_src) | Skill용 패키지 메타·의존성 (배포 시 함께 감) |
| `pyproject.toml` (개발 루트) | 개발용 (pytest 설정 등) — 또는 skill_src 것을 재사용 |
| `extract_and_render.py` | CLI 진입점 1: PPTX → extracted.json + 이미지 |
| `render_report.py` | CLI 진입점 2: findings.json + extracted.json → MD/HTML |
| `src/config.py` | 환경변수·기본값·dataclass (`Config`) |
| `src/extractor.py` | python-pptx로 PPTX 구조·텍스트·표·도형·노트·메타·임베디드 OLE 감지 → `extract(pptx_path) -> dict` |
| `src/slide_renderer.py` | PowerPoint COM으로 슬라이드 → 이미지 변환 (자동 모드 선택) → `render(pptx_path, out_dir) -> Dict[int, List[Path]]` |
| `src/locator.py` | EMU 좌표 → 한국어 위치 hint 변환 (순수 함수) → `to_hint(position_pct) -> str` |
| `src/reporter_md.py` | findings.json → 마크다운 리포트 → `render(findings, extracted, out_path) -> Path` |
| `src/reporter_html.py` | findings.json → HTML 리포트 (Jinja2, 썸네일 + 위치 박스 오버레이) → `render(findings, extracted, out_dir) -> Path` |
| `templates/report.html.j2` | HTML 리포트 템플릿 |
| `templates/style.css` | HTML 스타일 (심각도 색상 등) |
| `SKILL.md` | Claude(메인)가 따를 워크플로 가이드 (frontmatter + 본문) |
| `agents_src/*.md` | 7개 Subagent의 시스템 프롬프트 + 도구 권한 |

---

## Chunk 1: Phase 0~2 (셋업 + extractor + locator)

**[Chunk 1 cwd 면책]** 이 chunk의 모든 Bash 명령은 명시적 `cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"` 직후 실행을 가정한다. Task 0.1만 작업 루트(`C:/_PYTHON/0_CK_Project/8. Reviewer`)에서 실행하며 명시적으로 cd를 표기한다. 셸 cwd가 reset되거나 새 터미널 세션이면 매번 `cd`로 진입할 것. 절대 경로 사용이 안전한 명령(첫 venv 생성, 골든 파일 생성 등)은 본문에 절대 경로로 명시되어 있다.

### Task 0.1: git 초기화 + 개발 디렉토리 골격

**Files:**
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\.gitignore`
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\skill_src\` (디렉토리)
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\skill_src\src\__init__.py` (빈 파일)
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\skill_src\templates\` (디렉토리)
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\skill_src\tests\unit\` (디렉토리)
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\skill_src\tests\integration\` (디렉토리)
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\skill_src\tests\fixtures\golden\extractor_outputs\` (디렉토리)
- Create: `C:\_PYTHON\0_CK_Project\8. Reviewer\agents_src\` (디렉토리)

- [ ] **Step 1: 작업 디렉토리에서 git init**

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer"
git init
```

Expected: "Initialized empty Git repository in ..."

- [ ] **Step 2: .gitignore 작성**

`.gitignore`:
```
__pycache__/
*.py[cod]
*$py.class
*.egg-info/
.pytest_cache/
.coverage
htmlcov/
review_ws/
review_output/
dist/
build/
.venv/
.vscode/
.idea/
*.swp
*.tmp
~$*.pptx          # MS Office 임시 파일
```

- [ ] **Step 3: 디렉토리 구조 생성**

```bash
mkdir -p "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/src"
mkdir -p "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/templates"
mkdir -p "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/tests/unit"
mkdir -p "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/tests/integration"
mkdir -p "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/tests/fixtures/golden/extractor_outputs"
mkdir -p "C:/_PYTHON/0_CK_Project/8. Reviewer/agents_src"
```

- [ ] **Step 4: 빈 src/__init__.py 생성**

```python
# skill_src/src/__init__.py
```

- [ ] **Step 5: 첫 commit (개별 파일 명시)**

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer"
git add .gitignore
git add docs/superpowers/specs/2026-04-27-report-reviewer-design.md
git add docs/superpowers/plans/2026-04-27-report-reviewer.md
git add reference/pptx추출.png
git add skill_src/src/__init__.py
git commit -m "chore: 개발 디렉토리 골격 + git 초기화 + 설계·계획 문서"
```

**Note:** 사용자 글로벌 규칙("`git add .` 금지, 수정 파일만 개별 add")을 첫 commit부터 일관 적용. 디렉토리 add(`docs/`, `reference/`)는 사용 안 함.

---

### Task 0.2: pyproject.toml + 의존성 설치

**Files:**
- Create: `skill_src/pyproject.toml`

- [ ] **Step 1: pyproject.toml 작성**

`skill_src/pyproject.toml`:
```toml
[project]
name = "report-reviewer"
version = "0.1.0"
description = "사내 PPTX 보고서 자가점검 도구 (Claude Code Skill)"
requires-python = ">=3.10"
dependencies = [
    "python-pptx>=0.6.23",
    "Pillow>=10.0",
    "Jinja2>=3.1",
    "pywin32>=306; sys_platform == 'win32'",
    "pdf2image>=1.17",
]

[project.optional-dependencies]
dev = [
    "pytest>=7.4",
    "pytest-cov>=4.1",
    "pytest-mock>=3.12",
]

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[tool.setuptools.packages.find]
where = ["."]
include = ["src*"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["."]
```

- [ ] **Step 2: 가상환경 생성 + 설치**

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
python -m venv .venv
"C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/.venv/Scripts/pip" install -e ".[dev]"
```

Expected: 모든 의존성 설치 성공. pywin32는 Windows에서만 설치됨.

- [ ] **Step 3: pytest 동작 확인 (테스트 0개)**

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
"C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/.venv/Scripts/pytest" -v
```

Expected: "no tests ran" (정상 종료)

- [ ] **Step 4: commit**

```bash
git add skill_src/pyproject.toml
git commit -m "chore: skill_src pyproject.toml + 의존성 정의"
```

**Note:** `.venv/`는 `.gitignore`에 이미 포함됨.

---

### Task 0.3: src/config.py — 환경변수·기본값 dataclass

**Files:**
- Create: `skill_src/src/config.py`
- Create: `skill_src/tests/unit/test_config.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`skill_src/tests/unit/test_config.py`:
```python
import os
from src.config import Config


def test_default_config():
    cfg = Config.from_env({})
    assert cfg.batch_size == 5
    assert cfg.image_max_dim == 1928
    assert cfg.pdf_dpi == 150
    assert cfg.thumbnail_max_dim == 480


def test_env_override_batch_size():
    cfg = Config.from_env({"REPORT_REVIEWER_BATCH_SIZE": "10"})
    assert cfg.batch_size == 10


def test_env_override_image_dim():
    cfg = Config.from_env({"REPORT_REVIEWER_IMAGE_MAX_DIM": "1280"})
    assert cfg.image_max_dim == 1280


def test_invalid_int_raises():
    import pytest
    with pytest.raises(ValueError):
        Config.from_env({"REPORT_REVIEWER_BATCH_SIZE": "abc"})
```

- [ ] **Step 2: 테스트 실패 확인**

```bash
cd skill_src
.venv/Scripts/pytest tests/unit/test_config.py -v
```

Expected: FAIL with "ModuleNotFoundError: No module named 'src.config'"

- [ ] **Step 3: 최소 구현**

`skill_src/src/config.py`:
```python
from __future__ import annotations
from dataclasses import dataclass
from typing import Mapping


@dataclass(frozen=True)
class Config:
    batch_size: int = 5
    image_max_dim: int = 1928
    pdf_dpi: int = 150
    thumbnail_max_dim: int = 480

    @classmethod
    def from_env(cls, env: Mapping[str, str]) -> "Config":
        def _int(key: str, default: int) -> int:
            raw = env.get(key)
            if raw is None:
                return default
            return int(raw)  # ValueError on invalid

        return cls(
            batch_size=_int("REPORT_REVIEWER_BATCH_SIZE", 5),
            image_max_dim=_int("REPORT_REVIEWER_IMAGE_MAX_DIM", 1928),
            pdf_dpi=_int("REPORT_REVIEWER_PDF_DPI", 150),
            thumbnail_max_dim=_int("REPORT_REVIEWER_THUMBNAIL_MAX_DIM", 480),
        )
```

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_config.py -v
```

Expected: PASS (4 passed)

- [ ] **Step 5: commit**

```bash
git add skill_src/src/config.py skill_src/tests/unit/test_config.py
git commit -m "feat(config): 환경변수 기반 Config dataclass + 단위 테스트"
```

---

### Task 0.4: tests/conftest.py + fixture PPTX 4종 생성 스크립트

**Files:**
- Create: `skill_src/tests/conftest.py`
- Create: `skill_src/tests/fixtures/_make_fixtures.py` (한 번 실행, 결과는 git 추적)
- Create: `skill_src/tests/fixtures/sample_text_only.pptx` (생성됨)
- Create: `skill_src/tests/fixtures/sample_with_table.pptx` (생성됨)
- Create: `skill_src/tests/fixtures/sample_with_image.pptx` (생성됨)
- Create: `skill_src/tests/fixtures/sample_with_embedded.pptx` (생성 어려움 → manual)

- [ ] **Step 1: conftest.py 작성**

`skill_src/tests/conftest.py`:
```python
from pathlib import Path
import pytest


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def sample_text_only(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_text_only.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p} — run tests/fixtures/_make_fixtures.py")
    return p


@pytest.fixture
def sample_with_table(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_table.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p}")
    return p


@pytest.fixture
def sample_with_image(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_image.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p}")
    return p


@pytest.fixture
def sample_with_embedded(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_embedded.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p} — manually create with embedded OLE object")
    return p


@pytest.fixture
def tmp_work_dir(tmp_path: Path) -> Path:
    return tmp_path / "review_ws"
```

- [ ] **Step 2: fixture 생성 스크립트 작성**

`skill_src/tests/fixtures/_make_fixtures.py`:
```python
"""
Fixture PPTX 생성 스크립트.
한 번 실행하여 fixtures/*.pptx를 생성. 결과는 git에 commit하여 재실행 불필요.

sample_with_embedded.pptx는 PowerPoint UI에서 수동 생성해야 함
(임베디드 OLE 객체는 python-pptx로 생성 불가).
"""
from __future__ import annotations
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


HERE = Path(__file__).parent


def make_text_only() -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9
    prs.slide_height = Inches(7.5)

    # Slide 1: 표지
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title
    slide.shapes.title.text = "프론트 서스펜션 강도 해석"
    slide.placeholders[1].text = "작성자: 홍길동\n2026-04-27"

    # Slide 2: 목차
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "목차"
    slide.placeholders[1].text = (
        "1. 해석 목적\n2. 모델 정보\n3. 경계조건\n4. 결과\n5. 결론"
    )

    # Slide 3~5: 본문
    for i, t in enumerate(["해석 목적", "모델 정보", "결과"]):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = t
        slide.placeholders[1].text = f"{t}에 대한 설명입니다.\n참고 사항을 적습니다."
        # 발표자 노트
        slide.notes_slide.notes_text_frame.text = f"슬라이드 {i+3} 발표 노트"

    prs.save(HERE / "sample_text_only.pptx")
    print(f"Created: {HERE / 'sample_text_only.pptx'}")


def make_with_table() -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank
    slide.shapes.title.text = "결과 표"

    # 표 추가
    rows, cols = 4, 3
    left = Inches(2)
    top = Inches(2)
    width = Inches(6)
    height = Inches(3)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    headers = ["항목", "값", "단위"]
    data = [
        ["최대 응력", "240", "MPa"],
        ["허용 응력", "350", "MPa"],
        ["안전계수", "1.46", "-"],
    ]
    for c, h in enumerate(headers):
        table.cell(0, c).text = h
    for r, row_data in enumerate(data, start=1):
        for c, v in enumerate(row_data):
            table.cell(r, c).text = v

    prs.save(HERE / "sample_with_table.pptx")
    print(f"Created: {HERE / 'sample_with_table.pptx'}")


def make_with_image() -> None:
    """단순 도형(차트 대신)을 추가한 슬라이드. 실제 이미지 파일은 PIL로 즉석 생성."""
    from pptx.util import Emu
    from PIL import Image
    import io

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "그래프"

    # 더미 이미지 생성
    img = Image.new("RGB", (400, 300), color=(200, 220, 255))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, Inches(2), Inches(2), Inches(5), Inches(3.75))

    prs.save(HERE / "sample_with_image.pptx")
    print(f"Created: {HERE / 'sample_with_image.pptx'}")


if __name__ == "__main__":
    make_text_only()
    make_with_table()
    make_with_image()
    print("\nNOTE: sample_with_embedded.pptx는 PowerPoint에서 수동 생성:")
    print("  1. 새 PPT 만들기")
    print("  2. 슬라이드에 '삽입 > 개체 > Microsoft Excel Worksheet' 추가")
    print("  3. fixtures/sample_with_embedded.pptx로 저장")
```

- [ ] **Step 3: fixture 생성 실행**

```bash
cd skill_src
.venv/Scripts/python tests/fixtures/_make_fixtures.py
```

Expected: 3개 PPTX 파일 생성, 마지막 NOTE 메시지 출력.

- [ ] **Step 4: sample_with_embedded.pptx 수동 생성**

PowerPoint에서:
1. 새 PPT
2. 슬라이드 1장 추가 → 삽입 > 개체 > Microsoft Excel 워크시트 (또는 PDF) 임베드
3. `skill_src/tests/fixtures/sample_with_embedded.pptx`로 저장

- [ ] **Step 5: 모든 fixture 확인**

```bash
ls skill_src/tests/fixtures/*.pptx
```

Expected: 4개 파일 모두 존재.

- [ ] **Step 6: commit**

```bash
git add skill_src/tests/conftest.py skill_src/tests/fixtures/_make_fixtures.py skill_src/tests/fixtures/*.pptx
git commit -m "test: pytest conftest + fixture PPTX 4종 (text/table/image/embedded)"
```

---

### Task 1.1: extractor — 단순 텍스트 슬라이드 추출

**Files:**
- Create: `skill_src/src/extractor.py`
- Create: `skill_src/tests/unit/test_extractor.py`

- [ ] **Step 1: 실패하는 테스트 작성 (가장 단순한 케이스)**

`skill_src/tests/unit/test_extractor.py`:
```python
from pathlib import Path
import pytest
from src.extractor import extract


def test_extract_returns_metadata_and_slides(sample_text_only: Path):
    result = extract(sample_text_only)
    assert "metadata" in result
    assert "slides" in result
    assert result["metadata"]["slide_count"] == 5  # 표지+목차+3장
    assert len(result["slides"]) == 5


def test_extract_first_slide_has_title(sample_text_only: Path):
    result = extract(sample_text_only)
    slide1 = result["slides"][0]
    assert slide1["index"] == 1
    titles = [s["text"] for s in slide1["shapes"] if s["type"] == "Title"]
    assert any("프론트 서스펜션 강도 해석" in t for t in titles)


def test_extract_includes_file_path(sample_text_only: Path):
    result = extract(sample_text_only)
    assert result["metadata"]["file_path"] == str(sample_text_only)
```

- [ ] **Step 2: 테스트 실패 확인**

```bash
.venv/Scripts/pytest tests/unit/test_extractor.py -v
```

Expected: FAIL with "ModuleNotFoundError: No module named 'src.extractor'"

- [ ] **Step 3: 최소 구현 — metadata + 텍스트 도형만**

`skill_src/src/extractor.py`:
```python
from __future__ import annotations
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def extract(pptx_path: Path) -> dict[str, Any]:
    pptx_path = Path(pptx_path)
    prs = Presentation(str(pptx_path))

    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        shapes = _extract_shapes(slide, slide_index=idx)
        slides.append({
            "index": idx,
            "title": _get_slide_title(slide),
            "layout": slide.slide_layout.name if slide.slide_layout else None,
            "shapes": shapes,
            "notes": "",
            "image_path": None,
            "thumbnail_path": None,
            "embedded_image_paths": [],
            "has_embedded": False,
        })

    cp = prs.core_properties
    metadata = {
        "file_path": str(pptx_path),
        "title": cp.title or "",
        "author": cp.author or "",
        "created": cp.created.isoformat() if cp.created else "",
        "modified": cp.modified.isoformat() if cp.modified else "",
        "slide_count": len(slides),
    }
    return {"metadata": metadata, "slides": slides}


def _get_slide_title(slide) -> str:
    if slide.shapes.title is not None:
        return slide.shapes.title.text or ""
    return ""


def _extract_shapes(slide, slide_index: int) -> list[dict[str, Any]]:
    out = []
    for sh_idx, shape in enumerate(slide.shapes, start=1):
        shape_type = _classify_shape_type(shape)
        text = ""
        if shape.has_text_frame:
            text = shape.text_frame.text or ""
        out.append({
            "shape_id": f"s{slide_index}_sh{sh_idx}",
            "type": shape_type,
            "position_emu": {
                "left": int(shape.left) if shape.left is not None else 0,
                "top": int(shape.top) if shape.top is not None else 0,
                "width": int(shape.width) if shape.width is not None else 0,
                "height": int(shape.height) if shape.height is not None else 0,
            },
            "position_pct": {},
            "z_order": sh_idx,
            "text": text,
            "table": None,
            "image_ref": None,
            "embedded_progid": None,
        })
    return out


def _classify_shape_type(shape) -> str:
    # Title placeholder 우선 검출
    try:
        if shape.is_placeholder and shape.placeholder_format.idx == 0:
            return "Title"
    except Exception:
        pass
    st = shape.shape_type
    mapping = {
        MSO_SHAPE_TYPE.PICTURE: "Picture",
        MSO_SHAPE_TYPE.TABLE: "Table",
        MSO_SHAPE_TYPE.GROUP: "Group",
        MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT: "EmbeddedOLE",
        MSO_SHAPE_TYPE.TEXT_BOX: "TextBox",
    }
    return mapping.get(st, "Other")
```

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_extractor.py -v
```

Expected: PASS (3 passed)

- [ ] **Step 5: commit**

```bash
git add skill_src/src/extractor.py skill_src/tests/unit/test_extractor.py
git commit -m "feat(extractor): 메타데이터 + 텍스트 도형 추출 (단순 케이스)"
```

---

### Task 1.2: extractor — 표 추출 추가

**Files:**
- Modify: `skill_src/src/extractor.py`
- Modify: `skill_src/tests/unit/test_extractor.py`

- [ ] **Step 1: 실패하는 테스트 추가**

`skill_src/tests/unit/test_extractor.py`에 추가:
```python
def test_extract_table_shape(sample_with_table: Path):
    result = extract(sample_with_table)
    slide = result["slides"][0]
    tables = [s for s in slide["shapes"] if s["type"] == "Table"]
    assert len(tables) == 1
    tbl = tables[0]["table"]
    assert tbl is not None
    assert tbl["rows"] == 4
    assert tbl["cols"] == 3
    # 헤더 행
    assert tbl["cells"][0] == ["항목", "값", "단위"]
    # 데이터 행
    assert tbl["cells"][1] == ["최대 응력", "240", "MPa"]
```

- [ ] **Step 2: 테스트 실패 확인**

```bash
.venv/Scripts/pytest tests/unit/test_extractor.py::test_extract_table_shape -v
```

Expected: FAIL — table 필드가 None

- [ ] **Step 3: `_extract_shapes()` 함수를 다음 코드로 교체**

`skill_src/src/extractor.py`의 `_extract_shapes` 함수 전체를 아래로 교체:

```python
def _extract_shapes(slide, slide_index: int) -> list[dict[str, Any]]:
    out = []
    for sh_idx, shape in enumerate(slide.shapes, start=1):
        shape_type = _classify_shape_type(shape)
        text = ""
        if shape.has_text_frame:
            text = shape.text_frame.text or ""

        table_data = None
        if shape_type == "Table" and shape.has_table:
            tbl = shape.table
            cells = [[cell.text for cell in row.cells] for row in tbl.rows]
            table_data = {
                "rows": len(tbl.rows),
                "cols": len(tbl.columns),
                "cells": cells,
            }

        out.append({
            "shape_id": f"s{slide_index}_sh{sh_idx}",
            "type": shape_type,
            "position_emu": {
                "left": int(shape.left) if shape.left is not None else 0,
                "top": int(shape.top) if shape.top is not None else 0,
                "width": int(shape.width) if shape.width is not None else 0,
                "height": int(shape.height) if shape.height is not None else 0,
            },
            "position_pct": {},
            "z_order": sh_idx,
            "text": text,
            "table": table_data,
            "image_ref": None,
            "embedded_progid": None,
        })
    return out
```

**변경점**: `table_data` 변수 추가 + `out.append`의 `"table"` 키에 반영. `position_emu`는 Task 1.1에서 작성한 것을 그대로 유지.

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_extractor.py -v
```

Expected: 모든 테스트 PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/extractor.py skill_src/tests/unit/test_extractor.py
git commit -m "feat(extractor): 표 셀 데이터 추출"
```

---

### Task 1.3: extractor — position_pct (정규화 좌표) 추가

**Files:**
- Modify: `skill_src/src/extractor.py`
- Modify: `skill_src/tests/unit/test_extractor.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_extract_position_pct(sample_text_only: Path):
    result = extract(sample_text_only)
    slide1 = result["slides"][0]
    for shape in slide1["shapes"]:
        pct = shape["position_pct"]
        assert "left" in pct and 0.0 <= pct["left"] <= 1.0
        assert "top" in pct and 0.0 <= pct["top"] <= 1.0
        assert "width" in pct and 0.0 <= pct["width"] <= 1.0
        assert "height" in pct and 0.0 <= pct["height"] <= 1.0
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL — position_pct가 빈 dict

- [ ] **Step 3: `extract()`와 `_extract_shapes()` 함수 시그니처·본문 교체**

먼저 `extract()` 함수를 다음으로 교체 (slide_w/slide_h를 _extract_shapes에 전달):

```python
def extract(pptx_path: Path) -> dict[str, Any]:
    pptx_path = Path(pptx_path)
    prs = Presentation(str(pptx_path))
    slide_w = prs.slide_width or 0
    slide_h = prs.slide_height or 0

    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        shapes = _extract_shapes(slide, slide_index=idx, slide_w=slide_w, slide_h=slide_h)
        slides.append({
            "index": idx,
            "title": _get_slide_title(slide),
            "layout": slide.slide_layout.name if slide.slide_layout else None,
            "shapes": shapes,
            "notes": "",
            "image_path": None,
            "thumbnail_path": None,
            "embedded_image_paths": [],
            "has_embedded": False,
        })

    cp = prs.core_properties
    metadata = {
        "file_path": str(pptx_path),
        "title": cp.title or "",
        "author": cp.author or "",
        "created": cp.created.isoformat() if cp.created else "",
        "modified": cp.modified.isoformat() if cp.modified else "",
        "slide_count": len(slides),
    }
    return {"metadata": metadata, "slides": slides}
```

이어서 `_extract_shapes()` 함수를 다음으로 교체 (slide_w/slide_h 인자 추가 + position_pct 계산):

```python
def _extract_shapes(slide, slide_index: int, slide_w: int, slide_h: int) -> list[dict[str, Any]]:
    out = []
    for sh_idx, shape in enumerate(slide.shapes, start=1):
        shape_type = _classify_shape_type(shape)
        text = ""
        if shape.has_text_frame:
            text = shape.text_frame.text or ""

        table_data = None
        if shape_type == "Table" and shape.has_table:
            tbl = shape.table
            cells = [[cell.text for cell in row.cells] for row in tbl.rows]
            table_data = {
                "rows": len(tbl.rows),
                "cols": len(tbl.columns),
                "cells": cells,
            }

        left = int(shape.left) if shape.left is not None else 0
        top = int(shape.top) if shape.top is not None else 0
        width = int(shape.width) if shape.width is not None else 0
        height = int(shape.height) if shape.height is not None else 0

        position_pct: dict[str, float] = {}
        if slide_w and slide_h:
            position_pct = {
                "left": left / slide_w,
                "top": top / slide_h,
                "width": width / slide_w,
                "height": height / slide_h,
            }

        out.append({
            "shape_id": f"s{slide_index}_sh{sh_idx}",
            "type": shape_type,
            "position_emu": {"left": left, "top": top, "width": width, "height": height},
            "position_pct": position_pct,
            "z_order": sh_idx,
            "text": text,
            "table": table_data,
            "image_ref": None,
            "embedded_progid": None,
        })
    return out
```

**변경점**: `extract()`는 `prs.slide_width/height` 추출 후 `_extract_shapes`에 전달. `_extract_shapes`는 시그니처에 `slide_w, slide_h` 추가, 본문에 `position_pct` 계산 로직 추가.

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/extractor.py skill_src/tests/unit/test_extractor.py
git commit -m "feat(extractor): position_pct (정규화 좌표) 추가"
```

---

### Task 1.4: extractor — 발표자 노트 추출

**Files:**
- Modify: `skill_src/src/extractor.py`
- Modify: `skill_src/tests/unit/test_extractor.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_extract_speaker_notes(sample_text_only: Path):
    result = extract(sample_text_only)
    # _make_fixtures.py에서 슬라이드 3~5에 노트 추가됨
    slide3 = result["slides"][2]
    assert "발표 노트" in slide3["notes"]
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL — notes가 빈 문자열

- [ ] **Step 3: `extract()` 함수의 슬라이드 루프에 notes 추출 추가**

`extract()` 함수의 `for idx, slide in enumerate(prs.slides, ...)` 루프를 다음으로 교체:

```python
    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        shapes = _extract_shapes(slide, slide_index=idx, slide_w=slide_w, slide_h=slide_h)
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text or ""
        slides.append({
            "index": idx,
            "title": _get_slide_title(slide),
            "layout": slide.slide_layout.name if slide.slide_layout else None,
            "shapes": shapes,
            "notes": notes,
            "image_path": None,
            "thumbnail_path": None,
            "embedded_image_paths": [],
            "has_embedded": False,
        })
```

**변경점**: `notes` 변수 추가 (`has_notes_slide` 체크 후 `notes_text_frame.text`) + `slides.append`의 `"notes": ""` → `"notes": notes`.

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/extractor.py skill_src/tests/unit/test_extractor.py
git commit -m "feat(extractor): 발표자 노트 추출"
```

---

### Task 1.5: extractor — 임베디드 OLE 감지 (has_embedded 플래그)

**Files:**
- Modify: `skill_src/src/extractor.py`
- Modify: `skill_src/tests/unit/test_extractor.py`

**전제조건 경고:** 이 task는 `sample_with_embedded.pptx` fixture가 필요합니다. Task 0.4 Step 4에서 수동 생성하지 않았다면 conftest의 `pytest.skip()`이 작동하여 두 번째 테스트가 skip됩니다. 이 경우 첫 번째 테스트(`test_extract_no_embedded_for_text_only`)만 통과 확인하고, fixture 추가 후 두 번째 테스트를 별도 검증해야 합니다. **임베디드 검증 없이 task를 완료 처리하지 마세요.**

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_extract_no_embedded_for_text_only(sample_text_only: Path):
    result = extract(sample_text_only)
    for slide in result["slides"]:
        assert slide["has_embedded"] is False


def test_extract_detects_embedded(sample_with_embedded: Path):
    result = extract(sample_with_embedded)
    has_any_embedded = any(s["has_embedded"] for s in result["slides"])
    assert has_any_embedded is True
    # 임베디드 도형이 있는 슬라이드는 EmbeddedOLE 타입 + progid 비어있지 않음
    for slide in result["slides"]:
        if slide["has_embedded"]:
            ole_shapes = [s for s in slide["shapes"] if s["type"] == "EmbeddedOLE"]
            assert len(ole_shapes) >= 1
            assert ole_shapes[0]["embedded_progid"]  # non-empty
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL

- [ ] **Step 3: extractor.py 변경 — `_get_embedded_progid` 추가, `_extract_shapes`·`extract` 수정**

(1) `extractor.py` 상단 import에 추가:
```python
from pptx.oxml.ns import qn
```

(2) 파일에 새 헬퍼 함수 추가 (예: `_classify_shape_type` 다음):
```python
def _get_embedded_progid(shape) -> str | None:
    """EmbeddedOLE 도형의 ProgID 추출 (예: 'PowerPoint.Show.12', 'Excel.Sheet.12').

    python-pptx 버전 차이를 고려해 두 단계로 시도:
    1) `shape.ole_format.prog_id` (신규 API)
    2) OOXML <p:oleObj progId="..."> 속성에서 직접 추출 (fallback)
    """
    # (1) 신규 API
    try:
        if hasattr(shape, "ole_format") and shape.ole_format is not None:
            pid = shape.ole_format.prog_id
            if pid:
                return pid
    except Exception:
        pass

    # (2) Fallback: OOXML 직접 파싱
    try:
        gframe = shape._element  # p:graphicFrame
        ole_objs = gframe.findall(".//" + qn("p:oleObj"))
        if ole_objs:
            pid = ole_objs[0].get("progId")
            if pid:
                return pid
    except Exception:
        pass

    return None
```

(3) `_extract_shapes` 함수의 도형 처리 루프에서 `embedded_progid`를 채우도록 수정. 함수의 `out.append(...)` 직전에 다음 추가:
```python
        embedded_progid = None
        if shape_type == "EmbeddedOLE":
            embedded_progid = _get_embedded_progid(shape)
```
그리고 `out.append({...})`의 마지막 필드 `"embedded_progid": None` → `"embedded_progid": embedded_progid`로 변경.

(4) `extract()` 함수의 슬라이드 루프에서 `has_embedded` 계산. `slides.append({...})` 직전에 다음 추가:
```python
        has_embedded = any(s["type"] == "EmbeddedOLE" for s in shapes)
```
그리고 `slides.append`의 `"has_embedded": False` → `"has_embedded": has_embedded`로 변경.

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS (sample_with_embedded.pptx 존재 시)

- [ ] **Step 5: commit**

```bash
git add skill_src/src/extractor.py skill_src/tests/unit/test_extractor.py
git commit -m "feat(extractor): 임베디드 OLE 감지 (has_embedded + progid)"
```

---

### Task 1.6: extractor — 골든 파일 회귀 테스트

**Files:**
- Create: `skill_src/tests/fixtures/golden/extractor_outputs/sample_text_only.json`
- Modify: `skill_src/tests/unit/test_extractor.py`

- [ ] **Step 1: 골든 파일 생성 스크립트 한 번 실행 (수동)**

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
"C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/.venv/Scripts/python" -c "
import json
from pathlib import Path
from src.extractor import extract
result = extract(Path('tests/fixtures/sample_text_only.pptx'))
# 변동 가능 필드는 골든에서 정규화
result['metadata']['file_path'] = '<FIXTURE>'
result['metadata']['created'] = '<NORMALIZED>'
result['metadata']['modified'] = '<NORMALIZED>'
out = Path('tests/fixtures/golden/extractor_outputs/sample_text_only.json')
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print('Saved:', out)
"
```

**골든 파일 수동 검토 체크리스트** (commit 전 다음 항목을 직접 확인):
- [ ] `metadata.slide_count == 5`
- [ ] 5개 슬라이드 모두 `index`가 1~5
- [ ] 슬라이드 1의 `shapes` 중 `type=="Title"`이 있고 텍스트에 "프론트 서스펜션 강도 해석" 포함
- [ ] 슬라이드 3~5의 `notes`에 "발표 노트" 포함
- [ ] 모든 도형의 `position_pct.left/top` 값이 0.0~1.0 범위
- [ ] `metadata.file_path == "<FIXTURE>"`, `created/modified == "<NORMALIZED>"`
- [ ] 모든 슬라이드의 `has_embedded == false`

- [ ] **Step 2: 회귀 테스트 추가**

```python
import json

def test_extract_text_only_matches_golden(sample_text_only: Path, fixtures_dir: Path):
    result = extract(sample_text_only)
    # 변동 가능 필드 정규화 (골든과 동일하게)
    result["metadata"]["file_path"] = "<FIXTURE>"
    result["metadata"]["created"] = "<NORMALIZED>"
    result["metadata"]["modified"] = "<NORMALIZED>"
    golden_path = fixtures_dir / "golden" / "extractor_outputs" / "sample_text_only.json"
    expected = json.loads(golden_path.read_text(encoding="utf-8"))
    assert result == expected
```

- [ ] **Step 3: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_extractor.py -v
```

Expected: 모두 PASS. 만약 실패하면 골든 파일 업데이트 필요 — 그 변경이 의도적인지 검토.

- [ ] **Step 4: commit**

```bash
git add skill_src/tests/fixtures/golden/ skill_src/tests/unit/test_extractor.py
git commit -m "test(extractor): 골든 파일 회귀 테스트 (sample_text_only)"
```

---

### Task 2.1: locator — EMU/pct → 한국어 위치 hint (단순)

**Files:**
- Create: `skill_src/src/locator.py`
- Create: `skill_src/tests/unit/test_locator.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`skill_src/tests/unit/test_locator.py`:
```python
from src.locator import to_hint


def test_top_left():
    pct = {"left": 0.05, "top": 0.05, "width": 0.3, "height": 0.1}
    hint = to_hint(pct)
    assert "좌측 상단" in hint or "좌상단" in hint


def test_top_right():
    pct = {"left": 0.70, "top": 0.10, "width": 0.25, "height": 0.10}
    hint = to_hint(pct)
    assert "우측 상단" in hint or "우상단" in hint


def test_center():
    pct = {"left": 0.40, "top": 0.40, "width": 0.20, "height": 0.20}
    hint = to_hint(pct)
    assert "중앙" in hint or "가운데" in hint


def test_bottom_full_width():
    pct = {"left": 0.00, "top": 0.85, "width": 1.00, "height": 0.10}
    hint = to_hint(pct)
    assert "하단" in hint


def test_includes_pct_numbers():
    pct = {"left": 0.70, "top": 0.15, "width": 0.25, "height": 0.10}
    hint = to_hint(pct)
    assert "70%" in hint
    assert "15%" in hint


def test_empty_returns_unknown():
    hint = to_hint({})
    assert "위치 미상" in hint or "unknown" in hint.lower()
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL — module not found

- [ ] **Step 3: 최소 구현**

`skill_src/src/locator.py`:
```python
from __future__ import annotations
from typing import Mapping


def to_hint(position_pct: Mapping[str, float]) -> str:
    """슬라이드 내 정규화 좌표(0~1)를 한국어 위치 hint 문자열로 변환.

    예: {"left":0.7,"top":0.15,"width":0.25,"height":0.1}
        → "우측 상단 (좌측 70%, 상단 15%)"
    """
    if not position_pct or "left" not in position_pct or "top" not in position_pct:
        return "위치 미상"

    left = position_pct["left"]
    top = position_pct["top"]
    width = position_pct.get("width", 0.0)

    # 가로 영역
    center_x = left + width / 2
    if center_x < 0.33:
        h_zone = "좌측"
    elif center_x < 0.67:
        h_zone = "중앙"
    else:
        h_zone = "우측"

    # 세로 영역
    center_y = top + position_pct.get("height", 0.0) / 2
    if center_y < 0.33:
        v_zone = "상단"
    elif center_y < 0.67:
        v_zone = "중단"
    else:
        v_zone = "하단"

    # 중앙·중단이면 "가운데"
    if h_zone == "중앙" and v_zone == "중단":
        zone = "가운데"
    else:
        zone = f"{h_zone} {v_zone}"

    return f"{zone} (좌측 {int(left*100)}%, 상단 {int(top*100)}%)"
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: 모두 PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/locator.py skill_src/tests/unit/test_locator.py
git commit -m "feat(locator): EMU/pct → 한국어 위치 hint 변환 (9분할)"
```

---

### Task 2.2: locator — 슬라이드 번호 + 도형 식별자 통합 hint

**Files:**
- Modify: `skill_src/src/locator.py`
- Modify: `skill_src/tests/unit/test_locator.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
from src.locator import format_location


def test_format_location_full():
    """슬라이드 번호 + 슬라이드 제목 + 도형 hint 통합 형식."""
    loc = format_location(
        slide_index=5,
        slide_title="결과",
        shape_type="TextBox",
        position_pct={"left": 0.7, "top": 0.15, "width": 0.25, "height": 0.1},
    )
    assert "슬라이드 5" in loc
    assert "결과" in loc
    assert "TextBox" in loc or "텍스트" in loc
    assert "우측 상단" in loc or "우상단" in loc


def test_format_location_no_title():
    loc = format_location(slide_index=3, slide_title="", shape_type="Table", position_pct={"left":0.1,"top":0.1,"width":0.3,"height":0.3})
    assert "슬라이드 3" in loc


def test_format_location_matches_spec_example():
    """Spec §5.5.3 예시: '슬라이드 5 우측 상단 텍스트 박스 (좌측 70%, 상단 15%)' 정합성."""
    loc = format_location(
        slide_index=5,
        slide_title="",
        shape_type="TextBox",
        position_pct={"left": 0.70, "top": 0.15, "width": 0.25, "height": 0.10},
    )
    assert "슬라이드 5" in loc
    assert "우측 상단" in loc
    assert "텍스트 박스" in loc
    assert "70%" in loc
    assert "15%" in loc
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL — format_location 없음

- [ ] **Step 3: format_location 추가**

```python
_TYPE_KO = {
    "TextBox": "텍스트 박스",
    "Title": "제목",
    "Table": "표",
    "Picture": "그림",
    "Chart": "차트",
    "Group": "그룹",
    "EmbeddedOLE": "임베디드 객체",
    "Other": "도형",
}


def format_location(
    slide_index: int,
    slide_title: str,
    shape_type: str,
    position_pct: Mapping[str, float],
) -> str:
    pos = to_hint(position_pct)
    ko_type = _TYPE_KO.get(shape_type, shape_type)
    title_part = f" ({slide_title})" if slide_title else ""
    return f"슬라이드 {slide_index}{title_part}의 {pos} {ko_type}"
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/locator.py skill_src/tests/unit/test_locator.py
git commit -m "feat(locator): format_location (슬라이드 번호+제목+도형 통합 hint)"
```

---

## Chunk 1 완료 시점

이 시점에:
- 개발 환경 셋업 완료 (git init, venv, pyproject.toml)
- 4종 fixture PPTX 준비 완료 (마지막 1종은 수동)
- `extractor.py`: PPTX → extracted dict 완성 (메타·텍스트·표·position_pct·노트·임베디드 감지)
- `locator.py`: 위치 → 한국어 hint 변환 + spec 예시 정합성 검증 완성
- 모든 단위 테스트 통과 (단, sample_with_embedded.pptx 미생성 시 1개 skip)
- 골든 파일 회귀 테스트 1건 (sample_text_only)

**향후 확장 메모:**
- `extractor.py`가 향후 임베디드 OLE 처리 로직(progid 분기)으로 비대해지면 `_oleparser.py`로 분리 검토.
- 골든 파일 추가(table/image/embedded)는 Chunk 2 이후 점진 확장.

다음으로 **Chunk 2 (Phase 3: slide_renderer / PowerPoint COM)**로 진행.

---

## Chunk 2: Phase 3 (slide_renderer / PowerPoint COM)

**[Chunk 2 cwd 면책]** Chunk 1과 동일. 모든 Bash 명령은 `cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"` 직후 실행을 가정.

**[Chunk 2 핵심 설계 결정]**

| 항목 | 결정 |
|---|---|
| COM 백엔드 | `win32com.client.Dispatch("PowerPoint.Application")` |
| Application 가시성 | `WithWindow=False` 우선, 보조로 `Visible = 0` 시도 (일부 버전이 거부) |
| COM 멀티스레드 | `pythoncom.CoInitialize()` / `CoUninitialize()` 쌍으로 처리 |
| 리소스 정리 | `try/finally`로 Presentation.Close() + Application.Quit() 보장 |
| 모드 1 출력 | JPG (1928×1080) — `slide.Export(path, "JPG", 1928, 1080)`, 반환 타입 `dict[int, Path]` |
| 모드 2 임베디드 분기 | **PDF만 별도 추출** (DoVerb(3) + SaveAs + pdf2image). PowerPoint·Excel·Word 임베디드는 슬라이드 캡처에 의존 (1차 MVP) |
| 자동 모드 선택 | `extractor`의 `has_embedded` 플래그로 분기 (모든 슬라이드 통합 판단) |
| 썸네일 | Pillow `Image.Resampling.LANCZOS`로 480px 이내 리사이즈 |
| 단위 테스트 | `unittest.mock.patch`로 COM 호출 mock + render 분기 검증은 mode 함수 자체 spy |
| Silent fail 금지 | 모든 임베디드 처리 실패는 `logging.warning`으로 기록 |

**[1차 MVP 임베디드 처리 결정]**: 임베디드 **PowerPoint/Excel/Word는 슬라이드 캡처에만 의존** (별도 추출 안 함). 이유:
- nested PowerPoint COM Dispatch는 singleton 충돌·외부 Presentation 강제 종료 위험
- spec §11 "향후 확장"에 임베디드 객체 별도 분석을 명시 (2차 이후 subprocess spawn 검토)
- PDF는 pdf2image가 별도 프로세스라 nested 위험 없음 → 1차 포함

---

### Task 3.1: COM mock 헬퍼 + slide_renderer 빈 골격

**Files:**
- Create: `skill_src/src/slide_renderer.py` (빈 함수 시그니처만)
- Create: `skill_src/tests/unit/test_slide_renderer.py`
- Modify: `skill_src/tests/conftest.py` (COM mock fixture 추가)

- [ ] **Step 1: slide_renderer.py 빈 골격 작성**

`skill_src/src/slide_renderer.py`:
```python
"""슬라이드 → 이미지 변환 (PowerPoint COM 기반).

두 모드:
  - convert_pptx_to_images_without_embedding: 단순 변환 (slide.Export)
  - convert_pptx_with_embedded_to_images: 임베디드 OLE 포함 변환 (DoVerb 분기)

오케스트레이터:
  - render: extracted dict의 has_embedded 플래그 기반 슬라이드별 자동 모드 선택
"""
from __future__ import annotations
from pathlib import Path
from typing import Any


def convert_pptx_to_images_without_embedding(
    pptx_path: Path,
    out_dir: Path,
    width: int = 1928,
    height: int = 1080,
) -> dict[int, Path]:
    """모드 1: 단순 PPTX → 슬라이드 JPG 변환. 슬라이드 인덱스 → 출력 경로 dict."""
    raise NotImplementedError


def convert_pptx_with_embedded_to_images(
    pptx_path: Path,
    out_dir: Path,
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
) -> dict[int, list[Path]]:
    """모드 2: 임베디드 OLE 포함 변환. 슬라이드 인덱스 → [원본 슬라이드 경로, *임베디드 경로들]."""
    raise NotImplementedError


def render(
    pptx_path: Path,
    out_dir: Path,
    extracted: dict[str, Any],
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
    thumbnail_max_dim: int = 480,
) -> dict[int, list[Path]]:
    """자동 모드 선택 오케스트레이터. extracted dict의 has_embedded 플래그로 슬라이드별 분기.

    슬라이드별로 모드 1 또는 모드 2 호출, 썸네일 생성, 결과 dict 반환.
    """
    raise NotImplementedError
```

- [ ] **Step 2: COM mock fixture 추가**

`skill_src/tests/conftest.py`의 **파일 최상단 import 블록에 다음을 추가** (기존 `from pathlib import Path`, `import pytest` 아래):
```python
import sys
from unittest.mock import MagicMock
from PIL import Image as _PILImage
```

그리고 **fixture 정의 영역(기존 fixture들 아래)에 다음 fixture 추가**:
```python
@pytest.fixture
def mock_powerpoint_com(monkeypatch):
    """`win32com.client.Dispatch`와 `pythoncom`을 mock하여 실제 PowerPoint 호출 없이 테스트.

    - Dispatch("PowerPoint.Application") → MagicMock 반환
    - Presentation.Slides 순회는 mock slide 객체 리스트
    - slide.Export(path, "JPG", w, h) 호출 시 실제로 PIL 더미 JPG를 path에 생성
    """
    fake_win32 = MagicMock()
    fake_pythoncom = MagicMock()

    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32.client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    def make_app(slide_count: int = 3, embedded_progids_per_slide: dict[int, list[str]] | None = None):
        app = MagicMock(name="PowerPointApplication")
        app.Visible = 0
        pres = MagicMock(name="Presentation")
        slides = []
        for i in range(1, slide_count + 1):
            sl = MagicMock(name=f"Slide{i}")
            sl.SlideIndex = i

            # 임베디드 도형 mock (모드 2 테스트에서 사용)
            embedded = (embedded_progids_per_slide or {}).get(i, [])
            shapes = []
            for j, progid in enumerate(embedded, start=1):
                shape = MagicMock(name=f"Slide{i}Shape{j}")
                shape.Type = 7  # msoEmbeddedOLEObject
                shape.OLEFormat.ProgID = progid
                shapes.append(shape)
            sl.Shapes = shapes

            def _make_export(idx, w_default=1928, h_default=1080):
                def _export(path, fmt, w, h):
                    img = _PILImage.new("RGB", (w, h), color=(220, 230, 240))
                    img.save(path)
                return _export
            sl.Export.side_effect = _make_export(i)
            slides.append(sl)
        pres.Slides = slides
        app.Presentations.Open.return_value = pres
        fake_win32.client.Dispatch.return_value = app
        return app, pres, slides

    return make_app
```

**변경점**: `tmp_path` 인자 제거(미사용), `embedded_progids_per_slide` 파라미터를 처음부터 포함(Task 3.3에서 재수정 불필요).

- [ ] **Step 3: 골격 import만 검증하는 smoke 테스트**

`skill_src/tests/unit/test_slide_renderer.py`:
```python
import pytest
from pathlib import Path
from src.slide_renderer import (
    convert_pptx_to_images_without_embedding,
    convert_pptx_with_embedded_to_images,
    render,
)


def test_module_imports():
    """세 공개 함수가 import되고 NotImplementedError를 던지는지 확인."""
    with pytest.raises(NotImplementedError):
        convert_pptx_to_images_without_embedding(Path("x.pptx"), Path("out"))
    with pytest.raises(NotImplementedError):
        convert_pptx_with_embedded_to_images(Path("x.pptx"), Path("out"))
    with pytest.raises(NotImplementedError):
        render(Path("x.pptx"), Path("out"), {"slides": []})
```

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
```

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/slide_renderer.py skill_src/tests/conftest.py skill_src/tests/unit/test_slide_renderer.py
git commit -m "feat(slide_renderer): 빈 골격 + COM mock fixture"
```

---

### Task 3.2: slide_renderer — 모드 1 (without embedding) 구현

**Files:**
- Modify: `skill_src/src/slide_renderer.py`
- Modify: `skill_src/tests/unit/test_slide_renderer.py`

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/unit/test_slide_renderer.py`에 추가:
```python
def test_mode1_creates_one_image_per_slide(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=3)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")  # 존재만 하면 됨 (mock이라 실제 파싱 안 함)
    out_dir = tmp_path / "out"

    result = convert_pptx_to_images_without_embedding(pptx_path, out_dir)

    # dict[int, Path] 반환
    assert set(result.keys()) == {1, 2, 3}
    for idx, p in result.items():
        assert p.exists()
        assert p.suffix.lower() == ".jpg"


def test_mode1_calls_export_with_correct_size(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=2)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    convert_pptx_to_images_without_embedding(pptx_path, out_dir, width=1280, height=720)

    for sl in slides:
        # Export(path, "JPG", 1280, 720) 호출 검증
        args = sl.Export.call_args
        assert args[0][1] == "JPG"
        assert args[0][2] == 1280
        assert args[0][3] == 720


def test_mode1_quits_powerpoint_even_on_error(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=1)
    slides[0].Export.side_effect = RuntimeError("export failed")
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    with pytest.raises(RuntimeError):
        convert_pptx_to_images_without_embedding(pptx_path, out_dir)

    pres.Close.assert_called()
    app.Quit.assert_called()
```

- [ ] **Step 2: 테스트 실패 확인**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
```

Expected: FAIL (NotImplementedError)

- [ ] **Step 3: `convert_pptx_to_images_without_embedding` 구현**

`skill_src/src/slide_renderer.py`의 동일 함수를 다음으로 교체 (반환 타입 `dict[int, Path]`):
```python
import logging

logger = logging.getLogger(__name__)


def convert_pptx_to_images_without_embedding(
    pptx_path: Path,
    out_dir: Path,
    width: int = 1928,
    height: int = 1080,
) -> dict[int, Path]:
    """모드 1: 단순 PPTX → 슬라이드 JPG 변환. 슬라이드 인덱스 → 출력 경로 dict."""
    import win32com.client
    import pythoncom

    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.Visible = 0  # 일부 버전은 0 거부 — 무시
        except Exception:
            logger.debug("PowerPoint.Visible=0 거부됨, WithWindow=False에 의존")

        pres = app.Presentations.Open(str(pptx_path), WithWindow=False)

        out_paths: dict[int, Path] = {}
        for slide in pres.Slides:
            idx = slide.SlideIndex
            target = out_dir / f"slide_{idx:03d}.jpg"
            slide.Export(str(target), "JPG", width, height)
            out_paths[idx] = target
        return out_paths
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            logger.warning("Presentation.Close() 실패", exc_info=True)
        try:
            if app is not None:
                app.Quit()
        except Exception:
            logger.warning("Application.Quit() 실패", exc_info=True)
        pythoncom.CoUninitialize()
```

**Note**: `WithWindow=False`로 창 자체를 안 띄움. `Visible=0`은 일부 PowerPoint 버전이 거부하므로 보조 안전장치. `logging.warning`으로 silent fail 방지.

- [ ] **Step 4: 테스트 실행**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
```

Expected: 신규 3개 테스트(`test_mode1_*`) PASS, **기존 `test_module_imports`는 FAIL** (이제 `convert_pptx_to_images_without_embedding`이 NotImplementedError를 던지지 않음). Step 5에서 수정.

- [ ] **Step 5: `test_module_imports` 수정**

```python
def test_module_imports():
    """모드 2와 render는 아직 NotImplementedError."""
    with pytest.raises(NotImplementedError):
        convert_pptx_with_embedded_to_images(Path("x.pptx"), Path("out"))
    with pytest.raises(NotImplementedError):
        render(Path("x.pptx"), Path("out"), {"slides": []})
```

- [ ] **Step 6: 다시 테스트 + commit**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
git add skill_src/src/slide_renderer.py skill_src/tests/unit/test_slide_renderer.py
git commit -m "feat(slide_renderer): 모드 1 (단순 변환) 구현 + COM 정리 보장"
```

---

### Task 3.3: slide_renderer — 모드 2 (with embedded) 구현

**Files:**
- Modify: `skill_src/src/slide_renderer.py`
- Modify: `skill_src/tests/unit/test_slide_renderer.py`

**[Task 3.3 핵심 분기 표 — 1차 MVP 결정 반영]**

| ProgID 패턴 | 처리 |
|---|---|
| `AcroExch.*`, `*PDF*` (PDF 임베디드) | DoVerb(3) → SaveAs PDF → `pdf2image.convert_from_path(dpi=150)` → PNG (별도 이미지) |
| `PowerPoint.Show.*` | **별도 처리 없음** — 슬라이드 캡처(원본 슬라이드 export)에 의존. ProgID는 logger로 INFO 기록 |
| `Excel.Sheet.*`, `Word.Document.*` | **별도 처리 없음** — 슬라이드 캡처에 의존. ProgID는 logger로 INFO 기록 |
| 그 외 알 수 없음 | logger.warning |

**중요**: nested `Dispatch("PowerPoint.Application")` 회피를 위해 PowerPoint 임베디드는 1차 MVP에서 별도 추출 안 함. spec §11에 따라 2차 이후 subprocess spawn으로 안전 분리 검토.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/unit/test_slide_renderer.py`의 **상단 import 블록에 다음 추가**:
```python
from unittest.mock import MagicMock
from PIL import Image as _PILImage
```

그리고 **테스트 함수 영역에 다음 추가** (Task 3.1의 conftest에 임베디드 mock이 이미 포함되어 있으므로 conftest 추가 수정 불필요):
```python
def test_mode2_no_embedded_acts_like_mode1(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=2)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    # dict[int, list[Path]] 반환, 임베디드 없으면 각 리스트 길이 1
    assert set(result.keys()) == {1, 2}
    for paths in result.values():
        assert len(paths) == 1
        assert paths[0].exists()


def test_mode2_pptshow_uses_slide_capture_only(mock_powerpoint_com, tmp_path, caplog):
    """PowerPoint 임베디드는 별도 추출 안 함 (1차 MVP), 슬라이드 캡처만."""
    import logging
    app, pres, slides = mock_powerpoint_com(
        slide_count=1,
        embedded_progids_per_slide={1: ["PowerPoint.Show.12"]},
    )
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    with caplog.at_level(logging.INFO, logger="src.slide_renderer"):
        result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    # 슬라이드 캡처만 (임베디드 별도 이미지 없음)
    assert result[1] == [out_dir / "slide_001.jpg"]
    # ProgID가 로그에 INFO로 기록되었는지
    assert any("PowerPoint.Show" in rec.message for rec in caplog.records)
    # DoVerb(1) 호출되지 않아야 함 (1차 MVP는 처리 안 함)
    embedded_shape = slides[0].Shapes[0]
    embedded_shape.OLEFormat.DoVerb.assert_not_called()


def test_mode2_pdf_progid_extracts_separately(mock_powerpoint_com, tmp_path, monkeypatch):
    """PDF 임베디드는 DoVerb(3) + SaveAs + pdf2image 호출, 별도 PNG 생성."""
    app, pres, slides = mock_powerpoint_com(
        slide_count=1,
        embedded_progids_per_slide={1: ["AcroExch.Document.DC"]},
    )
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    # SaveAs가 실제로 PDF를 생성하도록 side_effect
    def _fake_saveas(path):
        Path(path).write_bytes(b"%PDF-1.4 dummy")  # 더미 PDF
    slides[0].Shapes[0].OLEFormat.Object.SaveAs.side_effect = _fake_saveas

    fake_pages = [_PILImage.new("RGB", (200, 200), color=(255, 0, 0))]
    fake_convert = MagicMock(return_value=fake_pages)
    monkeypatch.setattr("pdf2image.convert_from_path", fake_convert)

    result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    embedded_shape = slides[0].Shapes[0]
    embedded_shape.OLEFormat.DoVerb.assert_any_call(3)
    fake_convert.assert_called_once()
    # 슬라이드 캡처 + 임베디드 PDF 추출 = 2개 경로
    assert len(result[1]) == 2
    # 두 번째 경로는 임베디드 PNG
    assert result[1][1].suffix.lower() == ".png"


def test_mode2_pdf_saveas_failure_logs_warning(mock_powerpoint_com, tmp_path, caplog):
    """PDF 임베디드 SaveAs 실패 시 silent하지 않고 logging.warning."""
    import logging
    app, pres, slides = mock_powerpoint_com(
        slide_count=1,
        embedded_progids_per_slide={1: ["AcroExch.Document.DC"]},
    )
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    slides[0].Shapes[0].OLEFormat.Object.SaveAs.side_effect = RuntimeError("Acrobat blocked")

    with caplog.at_level(logging.WARNING, logger="src.slide_renderer"):
        result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    assert any("PDF 임베디드" in rec.message or "SaveAs" in rec.message for rec in caplog.records)
    # 슬라이드 캡처는 성공
    assert result[1][0].exists()
```

- [ ] **Step 2: 테스트 실패 확인**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
```

Expected: 신규 4개 테스트(`test_mode2_*`) FAIL (NotImplementedError). 기존 모드 1 테스트는 PASS 유지.

- [ ] **Step 3: `convert_pptx_with_embedded_to_images` + `_process_embedded` 구현**

`skill_src/src/slide_renderer.py`의 `convert_pptx_with_embedded_to_images`를 다음으로 교체:
```python
def convert_pptx_with_embedded_to_images(
    pptx_path: Path,
    out_dir: Path,
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
) -> dict[int, list[Path]]:
    """모드 2: 임베디드 OLE 포함 변환. 슬라이드 인덱스 → [원본 슬라이드 경로, *임베디드 경로들].

    1차 MVP: PDF 임베디드만 별도 추출. PowerPoint/Excel/Word 임베디드는 슬라이드 캡처에 의존.
    """
    import win32com.client
    import pythoncom

    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    embed_dir = out_dir / "embedded"
    embed_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.Visible = 0
        except Exception:
            logger.debug("PowerPoint.Visible=0 거부됨, WithWindow=False에 의존")
        pres = app.Presentations.Open(str(pptx_path), WithWindow=False)

        result: dict[int, list[Path]] = {}
        for slide in pres.Slides:
            idx = slide.SlideIndex
            paths: list[Path] = []

            # 1) 원본 슬라이드 캡처 (모든 슬라이드 공통)
            main_target = out_dir / f"slide_{idx:03d}.jpg"
            slide.Export(str(main_target), "JPG", width, height)
            paths.append(main_target)

            # 2) 슬라이드 내 임베디드 OLE 도형 순회
            for j, shape in enumerate(slide.Shapes, start=1):
                if getattr(shape, "Type", None) != 7:  # msoEmbeddedOLEObject
                    continue
                try:
                    progid = shape.OLEFormat.ProgID or ""
                except Exception:
                    progid = ""

                emb_paths = _process_embedded(
                    shape=shape,
                    progid=progid,
                    slide_idx=idx,
                    shape_idx=j,
                    embed_dir=embed_dir,
                    pdf_dpi=pdf_dpi,
                )
                paths.extend(emb_paths)

            result[idx] = paths
        return result
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            logger.warning("Presentation.Close() 실패", exc_info=True)
        try:
            if app is not None:
                app.Quit()
        except Exception:
            logger.warning("Application.Quit() 실패", exc_info=True)
        pythoncom.CoUninitialize()


def _process_embedded(
    shape,
    progid: str,
    slide_idx: int,
    shape_idx: int,
    embed_dir: Path,
    pdf_dpi: int,
) -> list[Path]:
    """임베디드 OLE 도형 1개 처리. PDF만 별도 추출, 그 외는 ProgID 기록만.

    nested PowerPoint COM Dispatch를 회피하기 위해 PowerPoint 임베디드는 처리하지 않음.
    """
    out: list[Path] = []
    pid_lower = progid.lower()

    # PDF 임베디드: 별도 추출
    if "acroexch" in pid_lower or "pdf" in pid_lower:
        try:
            shape.OLEFormat.DoVerb(3)
        except Exception:
            logger.warning(
                "슬라이드 %d 임베디드 PDF DoVerb(3) 실패 (Acrobat 자동화 차단 가능). slide_%03d 캡처에 의존.",
                slide_idx, slide_idx, exc_info=True,
            )
            return out

        tmp_pdf = embed_dir / f"slide_{slide_idx:03d}_emb{shape_idx:02d}.pdf"
        try:
            shape.OLEFormat.Object.SaveAs(str(tmp_pdf))
        except Exception:
            logger.warning(
                "슬라이드 %d 임베디드 PDF SaveAs 실패. slide_%03d 캡처에 의존.",
                slide_idx, slide_idx, exc_info=True,
            )
            return out

        if not tmp_pdf.exists():
            logger.warning("슬라이드 %d 임베디드 PDF 저장 결과 파일 없음.", slide_idx)
            return out

        try:
            from pdf2image import convert_from_path
            pages = convert_from_path(str(tmp_pdf), dpi=pdf_dpi)
            for p_idx, page in enumerate(pages, start=1):
                out_png = embed_dir / f"slide_{slide_idx:03d}_emb{shape_idx:02d}_p{p_idx:02d}.png"
                page.save(out_png, "PNG")
                out.append(out_png)
        except Exception:
            logger.warning(
                "슬라이드 %d 임베디드 PDF → PNG 변환 실패 (Poppler 미설치 가능).",
                slide_idx, exc_info=True,
            )
        return out

    # PowerPoint 임베디드: 1차 MVP 미지원, 슬라이드 캡처에 의존
    if "powerpoint.show" in pid_lower:
        logger.info(
            "슬라이드 %d에 PowerPoint 임베디드(%s) 발견. 1차 MVP는 슬라이드 캡처에 의존.",
            slide_idx, progid,
        )
        return out

    # Excel/Word: 슬라이드 캡처에 의존
    if "excel" in pid_lower or "word" in pid_lower:
        logger.info(
            "슬라이드 %d에 Office 임베디드(%s) 발견. 슬라이드 캡처에 의존.",
            slide_idx, progid,
        )
        return out

    # 그 외 알 수 없는 ProgID
    logger.warning(
        "슬라이드 %d에 알 수 없는 임베디드 ProgID(%s). 슬라이드 캡처에 의존.",
        slide_idx, progid,
    )
    return out
```

**핵심 변경**: PowerPoint 임베디드 분기 제거 (nested Dispatch 회피), 모든 silent fail에 `logger.warning/info` 추가, `_process_embedded` 시그니처에서 `width/height` 인자 제거 (PDF만 처리하므로 불필요).

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
```

Expected: 신규 4개 테스트(`test_mode2_no_embedded_acts_like_mode1`, `test_mode2_pptshow_uses_slide_capture_only`, `test_mode2_pdf_progid_extracts_separately`, `test_mode2_pdf_saveas_failure_logs_warning`) 모두 PASS. mock의 SaveAs side_effect로 PDF 경로가 결정론적으로 1개 PNG 생성됨.

- [ ] **Step 5: commit**

```bash
git add skill_src/src/slide_renderer.py skill_src/tests/unit/test_slide_renderer.py
git commit -m "feat(slide_renderer): 모드 2 (PDF 임베디드만 별도 추출) + logging"
```

---

### Task 3.4: slide_renderer — `render()` 자동 모드 선택 + 썸네일 통합

**Files:**
- Modify: `skill_src/src/slide_renderer.py`
- Modify: `skill_src/tests/unit/test_slide_renderer.py`

- [ ] **Step 1: 실패하는 테스트 추가 — mode 함수 spy로 분기 검증**

```python
def test_render_calls_mode1_when_no_embedded(monkeypatch, tmp_path):
    """render가 has_embedded=False 시 mode 1을 호출하고 mode 2는 호출 안 함."""
    from PIL import Image as _PI
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    fake_jpg = out_dir / "slide_001.jpg"
    _PI.new("RGB", (100, 100)).save(fake_jpg)

    mode1_calls = []
    def fake_mode1(pptx, out, **kwargs):
        mode1_calls.append((pptx, out))
        return {1: fake_jpg}

    mode2_calls = []
    def fake_mode2(pptx, out, **kwargs):
        mode2_calls.append((pptx, out))
        return {}

    monkeypatch.setattr("src.slide_renderer.convert_pptx_to_images_without_embedding", fake_mode1)
    monkeypatch.setattr("src.slide_renderer.convert_pptx_with_embedded_to_images", fake_mode2)

    extracted = {"slides": [{"index": 1, "has_embedded": False}]}
    render(tmp_path / "in.pptx", out_dir, extracted)

    assert len(mode1_calls) == 1
    assert len(mode2_calls) == 0


def test_render_calls_mode2_when_any_embedded(monkeypatch, tmp_path):
    """render가 has_embedded=True가 하나라도 있으면 mode 2를 호출하고 mode 1은 호출 안 함."""
    from PIL import Image as _PI
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    fake_jpg1 = out_dir / "slide_001.jpg"
    fake_jpg2 = out_dir / "slide_002.jpg"
    _PI.new("RGB", (100, 100)).save(fake_jpg1)
    _PI.new("RGB", (100, 100)).save(fake_jpg2)

    mode1_calls = []
    def fake_mode1(pptx, out, **kwargs):
        mode1_calls.append((pptx, out))
        return {}

    mode2_calls = []
    def fake_mode2(pptx, out, **kwargs):
        mode2_calls.append((pptx, out))
        return {1: [fake_jpg1], 2: [fake_jpg2]}

    monkeypatch.setattr("src.slide_renderer.convert_pptx_to_images_without_embedding", fake_mode1)
    monkeypatch.setattr("src.slide_renderer.convert_pptx_with_embedded_to_images", fake_mode2)

    extracted = {"slides": [
        {"index": 1, "has_embedded": False},
        {"index": 2, "has_embedded": True},
    ]}
    render(tmp_path / "in.pptx", out_dir, extracted)

    assert len(mode2_calls) == 1
    assert len(mode1_calls) == 0


def test_render_creates_thumbnails(mock_powerpoint_com, tmp_path):
    """실제 mode 1 + 썸네일 생성 통합 검증."""
    app, pres, slides = mock_powerpoint_com(slide_count=1)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"
    extracted = {"slides": [{"index": 1, "has_embedded": False}]}

    render(pptx_path, out_dir, extracted, thumbnail_max_dim=300)

    thumb = out_dir / "thumbnails" / "slide_001.jpg"
    assert thumb.exists()
    from PIL import Image
    with Image.open(thumb) as img:
        assert max(img.size) <= 300
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL (NotImplementedError)

- [ ] **Step 3: `render()` 구현**

`skill_src/src/slide_renderer.py`의 `render` 함수를 다음으로 교체:
```python
def render(
    pptx_path: Path,
    out_dir: Path,
    extracted: dict[str, Any],
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
    thumbnail_max_dim: int = 480,
) -> dict[int, list[Path]]:
    """자동 모드 선택 오케스트레이터.

    extracted dict의 슬라이드별 has_embedded 플래그가 하나라도 True면 모드 2,
    아니면 모드 1을 사용한다. 썸네일은 항상 생성.
    """
    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    has_any_embedded = any(s.get("has_embedded") for s in extracted.get("slides", []))

    if has_any_embedded:
        result = convert_pptx_with_embedded_to_images(
            pptx_path, out_dir, width=width, height=height, pdf_dpi=pdf_dpi
        )
    else:
        paths_dict = convert_pptx_to_images_without_embedding(
            pptx_path, out_dir, width=width, height=height
        )
        # 모드 1은 dict[int, Path] 반환 → render는 dict[int, list[Path]]로 통일
        result = {idx: [p] for idx, p in paths_dict.items()}

    # 썸네일 생성
    thumb_dir = out_dir / "thumbnails"
    thumb_dir.mkdir(parents=True, exist_ok=True)
    _generate_thumbnails(result, thumb_dir, max_dim=thumbnail_max_dim)

    return result


def _generate_thumbnails(
    rendered: dict[int, list[Path]],
    thumb_dir: Path,
    max_dim: int,
) -> None:
    """원본 슬라이드 이미지(각 슬라이드의 첫 번째 경로)를 썸네일로 축소."""
    from PIL import Image

    for slide_idx, paths in rendered.items():
        if not paths:
            continue
        src = paths[0]
        target = thumb_dir / f"slide_{slide_idx:03d}.jpg"
        with Image.open(src) as img:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            # JPG는 RGB 강제
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.save(target, "JPG", quality=85)
```

**변경점**:
- 모드 1 결과를 `paths_dict.items()`로 안전하게 dict 변환 (파일명 파싱 의존 제거)
- `Image.LANCZOS` → `Image.Resampling.LANCZOS` (Pillow 10+ 권장)

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
```

Expected: 신규 3개 테스트 PASS. `test_module_imports`는 이제 `render`도 NotImplementedError가 아니므로 수정.

- [ ] **Step 5: `test_module_imports` 최종 수정**

```python
def test_module_imports():
    """모든 공개 함수가 import 가능 (구현 완료)."""
    assert callable(convert_pptx_to_images_without_embedding)
    assert callable(convert_pptx_with_embedded_to_images)
    assert callable(render)
```

- [ ] **Step 6: 다시 테스트 + commit**

```bash
.venv/Scripts/pytest tests/unit/test_slide_renderer.py -v
git add skill_src/src/slide_renderer.py skill_src/tests/unit/test_slide_renderer.py
git commit -m "feat(slide_renderer): render() 자동 모드 선택 + Pillow 썸네일"
```

---

### Task 3.5: 실제 PowerPoint COM 통합 테스트 (수동)

**Files:**
- Create: `skill_src/tests/integration/test_slide_renderer_real.py`

**[전제조건]** Windows + MS PowerPoint 설치 필수. CI에서 자동 실행 안 함, 사용자 로컬에서 수동.

- [ ] **Step 1: 통합 테스트 작성**

`skill_src/tests/integration/test_slide_renderer_real.py`:
```python
"""실제 PowerPoint COM을 호출하는 통합 테스트.

수동 실행:
  .venv/Scripts/pytest tests/integration/test_slide_renderer_real.py -v

CI 환경에서 자동 실행하지 않으려면 pyproject.toml에 마커 등록 + addopts로 기본 제외 (Step 2 참조).
"""
import sys
import pytest
from pathlib import Path


pytestmark = [
    pytest.mark.real_com,  # CI 기본 제외 마커
    pytest.mark.skipif(sys.platform != "win32", reason="Windows + MS PowerPoint 필요"),
]


def test_real_mode1_text_only(sample_text_only: Path, tmp_path: Path):
    from src.slide_renderer import convert_pptx_to_images_without_embedding

    out_dir = tmp_path / "out"
    result = convert_pptx_to_images_without_embedding(sample_text_only, out_dir)

    assert set(result.keys()) == {1, 2, 3, 4, 5}  # fixture는 5장
    for idx, p in result.items():
        assert p.exists()
        assert p.stat().st_size > 0


def test_real_render_with_thumbnails(sample_text_only: Path, tmp_path: Path):
    from src.extractor import extract
    from src.slide_renderer import render

    extracted = extract(sample_text_only)
    out_dir = tmp_path / "out"
    result = render(sample_text_only, out_dir, extracted)

    assert len(result) == 5
    thumb_dir = out_dir / "thumbnails"
    assert len(list(thumb_dir.glob("slide_*.jpg"))) == 5


def test_real_render_with_embedded(sample_with_embedded: Path, tmp_path: Path):
    """sample_with_embedded.pptx 부재 시 conftest의 fixture가 자동 skip."""
    from src.extractor import extract
    from src.slide_renderer import render

    extracted = extract(sample_with_embedded)
    out_dir = tmp_path / "out"
    result = render(sample_with_embedded, out_dir, extracted)

    # PDF 임베디드가 있고 Acrobat 자동화가 허용된 경우에만 multi-path
    # PowerPoint/Excel/Word 임베디드는 슬라이드 캡처에 의존(1차 MVP)이라 단일 경로
    # 따라서 이 테스트는 "결과 dict가 정상 생성"까지만 검증
    assert len(result) > 0
    for paths in result.values():
        assert len(paths) >= 1  # 최소 슬라이드 캡처 1개
```

- [ ] **Step 2: pyproject.toml에 마커 등록 + CI 기본 제외**

`skill_src/pyproject.toml`의 `[tool.pytest.ini_options]` 섹션을 다음으로 교체:
```toml
[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["."]
markers = [
    "real_com: requires real MS PowerPoint COM (Windows only, manual execution)",
]
addopts = "-m 'not real_com'"
```

**효과**: 일반 `pytest` 호출은 `real_com` 마커 테스트를 자동 제외. 통합 테스트 실행 시 `pytest -m real_com` 또는 명시적 파일 지정.

- [ ] **Step 3: 수동 통합 실행** (real_com 마커 명시)

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
"C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/.venv/Scripts/pytest" tests/integration/test_slide_renderer_real.py -v -m real_com
```

**Note**: `-m real_com`로 addopts의 기본 제외(`-m 'not real_com'`)를 덮어씀.

**예상 동작:**
- PowerPoint가 백그라운드에서 잠깐 떴다가 종료
- `tmp_path/out/slide_001.jpg` ~ `slide_005.jpg` 생성됨
- `tmp_path/out/thumbnails/slide_001.jpg` ~ 생성됨
- 임베디드 fixture 있을 시 result에 multi-path 슬라이드 1개 이상

**검증 체크리스트:**
- [ ] PowerPoint 프로세스가 테스트 종료 후 좀비로 남지 않음 (작업 관리자 확인)
- [ ] 생성된 JPG 파일이 1928×1080 또는 그에 비례한 크기
- [ ] 썸네일이 480px 이내
- [ ] 임베디드 테스트가 (fixture 있을 시) 통과 — 없으면 skip

- [ ] **Step 4: commit**

```bash
git add skill_src/tests/integration/test_slide_renderer_real.py skill_src/pyproject.toml
git commit -m "test(slide_renderer): 실제 PowerPoint COM 통합 테스트 (수동 실행)"
```

---

## Chunk 2 완료 시점

이 시점에:
- `slide_renderer.py`: 두 모드(without/with embedding) + `render()` 자동 모드 선택 + 썸네일 완성
- COM mock 단위 테스트 (자동 실행)
- 실제 PowerPoint 통합 테스트 (수동 실행)
- COM 리소스 정리 보장 (try/finally + Visible=0)
- pdf2image 통합 (PDF 임베디드 처리)

**향후 확장 메모:**
- Excel/Word/PowerPoint 임베디드의 별도 추출은 2차 이후 (현재는 슬라이드 캡처로 의존, subprocess spawn 검토)
- COM 호출이 느릴 경우 슬라이드별 병렬화는 GIL·COM 한계로 어려움 → 그대로 유지
- 임베디드 처리 디테일이 비대해지면 `_embedded_handlers.py`로 분리 검토

다음으로 **Chunk 3 (Phase 4~5: reporter_md + reporter_html)**로 진행.

---

## Chunk 3: Phase 4~5 (reporter_md + reporter_html)

**[Chunk 3 cwd 면책]** Chunk 1·2와 동일. 모든 Bash 명령은 `cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"` 직후 실행을 가정.

**[Chunk 3 핵심 설계 결정]**

| 항목 | 결정 |
|---|---|
| 입력 | `findings.json` + `extracted.json` (별도 로드, SSoT는 `slide_summaries.json`이지만 Markdown/HTML은 두 파일만 필요) |
| 마크다운 구조 | (1) 요약 헤더 → (2) 심각도별 섹션 → (3) 카테고리별 섹션 → (4) 슬라이드별 섹션 |
| HTML 구조 | (1) 요약 카드 → (2) 슬라이드 카드 그리드 (썸네일 + 위치 박스 + 이슈 카드) → (3) 카테고리·심각도 필터 |
| 위치 박스 오버레이 | 썸네일 위에 `<div style="position:absolute">`로 박스 그리기 (finding의 `position_pct` 기반) |
| 심각도 색상 | Critical=빨강 (#dc2626), Warning=주황 (#f97316), Info=파랑 (#0ea5e9) |
| 빈 findings | "검토 결과 이슈 없음" 친절 메시지 |
| 마크다운 라이브러리 | 표준 라이브러리만 (외부 의존 없음, 직접 문자열 조립) |
| HTML 라이브러리 | Jinja2 (이미 dependencies) |

**[finding의 `position_pct` 필드 forward note]**: spec §5.5.3 보강 — 2단계 카테고리 SA(Chunk 4)가 finding 작성 시 `shape_id`로 `extracted.json`의 도형을 조회하여 그 `position_pct`를 복사. 도형을 특정할 수 없는 finding(예: 슬라이드 전체 결론)은 `position_pct` 생략 가능 — 리포트는 박스 없이 카드만 표시. Chunk 3의 `reporter_html`은 `f.get("position_pct") or {}`로 안전 처리.

---

### Task 4.1: reporter_md — 빈 findings → "이슈 없음" 메시지

**Files:**
- Create: `skill_src/src/reporter_md.py`
- Create: `skill_src/tests/unit/test_reporter_md.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`skill_src/tests/unit/test_reporter_md.py`:
```python
from pathlib import Path
from src.reporter_md import render


def test_empty_findings(tmp_path):
    findings = {"summary": {"total_issues": 0, "by_severity": {}, "by_category": {}}, "findings": []}
    extracted = {"metadata": {"title": "테스트 보고서", "slide_count": 3}, "slides": []}
    out_path = tmp_path / "review.md"
    result = render(findings, extracted, out_path)

    assert result == out_path
    assert out_path.exists()
    text = out_path.read_text(encoding="utf-8")
    assert "이슈 없음" in text or "발견된 이슈가 없습니다" in text
    assert "테스트 보고서" in text
```

- [ ] **Step 2: 테스트 실패 확인**

```bash
.venv/Scripts/pytest tests/unit/test_reporter_md.py -v
```

Expected: FAIL — module not found

- [ ] **Step 3: 최소 구현**

`skill_src/src/reporter_md.py`:
```python
from __future__ import annotations
from pathlib import Path
from typing import Any


def render(findings: dict[str, Any], extracted: dict[str, Any], out_path: Path) -> Path:
    """findings.json + extracted.json → 마크다운 리포트 파일 생성. 출력 경로 반환."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    title = extracted.get("metadata", {}).get("title", "보고서")
    summary = findings.get("summary", {})
    total = summary.get("total_issues", 0)

    lines: list[str] = []
    lines.append(f"# 보고서 검토 결과: {title}")
    lines.append("")
    lines.append(f"슬라이드 수: {extracted.get('metadata', {}).get('slide_count', 0)}")
    lines.append(f"총 이슈: {total}개")
    lines.append("")

    if total == 0:
        lines.append("## 검토 결과 이슈 없음")
        lines.append("")
        lines.append("발견된 이슈가 없습니다. 보고서를 그대로 제출 가능합니다.")
    else:
        lines.append("(상세 내용은 후속 task에서 추가)")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path
```

- [ ] **Step 4: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_reporter_md.py -v
```

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/reporter_md.py skill_src/tests/unit/test_reporter_md.py
git commit -m "feat(reporter_md): 빈 findings용 '이슈 없음' 메시지"
```

---

### Task 4.2: reporter_md — 1개 finding 기본 형식

**Files:**
- Modify: `skill_src/src/reporter_md.py`
- Modify: `skill_src/tests/unit/test_reporter_md.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_single_finding_includes_all_fields(tmp_path):
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"critical": 1}, "by_category": {"data": 1}},
        "findings": [
            {
                "id": "F001",
                "category": "data",
                "severity": "critical",
                "slide_index": 5,
                "shape_id": "s5_sh3",
                "position_hint": "슬라이드 5 우측 상단 텍스트 박스 (좌측 70%, 상단 15%)",
                "quoted_text": "최대 응력 250 MPa",
                "issue": "표 5의 셀에는 240 MPa로 기재되어 있어 본문 인용과 불일치",
                "suggestion": "본문을 240 MPa로 수정",
                "evidence": "slide_5 표(s5_sh4) 행 3·열 2 = '240'",
            }
        ],
    }
    extracted = {"metadata": {"title": "테스트", "slide_count": 5}, "slides": []}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)

    text = out_path.read_text(encoding="utf-8")
    # 모든 필드가 마크다운에 포함되는지
    assert "F001" in text
    assert "data" in text or "데이터" in text
    assert "critical" in text or "Critical" in text or "심각" in text
    assert "슬라이드 5" in text
    assert "최대 응력 250 MPa" in text
    assert "본문 인용과 불일치" in text
    assert "240 MPa로 수정" in text
    assert "행 3·열 2" in text
```

- [ ] **Step 2: 테스트 실패 확인**

Expected: FAIL — finding 내용이 출력에 없음

- [ ] **Step 3: render() 확장 — finding 1개 렌더링**

`reporter_md.py`의 `render` 함수에서 `if total == 0:` 분기 뒤에 다음 추가 (else 블록 교체):
```python
    if total == 0:
        lines.append("## 검토 결과 이슈 없음")
        lines.append("")
        lines.append("발견된 이슈가 없습니다. 보고서를 그대로 제출 가능합니다.")
    else:
        lines.append("## 발견된 이슈")
        lines.append("")
        for f in findings.get("findings", []):
            lines.extend(_format_finding(f))
            lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


_SEVERITY_KO = {"critical": "🔴 Critical", "warning": "🟠 Warning", "info": "🔵 Info"}
_CATEGORY_KO = {
    "typo": "오타",
    "terminology": "용어 통일",
    "data": "데이터",
    "conclusion": "결론 검증",
    "improvement": "개선 제안",
    "logic": "논리·강도",
}


def _format_finding(f: dict[str, Any]) -> list[str]:
    """단일 finding을 마크다운 블록(라인 리스트)으로."""
    sev = _SEVERITY_KO.get(f.get("severity", "info"), f.get("severity", "info"))
    cat = _CATEGORY_KO.get(f.get("category", ""), f.get("category", ""))
    block = [
        f"### [{f.get('id', '?')}] {sev} · {cat} · {f.get('position_hint', '')}",
        "",
        f"**원문 인용**: \"{f.get('quoted_text', '')}\"",
        "",
        f"**문제**: {f.get('issue', '')}",
        "",
        f"**개선 제안**: {f.get('suggestion', '')}",
    ]
    evidence = f.get("evidence")
    if evidence:
        block.extend(["", f"**근거**: {evidence}"])
    return block
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS (기존 + 신규 모두)

- [ ] **Step 5: commit**

```bash
git add skill_src/src/reporter_md.py skill_src/tests/unit/test_reporter_md.py
git commit -m "feat(reporter_md): 단일 finding 렌더 (id/심각도/카테고리/위치/인용/제안/근거)"
```

---

### Task 4.3: reporter_md — 슬라이드별 그룹 섹션 추가

**Files:**
- Modify: `skill_src/src/reporter_md.py`
- Modify: `skill_src/tests/unit/test_reporter_md.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_findings_grouped_by_slide(tmp_path):
    findings = {
        "summary": {"total_issues": 3, "by_severity": {"critical": 1, "warning": 2}, "by_category": {"typo": 2, "data": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "슬라이드 1", "quoted_text": "오타1", "issue": "...", "suggestion": "..."},
            {"id": "F002", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2", "quoted_text": "오타2", "issue": "...", "suggestion": "..."},
            {"id": "F003", "category": "data", "severity": "critical", "slide_index": 1,
             "shape_id": "s1_sh3", "position_hint": "슬라이드 1", "quoted_text": "값", "issue": "...", "suggestion": "..."},
        ],
    }
    extracted = {"metadata": {"title": "테스트", "slide_count": 3}, "slides": [
        {"index": 1, "title": "표지"}, {"index": 2, "title": "본문"}, {"index": 3, "title": "결론"}
    ]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")

    # 슬라이드별 섹션 존재
    assert "## 슬라이드별 이슈" in text
    assert "### 슬라이드 1" in text or "### 슬라이드 1: 표지" in text
    assert "### 슬라이드 2" in text or "### 슬라이드 2: 본문" in text
    # 슬라이드 1에 F001과 F003 둘 다 나와야 함
    s1_section_start = text.find("### 슬라이드 1")
    s2_section_start = text.find("### 슬라이드 2")
    s1_block = text[s1_section_start:s2_section_start]
    assert "F001" in s1_block
    assert "F003" in s1_block
```

- [ ] **Step 2: 테스트 실패 확인**

- [ ] **Step 3: 슬라이드별 그룹 섹션 추가 — `render()` 수정**

`reporter_md.py`의 `render` 함수에서 `## 발견된 이슈` 섹션 뒤에 슬라이드별 섹션 추가:
```python
    else:
        lines.append("## 발견된 이슈")
        lines.append("")
        for f in findings.get("findings", []):
            lines.extend(_format_finding(f))
            lines.append("")

        # 슬라이드별 그룹
        lines.append("## 슬라이드별 이슈")
        lines.append("")
        slides_meta = {s["index"]: s.get("title", "") for s in extracted.get("slides", [])}
        by_slide: dict[int, list[dict[str, Any]]] = {}
        for f in findings.get("findings", []):
            by_slide.setdefault(f.get("slide_index", 0), []).append(f)
        for slide_idx in sorted(by_slide.keys()):
            title = slides_meta.get(slide_idx, "")
            heading = f"### 슬라이드 {slide_idx}"
            if title:
                heading += f": {title}"
            lines.append(heading)
            lines.append("")
            for f in by_slide[slide_idx]:
                sev = _SEVERITY_KO.get(f.get("severity", ""), "")
                cat = _CATEGORY_KO.get(f.get("category", ""), "")
                lines.append(f"- [{f.get('id', '?')}] {sev} · {cat} · {f.get('issue', '')}")
            lines.append("")
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/reporter_md.py skill_src/tests/unit/test_reporter_md.py
git commit -m "feat(reporter_md): 슬라이드별 그룹 섹션 추가"
```

---

### Task 4.4: reporter_md — 카테고리별 그룹 + 요약 헤더 보강

**Files:**
- Modify: `skill_src/src/reporter_md.py`
- Modify: `skill_src/tests/unit/test_reporter_md.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_summary_header_and_category_grouping(tmp_path):
    findings = {
        "summary": {
            "total_issues": 3,
            "by_severity": {"critical": 1, "warning": 2, "info": 0},
            "by_category": {"typo": 2, "data": 1, "terminology": 0, "conclusion": 0, "improvement": 0, "logic": 0},
        },
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "슬라이드 1", "quoted_text": "x", "issue": "i1", "suggestion": "s1"},
            {"id": "F002", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2", "quoted_text": "y", "issue": "i2", "suggestion": "s2"},
            {"id": "F003", "category": "data", "severity": "critical", "slide_index": 1,
             "shape_id": "s1_sh3", "position_hint": "슬라이드 1", "quoted_text": "z", "issue": "i3", "suggestion": "s3"},
        ],
    }
    extracted = {"metadata": {"title": "T", "slide_count": 3}, "slides": [{"index": i} for i in [1, 2, 3]]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")

    # 요약 헤더에 심각도·카테고리 통계
    assert "Critical: 1" in text or "🔴" in text
    assert "Warning: 2" in text or "🟠" in text
    # 카테고리별 섹션
    assert "## 카테고리별 이슈" in text
    assert "### 오타" in text or "### 오타 (2건)" in text
    assert "### 데이터" in text or "### 데이터 (1건)" in text
```

- [ ] **Step 2: 테스트 실패 확인**

- [ ] **Step 3: render() 수정 — 요약 헤더 + 카테고리 섹션**

`reporter_md.py`의 `render`에서 두 위치를 수정:

**(A) 요약 헤더 통계 — `if total == 0:` 분기 진입 직전에 한 번만 추가** (분기 안에 넣지 말 것):
```python
    # 요약 헤더 보강 (분기 진입 전, 총 이슈 라인 다음)
    by_sev = summary.get("by_severity", {})
    if by_sev:
        sev_parts = []
        for sk, ko in [("critical", "Critical"), ("warning", "Warning"), ("info", "Info")]:
            n = by_sev.get(sk, 0)
            if n > 0:
                sev_parts.append(f"{_SEVERITY_KO.get(sk, ko)}: {n}")
        if sev_parts:
            lines.append("심각도: " + " · ".join(sev_parts))
            lines.append("")

    # 이어서 기존 `if total == 0: ... else: ...` 블록은 그대로 유지
```

**Note**: `_SEVERITY_KO`가 이모지 prefix를 포함하므로 실제 출력은 `심각도: 🔴 Critical: 1 · 🟠 Warning: 2` 형태. 골든 파일(Task 4.5) 검토 체크리스트의 "심각도 통계 형식"은 이 형식을 정답으로 인정.

그리고 슬라이드별 섹션(`## 슬라이드별 이슈`) 다음에 추가:
```python
        # 카테고리별 그룹
        lines.append("## 카테고리별 이슈")
        lines.append("")
        by_cat: dict[str, list[dict[str, Any]]] = {}
        for f in findings.get("findings", []):
            by_cat.setdefault(f.get("category", ""), []).append(f)
        for cat in sorted(by_cat.keys()):
            cat_ko = _CATEGORY_KO.get(cat, cat)
            lines.append(f"### {cat_ko} ({len(by_cat[cat])}건)")
            lines.append("")
            for f in by_cat[cat]:
                sev = _SEVERITY_KO.get(f.get("severity", ""), "")
                lines.append(f"- [{f.get('id', '?')}] {sev} · 슬라이드 {f.get('slide_index')} · {f.get('issue', '')}")
            lines.append("")
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/reporter_md.py skill_src/tests/unit/test_reporter_md.py
git commit -m "feat(reporter_md): 요약 헤더(심각도 통계) + 카테고리별 그룹"
```

---

### Task 4.5: reporter_md — 통합 골든 파일 회귀 테스트

**Files:**
- Create: `skill_src/tests/fixtures/golden/reporter_md/sample_review.md`
- Modify: `skill_src/tests/unit/test_reporter_md.py`

- [ ] **Step 1: 골든 파일 작성 — 한 번 실행 후 검토**

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
"C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/.venv/Scripts/python" -c "
import json
from pathlib import Path
from src.reporter_md import render

findings = {
    'summary': {
        'total_issues': 2,
        'by_severity': {'critical': 1, 'warning': 1, 'info': 0},
        'by_category': {'typo': 1, 'data': 1},
    },
    'findings': [
        {'id': 'F001', 'category': 'typo', 'severity': 'warning', 'slide_index': 2,
         'shape_id': 's2_sh1', 'position_hint': '슬라이드 2 본문 영역', 'quoted_text': '안응력',
         'issue': '오타: 안응력 → 인장응력으로 보임', 'suggestion': '인장응력으로 수정', 'evidence': ''},
        {'id': 'F002', 'category': 'data', 'severity': 'critical', 'slide_index': 5,
         'shape_id': 's5_sh3', 'position_hint': '슬라이드 5 우측 상단 텍스트 박스', 'quoted_text': '최대 응력 250 MPa',
         'issue': '표 5의 셀에는 240 MPa로 기재됨', 'suggestion': '본문 240 MPa 또는 표 250 MPa로 통일',
         'evidence': '표 5(s5_sh4) 행 3·열 2 = 240'},
    ],
}
extracted = {
    'metadata': {'title': '테스트 보고서', 'slide_count': 5},
    'slides': [{'index': i, 'title': f'슬라이드{i}'} for i in range(1, 6)],
}
out = Path('tests/fixtures/golden/reporter_md/sample_review.md')
out.parent.mkdir(parents=True, exist_ok=True)
render(findings, extracted, out)
print('Saved:', out)
"
```

**검토 체크리스트** (commit 전 확인):
- [ ] 제목 "테스트 보고서" 등장
- [ ] 심각도 통계 "Critical: 1 · Warning: 1" 출력
- [ ] F001, F002 모두 "발견된 이슈" 섹션에 등장
- [ ] 슬라이드 2에 F001, 슬라이드 5에 F002 그룹핑
- [ ] 카테고리 "오타", "데이터" 섹션 분리
- [ ] 마크다운 형식 깨짐 없음

- [ ] **Step 2: 회귀 테스트 추가**

```python
def test_render_matches_golden(tmp_path, fixtures_dir):
    findings = {
        "summary": {"total_issues": 2, "by_severity": {"critical": 1, "warning": 1, "info": 0},
                    "by_category": {"typo": 1, "data": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2 본문 영역", "quoted_text": "안응력",
             "issue": "오타: 안응력 → 인장응력으로 보임", "suggestion": "인장응력으로 수정", "evidence": ""},
            {"id": "F002", "category": "data", "severity": "critical", "slide_index": 5,
             "shape_id": "s5_sh3", "position_hint": "슬라이드 5 우측 상단 텍스트 박스", "quoted_text": "최대 응력 250 MPa",
             "issue": "표 5의 셀에는 240 MPa로 기재됨", "suggestion": "본문 240 MPa 또는 표 250 MPa로 통일",
             "evidence": "표 5(s5_sh4) 행 3·열 2 = 240"},
        ],
    }
    extracted = {"metadata": {"title": "테스트 보고서", "slide_count": 5},
                 "slides": [{"index": i, "title": f"슬라이드{i}"} for i in range(1, 6)]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    actual = out_path.read_text(encoding="utf-8")
    expected = (fixtures_dir / "golden" / "reporter_md" / "sample_review.md").read_text(encoding="utf-8")
    assert actual == expected
```

- [ ] **Step 3: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 4: commit**

```bash
git add skill_src/tests/fixtures/golden/reporter_md/ skill_src/tests/unit/test_reporter_md.py
git commit -m "test(reporter_md): 통합 골든 파일 회귀 테스트"
```

---

### Task 5.1: reporter_html — 빈 findings용 기본 템플릿

**Files:**
- Create: `skill_src/templates/report.html.j2`
- Create: `skill_src/templates/style.css`
- Create: `skill_src/src/reporter_html.py`
- Create: `skill_src/tests/unit/test_reporter_html.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`skill_src/tests/unit/test_reporter_html.py`:
```python
from pathlib import Path
from src.reporter_html import render


def test_empty_findings_creates_html(tmp_path):
    findings = {"summary": {"total_issues": 0, "by_severity": {}, "by_category": {}}, "findings": []}
    extracted = {"metadata": {"title": "T", "slide_count": 0}, "slides": []}
    out_dir = tmp_path / "out"
    result = render(findings, extracted, out_dir)

    assert result.exists()
    assert result.suffix == ".html"
    assert result.name == "review.html"
    text = result.read_text(encoding="utf-8")
    assert "<!DOCTYPE html>" in text
    assert "T" in text
    assert "이슈 없음" in text or "no issues" in text.lower()
    # CSS 복사 확인
    assert (out_dir / "assets" / "style.css").exists()
```

- [ ] **Step 2: 테스트 실패 확인**

- [ ] **Step 3: 템플릿 + 구현 작성**

**Note**: 템플릿(`report.html.j2`)은 Task 5.1에서 한 번에 작성하지만, 변수 중 `slides_with_findings`, `boxes`, `severity_label`, `category_label` 등은 Task 5.2~5.3에서 채워짐. Task 5.1의 reporter_html.py는 빈 리스트만 전달 → 빈 findings 테스트만 통과. 템플릿이 미리 완전히 작성되는 이유는 후속 task에서 템플릿 재수정을 피하기 위함.

`skill_src/templates/style.css`:
```css
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Malgun Gothic", sans-serif; max-width: 1200px; margin: 2em auto; padding: 0 1em; color: #1a1a1a; }
h1 { border-bottom: 2px solid #1a1a1a; padding-bottom: 0.3em; }
.summary { background: #f5f5f5; padding: 1em; border-radius: 8px; margin-bottom: 2em; }
.severity-critical { color: #dc2626; font-weight: bold; }
.severity-warning { color: #f97316; font-weight: bold; }
.severity-info { color: #0ea5e9; font-weight: bold; }
.slide-card { border: 1px solid #ddd; border-radius: 8px; padding: 1em; margin: 1em 0; }
.slide-thumb-wrap { position: relative; display: inline-block; }
.slide-thumb-wrap img { display: block; max-width: 480px; height: auto; }
.position-box { position: absolute; border: 2px solid #dc2626; background: rgba(220, 38, 38, 0.1); pointer-events: none; }
.finding-card { border-left: 4px solid #888; padding: 0.5em 1em; margin: 0.5em 0; background: #fafafa; }
.finding-card.critical { border-left-color: #dc2626; }
.finding-card.warning { border-left-color: #f97316; }
.finding-card.info { border-left-color: #0ea5e9; }
.empty { text-align: center; padding: 3em; color: #666; }
```

`skill_src/templates/report.html.j2`:
```html
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>보고서 검토 결과: {{ title }}</title>
<link rel="stylesheet" href="assets/style.css">
</head>
<body>
<h1>보고서 검토 결과: {{ title }}</h1>
<div class="summary">
  <p>슬라이드 수: {{ slide_count }}</p>
  <p>총 이슈: {{ total_issues }}개</p>
  {% if total_issues > 0 and severity_counts %}
  <p>심각도:
    {% for sev_key, sev_label, count in severity_counts %}
      <span class="severity-{{ sev_key }}">{{ sev_label }}: {{ count }}</span>{% if not loop.last %} · {% endif %}
    {% endfor %}
  </p>
  {% endif %}
</div>

{% if total_issues == 0 %}
<div class="empty">
  <h2>검토 결과 이슈 없음</h2>
  <p>발견된 이슈가 없습니다.</p>
</div>
{% else %}
{% for slide in slides_with_findings %}
<div class="slide-card">
  <h3>슬라이드 {{ slide.index }}{% if slide.title %}: {{ slide.title }}{% endif %}</h3>
  {% if slide.thumbnail_rel %}
  <div class="slide-thumb-wrap">
    <img src="{{ slide.thumbnail_rel }}" alt="슬라이드 {{ slide.index }}">
    {% for box in slide.boxes %}
    <div class="position-box" style="left:{{ box.left }}%; top:{{ box.top }}%; width:{{ box.width }}%; height:{{ box.height }}%;"></div>
    {% endfor %}
  </div>
  {% endif %}
  {% for f in slide.findings %}
  <div class="finding-card {{ f.severity }}">
    <strong>[{{ f.id }}] {{ f.severity_label }} · {{ f.category_label }}</strong>
    <p>원문: "{{ f.quoted_text }}"</p>
    <p>문제: {{ f.issue }}</p>
    <p>개선 제안: {{ f.suggestion }}</p>
    {% if f.evidence %}<p>근거: {{ f.evidence }}</p>{% endif %}
  </div>
  {% endfor %}
</div>
{% endfor %}
{% endif %}
</body>
</html>
```

`skill_src/src/reporter_html.py`:
```python
from __future__ import annotations
import shutil
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape


_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"

_SEVERITY_LABEL = {"critical": "Critical", "warning": "Warning", "info": "Info"}
_CATEGORY_LABEL = {
    "typo": "오타", "terminology": "용어 통일", "data": "데이터",
    "conclusion": "결론 검증", "improvement": "개선 제안", "logic": "논리·강도",
}


def render(findings: dict[str, Any], extracted: dict[str, Any], out_dir: Path) -> Path:
    """findings.json + extracted.json → HTML 리포트 파일 생성. 출력 HTML 경로 반환."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = out_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    # CSS 복사
    css_src = _TEMPLATES_DIR / "style.css"
    css_dst = assets_dir / "style.css"
    shutil.copy(css_src, css_dst)

    env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)), autoescape=select_autoescape(["html"]))
    template = env.get_template("report.html.j2")

    title = extracted.get("metadata", {}).get("title", "보고서")
    slide_count = extracted.get("metadata", {}).get("slide_count", 0)
    summary = findings.get("summary", {})
    total = summary.get("total_issues", 0)

    severity_counts = []
    by_sev = summary.get("by_severity", {})
    for sk in ("critical", "warning", "info"):
        if by_sev.get(sk, 0) > 0:
            severity_counts.append((sk, _SEVERITY_LABEL[sk], by_sev[sk]))

    rendered = template.render(
        title=title,
        slide_count=slide_count,
        total_issues=total,
        severity_counts=severity_counts,
        slides_with_findings=[],  # Task 5.2에서 채움
    )
    out_path = out_dir / "review.html"
    out_path.write_text(rendered, encoding="utf-8")
    return out_path
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/templates/ skill_src/src/reporter_html.py skill_src/tests/unit/test_reporter_html.py
git commit -m "feat(reporter_html): Jinja2 기본 템플릿 + 빈 findings 처리 + CSS"
```

---

### Task 5.2: reporter_html — 슬라이드 카드 + 썸네일 임베드

**Files:**
- Modify: `skill_src/src/reporter_html.py`
- Modify: `skill_src/tests/unit/test_reporter_html.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_html_includes_slide_cards_with_thumbnails(tmp_path):
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"warning": 1}, "by_category": {"typo": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2", "quoted_text": "오타",
             "issue": "오타입니다", "suggestion": "수정 필요", "evidence": ""}
        ],
    }
    # 가짜 썸네일 경로
    thumb = tmp_path / "ws" / "thumbnails" / "slide_002.jpg"
    thumb.parent.mkdir(parents=True, exist_ok=True)
    thumb.write_bytes(b"fake jpg")
    extracted = {
        "metadata": {"title": "T", "slide_count": 3},
        "slides": [
            {"index": 1, "title": "표지", "thumbnail_path": None},
            {"index": 2, "title": "본문", "thumbnail_path": str(thumb)},
            {"index": 3, "title": "결론", "thumbnail_path": None},
        ],
    }
    out_dir = tmp_path / "out"
    render(findings, extracted, out_dir)

    text = (out_dir / "review.html").read_text(encoding="utf-8")
    # 슬라이드 2 카드만 존재 (이슈가 있는 슬라이드)
    assert "슬라이드 2" in text
    assert "F001" in text
    # 썸네일이 assets로 복사되어 상대 경로로 참조
    assert (out_dir / "assets" / "thumbnails" / "slide_002.jpg").exists()
    assert "assets/thumbnails/slide_002.jpg" in text
```

- [ ] **Step 2: 테스트 실패 확인**

- [ ] **Step 3: render() 수정 — 슬라이드 카드 빌드 + 썸네일 복사**

`reporter_html.py`의 `render` 함수 끝부분(template.render 직전)에서 `slides_with_findings` 채우기 + 썸네일 복사:
```python
def render(findings, extracted, out_dir):
    # ... (기존 부분)

    # 이슈가 있는 슬라이드만 카드로
    slides_meta = {s["index"]: s for s in extracted.get("slides", [])}
    by_slide: dict[int, list[dict]] = {}
    for f in findings.get("findings", []):
        by_slide.setdefault(f.get("slide_index", 0), []).append(f)

    thumb_out_dir = assets_dir / "thumbnails"
    thumb_out_dir.mkdir(parents=True, exist_ok=True)

    slides_with_findings = []
    for slide_idx in sorted(by_slide.keys()):
        meta = slides_meta.get(slide_idx, {})
        thumb_rel = None
        thumb_src = meta.get("thumbnail_path")
        if thumb_src and Path(thumb_src).exists():
            thumb_dst = thumb_out_dir / f"slide_{slide_idx:03d}.jpg"
            shutil.copy(thumb_src, thumb_dst)
            thumb_rel = f"assets/thumbnails/slide_{slide_idx:03d}.jpg"

        formatted_findings = []
        for f in by_slide[slide_idx]:
            formatted_findings.append({
                "id": f.get("id", "?"),
                "severity": f.get("severity", "info"),
                "severity_label": _SEVERITY_LABEL.get(f.get("severity", "info"), ""),
                "category_label": _CATEGORY_LABEL.get(f.get("category", ""), f.get("category", "")),
                "quoted_text": f.get("quoted_text", ""),
                "issue": f.get("issue", ""),
                "suggestion": f.get("suggestion", ""),
                "evidence": f.get("evidence", ""),
            })

        slides_with_findings.append({
            "index": slide_idx,
            "title": meta.get("title", ""),
            "thumbnail_rel": thumb_rel,
            "boxes": [],  # Task 5.3에서 채움
            "findings": formatted_findings,
        })

    rendered = template.render(
        title=title,
        slide_count=slide_count,
        total_issues=total,
        severity_counts=severity_counts,
        slides_with_findings=slides_with_findings,
    )
    # ... (out_path.write_text 그대로)
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS

- [ ] **Step 5: commit**

```bash
git add skill_src/src/reporter_html.py skill_src/tests/unit/test_reporter_html.py
git commit -m "feat(reporter_html): 슬라이드 카드 + 썸네일 임베드"
```

---

### Task 5.3: reporter_html — 위치 박스 오버레이

**Files:**
- Modify: `skill_src/src/reporter_html.py`
- Modify: `skill_src/tests/unit/test_reporter_html.py`

- [ ] **Step 1: 실패하는 테스트 추가**

```python
def test_html_includes_position_boxes(tmp_path):
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"warning": 1}, "by_category": {"typo": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh3", "position_hint": "슬라이드 1 우측 상단",
             "position_pct": {"left": 0.7, "top": 0.15, "width": 0.25, "height": 0.10},
             "quoted_text": "x", "issue": "i", "suggestion": "s", "evidence": ""}
        ],
    }
    thumb = tmp_path / "ws" / "thumbnails" / "slide_001.jpg"
    thumb.parent.mkdir(parents=True, exist_ok=True)
    thumb.write_bytes(b"fake")
    extracted = {
        "metadata": {"title": "T", "slide_count": 1},
        "slides": [{"index": 1, "title": "표지", "thumbnail_path": str(thumb)}],
    }
    out_dir = tmp_path / "out"
    render(findings, extracted, out_dir)

    text = (out_dir / "review.html").read_text(encoding="utf-8")
    # position-box div가 left:70%, top:15%로 출력되는지
    assert "position-box" in text
    assert "left:70" in text or "left: 70" in text
    assert "top:15" in text or "top: 15" in text
    assert "width:25" in text or "width: 25" in text
```

- [ ] **Step 2: 테스트 실패 확인**

- [ ] **Step 3: render() 수정 — position_pct 기반 박스 좌표 추가**

`reporter_html.py`의 `slides_with_findings.append` 직전 부분 수정:
```python
        boxes = []
        for f in by_slide[slide_idx]:
            pct = f.get("position_pct") or {}
            if "left" in pct and "top" in pct:
                boxes.append({
                    "left": int(pct["left"] * 100),
                    "top": int(pct["top"] * 100),
                    "width": int(pct.get("width", 0) * 100),
                    "height": int(pct.get("height", 0) * 100),
                })

        slides_with_findings.append({
            "index": slide_idx,
            "title": meta.get("title", ""),
            "thumbnail_rel": thumb_rel,
            "boxes": boxes,  # ← 채움
            "findings": formatted_findings,
        })
```

- [ ] **Step 4: 테스트 통과 확인**

Expected: PASS (템플릿이 이미 boxes 처리)

- [ ] **Step 5: commit**

```bash
git add skill_src/src/reporter_html.py skill_src/tests/unit/test_reporter_html.py
git commit -m "feat(reporter_html): 슬라이드 썸네일 위 위치 박스 오버레이"
```

---

### Task 5.4: reporter_html — 통합 회귀 테스트 + 빌드 검증

**Files:**
- Modify: `skill_src/tests/unit/test_reporter_html.py`

- [ ] **Step 1: 통합 회귀 테스트 추가**

```python
def test_html_full_integration(tmp_path):
    """모든 카테고리·심각도가 섞인 findings로 HTML 렌더 + DOM 구조 검증."""
    thumb1 = tmp_path / "ws" / "thumbnails" / "slide_001.jpg"
    thumb2 = tmp_path / "ws" / "thumbnails" / "slide_002.jpg"
    for t in (thumb1, thumb2):
        t.parent.mkdir(parents=True, exist_ok=True)
        t.write_bytes(b"fake")

    findings = {
        "summary": {
            "total_issues": 3,
            "by_severity": {"critical": 1, "warning": 1, "info": 1},
            "by_category": {"typo": 1, "data": 1, "logic": 1},
        },
        "findings": [
            {"id": "F001", "category": "typo", "severity": "info", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "S1", "position_pct": {"left":0.1,"top":0.1,"width":0.3,"height":0.1},
             "quoted_text": "오타", "issue": "i1", "suggestion": "s1", "evidence": ""},
            {"id": "F002", "category": "data", "severity": "critical", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "S2", "position_pct": {"left":0.5,"top":0.5,"width":0.4,"height":0.2},
             "quoted_text": "데이터", "issue": "i2", "suggestion": "s2", "evidence": "표 ..."},
            {"id": "F003", "category": "logic", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh2", "position_hint": "S2-2", "position_pct": {"left":0.0,"top":0.8,"width":1.0,"height":0.2},
             "quoted_text": "결론", "issue": "i3", "suggestion": "s3", "evidence": ""},
        ],
    }
    extracted = {
        "metadata": {"title": "통합 테스트", "slide_count": 3},
        "slides": [
            {"index": 1, "title": "표지", "thumbnail_path": str(thumb1)},
            {"index": 2, "title": "본문", "thumbnail_path": str(thumb2)},
            {"index": 3, "title": "결론", "thumbnail_path": None},
        ],
    }
    out_dir = tmp_path / "out"
    render(findings, extracted, out_dir)

    text = (out_dir / "review.html").read_text(encoding="utf-8")
    # 3개 finding 모두 등장
    for fid in ("F001", "F002", "F003"):
        assert fid in text
    # 심각도별 통계
    assert "Critical: 1" in text
    assert "Warning: 1" in text
    assert "Info: 1" in text
    # 슬라이드 1, 2 카드만 (3은 이슈 없음)
    assert text.count('<div class="slide-card">') == 2
    # 위치 박스 3개 (각 finding당 1개)
    assert text.count("position-box") == 3
    # CSS·썸네일 자산 모두 복사됨
    assert (out_dir / "assets" / "style.css").exists()
    assert (out_dir / "assets" / "thumbnails" / "slide_001.jpg").exists()
    assert (out_dir / "assets" / "thumbnails" / "slide_002.jpg").exists()
```

- [ ] **Step 2: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_reporter_html.py -v
```

Expected: 모든 테스트 PASS

- [ ] **Step 3: commit**

```bash
git add skill_src/tests/unit/test_reporter_html.py
git commit -m "test(reporter_html): 통합 회귀 테스트 (3 finding, 2 슬라이드, CSS·썸네일 검증)"
```

---

## Chunk 3 완료 시점

이 시점에:
- `reporter_md.py`: 빈 findings 처리 + 단일 finding + 슬라이드별/카테고리별 그룹 + 요약 헤더 완성
- `reporter_html.py`: Jinja2 템플릿 + 슬라이드 카드 + 썸네일 + 위치 박스 오버레이 + CSS 자산 복사
- 통합 골든 파일 회귀 테스트 1건 (마크다운)
- 통합 회귀 테스트 1건 (HTML)
- 모든 단위 테스트 통과

**향후 확장 메모:**
- 카테고리별 색상·아이콘 추가 시 CSS·`_CATEGORY_LABEL` 확장
- 슬라이드별 가나다 정렬·필터링은 JavaScript로 추후 (vanilla JS, 외부 의존 없음)
- 마크다운에서 카테고리별 finding의 detail은 현재 한 줄 요약 — 필요 시 detail 펼침 추가

다음으로 **Chunk 4 (Phase 6~7: CLI + SKILL.md + 7 Subagents + 글로벌 배포)**로 진행.

---

## Chunk 4: Phase 6~7 (CLI + SKILL.md + 7 Subagents + 글로벌 배포)

**[Chunk 4 cwd 면책]** 이 chunk의 Bash 명령은 task별 헤더의 cwd를 따른다 (개발 시 `skill_src/`, 배포 task에서 `~/.claude/...`).

**[Chunk 4 핵심 설계 결정 — spec §12 미결정 항목 해결]**

| 항목 | 결정 |
|---|---|
| **1단계 SA 입력 전달 방식** | **임시 파일 경로 전달** (`<work_dir>/slide_inputs/slide_NNN.json`). Task prompt에 인라인 JSON 넣지 않음 — 슬라이드별 입력 토큰 절약. SA는 Read로 JSON 읽고, 슬라이드 이미지도 Read로 읽음 |
| **2단계 SA의 extracted.json Read 범위** | **`slide_summaries.json` 전체 + `extracted.json` 전체 Read 권한** (1차 MVP). 슬라이드 단위 분할은 2차 이후 검토 (현재 보고서 평균 30장이면 충분히 처리 가능) |
| **6개 카테고리 SA 결과 병합** | **각 SA가 자기 카테고리 ID prefix 사용** (typo→T001, terminology→TM001, data→D001, conclusion→C001, improvement→I001, logic→L001). 메인 Claude가 단순 concat. ID 충돌 없음 |
| **이미지 토큰 한계 감지** | **사전 측정** (Pillow로 이미지 크기 → 토큰 추정 약 1.6×width×height/750) + **API 에러 catch** 시 다운스케일 재시도 (1928→1280→960). 메인 Claude가 SKILL.md 가이드에 따라 처리 |
| **Subagent 도구 권한** | 모두 `Read`만. Bash·Write 권한 없음 (JSON 결과는 stdout으로 반환) |
| **Subagent 출력 파싱** | 메인 Claude가 SA 응답에서 JSON 코드 블록 추출 → `json.loads` |

---

### Task 6.1: extract_and_render.py CLI

**Files:**
- Create: `skill_src/extract_and_render.py`
- Create: `skill_src/tests/integration/test_cli_extract_and_render.py`

- [ ] **Step 1: 실패하는 통합 테스트 작성**

`skill_src/tests/integration/test_cli_extract_and_render.py`:
```python
import json
import sys
import subprocess
from pathlib import Path
import pytest

pytestmark = pytest.mark.real_com  # PowerPoint 필요


def test_cli_creates_extracted_json_and_images(sample_text_only: Path, tmp_path: Path):
    work_dir = tmp_path / "ws"
    skill_dir = Path(__file__).parent.parent.parent  # skill_src/
    cli = skill_dir / "extract_and_render.py"

    result = subprocess.run(
        [sys.executable, str(cli), str(sample_text_only), "--out", str(work_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )
    assert result.returncode == 0, f"stderr: {result.stderr}"

    extracted_path = work_dir / "extracted.json"
    assert extracted_path.exists()
    data = json.loads(extracted_path.read_text(encoding="utf-8"))
    assert data["metadata"]["slide_count"] == 5
    # 슬라이드별 image_path가 채워졌는지
    for s in data["slides"]:
        assert s["image_path"] is not None
        assert Path(s["image_path"]).exists()
        assert s["thumbnail_path"] is not None
        assert Path(s["thumbnail_path"]).exists()


def test_cli_creates_slide_input_jsons(sample_text_only: Path, tmp_path: Path):
    """1단계 SA가 읽을 슬라이드별 입력 파일도 생성."""
    work_dir = tmp_path / "ws"
    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "extract_and_render.py"

    subprocess.run(
        [sys.executable, str(cli), str(sample_text_only), "--out", str(work_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )

    inputs_dir = work_dir / "slide_inputs"
    assert inputs_dir.exists()
    inputs = sorted(inputs_dir.glob("slide_*.json"))
    assert len(inputs) == 5
    # 각 입력 파일에 슬라이드 데이터 + 이미지 경로
    sample = json.loads(inputs[0].read_text(encoding="utf-8"))
    assert "index" in sample
    assert "shapes" in sample
    assert "image_path" in sample
```

- [ ] **Step 2: extract_and_render.py 작성**

`skill_src/extract_and_render.py`:
```python
"""CLI: PPTX → extracted.json + 슬라이드 이미지 + 슬라이드별 SA 입력 파일."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

from src.config import Config
from src.extractor import extract
from src.slide_renderer import render


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="PPTX 추출 + 슬라이드 이미지 변환 + 1단계 SA 입력 생성"
    )
    parser.add_argument("pptx_path", type=Path, help="입력 PPTX 파일 경로")
    parser.add_argument("--out", type=Path, required=True, help="출력 작업 디렉토리")
    args = parser.parse_args(argv)

    pptx_path: Path = args.pptx_path.resolve()
    work_dir: Path = args.out.resolve()
    work_dir.mkdir(parents=True, exist_ok=True)

    if not pptx_path.exists():
        print(f"ERROR: PPTX 파일 없음: {pptx_path}", file=sys.stderr)
        return 2

    cfg = Config.from_env(dict(__import__("os").environ))

    # ① 구조 추출
    print(f"[1/3] 구조 추출 중: {pptx_path.name}")
    extracted = extract(pptx_path)

    # ② 슬라이드 이미지 변환 (자동 모드 선택)
    print(f"[2/3] 슬라이드 이미지 변환 중 (모드 자동 선택)")
    rendered = render(
        pptx_path, work_dir, extracted,
        width=cfg.image_max_dim, height=int(cfg.image_max_dim * 9 / 16),
        pdf_dpi=cfg.pdf_dpi, thumbnail_max_dim=cfg.thumbnail_max_dim,
    )

    # extracted에 image_path / thumbnail_path / embedded_image_paths 채우기
    thumb_dir = work_dir / "thumbnails"
    for slide in extracted["slides"]:
        idx = slide["index"]
        paths = rendered.get(idx, [])
        if paths:
            slide["image_path"] = str(paths[0])
            slide["embedded_image_paths"] = [str(p) for p in paths[1:]]
        thumb = thumb_dir / f"slide_{idx:03d}.jpg"
        if thumb.exists():
            slide["thumbnail_path"] = str(thumb)

    # extracted.json 저장
    extracted_path = work_dir / "extracted.json"
    extracted_path.write_text(
        json.dumps(extracted, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  → {extracted_path}")

    # ③ 슬라이드별 SA 입력 JSON 생성
    print(f"[3/3] 슬라이드별 SA 입력 JSON 생성")
    inputs_dir = work_dir / "slide_inputs"
    inputs_dir.mkdir(parents=True, exist_ok=True)
    for slide in extracted["slides"]:
        idx = slide["index"]
        target = inputs_dir / f"slide_{idx:03d}.json"
        target.write_text(
            json.dumps(slide, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(f"  → {inputs_dir}/ ({len(extracted['slides'])}개)")

    print("완료. 다음 단계: SKILL.md 워크플로 따라 1단계 SA dispatch.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 3: 단위 테스트도 가벼운 것 추가** (CLI argparse 검증)

`skill_src/tests/unit/test_cli_extract.py`:
```python
import sys
from pathlib import Path
import pytest


def test_cli_missing_pptx_returns_error(tmp_path, monkeypatch):
    """존재하지 않는 PPTX 입력 시 종료 코드 2."""
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))
    from extract_and_render import main

    rc = main([str(tmp_path / "nonexistent.pptx"), "--out", str(tmp_path / "out")])
    assert rc == 2
```

- [ ] **Step 4: 테스트 실행**

```bash
.venv/Scripts/pytest tests/unit/test_cli_extract.py -v
```

Expected: PASS

```bash
.venv/Scripts/pytest tests/integration/test_cli_extract_and_render.py -v -m real_com
```

Expected: PASS (Windows + MS PowerPoint 환경)

- [ ] **Step 5: commit**

```bash
git add skill_src/extract_and_render.py skill_src/tests/unit/test_cli_extract.py skill_src/tests/integration/test_cli_extract_and_render.py
git commit -m "feat(cli): extract_and_render.py — 추출+이미지+SA 입력 파일 생성"
```

---

### Task 6.2: render_report.py CLI

**Files:**
- Create: `skill_src/render_report.py`
- Create: `skill_src/tests/integration/test_cli_render_report.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`skill_src/tests/integration/test_cli_render_report.py`:
```python
import json
import sys
import subprocess
from pathlib import Path


def test_render_report_creates_md_and_html(tmp_path):
    """findings.json + extracted.json → review.md + review.html."""
    work_dir = tmp_path / "ws"
    work_dir.mkdir()
    out_dir = tmp_path / "out"

    # 가짜 입력
    extracted = {"metadata": {"title": "T", "slide_count": 1}, "slides": [{"index": 1, "title": "표지", "thumbnail_path": None}]}
    findings = {"summary": {"total_issues": 0, "by_severity": {}, "by_category": {}}, "findings": []}
    (work_dir / "extracted.json").write_text(json.dumps(extracted, ensure_ascii=False), encoding="utf-8")
    (work_dir / "findings.json").write_text(json.dumps(findings, ensure_ascii=False), encoding="utf-8")

    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "render_report.py"
    result = subprocess.run(
        [sys.executable, str(cli), str(work_dir), "--out", str(out_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )
    assert result.returncode == 0, f"stderr: {result.stderr}"
    assert (out_dir / "review.md").exists()
    assert (out_dir / "review.html").exists()
    assert (out_dir / "assets" / "style.css").exists()


def test_render_report_missing_findings_errors(tmp_path):
    """findings.json 없으면 에러."""
    work_dir = tmp_path / "ws"
    work_dir.mkdir()
    (work_dir / "extracted.json").write_text("{}", encoding="utf-8")
    out_dir = tmp_path / "out"

    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "render_report.py"
    result = subprocess.run(
        [sys.executable, str(cli), str(work_dir), "--out", str(out_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )
    assert result.returncode != 0
    assert "findings.json" in result.stderr
```

- [ ] **Step 2: render_report.py 작성**

`skill_src/render_report.py`:
```python
"""CLI: <work_dir>/extracted.json + findings.json → review.md + review.html."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

from src.reporter_md import render as render_md
from src.reporter_html import render as render_html


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="findings.json + extracted.json → 마크다운·HTML 리포트"
    )
    parser.add_argument("work_dir", type=Path, help="extract_and_render.py가 생성한 작업 디렉토리")
    parser.add_argument("--out", type=Path, required=True, help="리포트 출력 디렉토리")
    args = parser.parse_args(argv)

    work_dir: Path = args.work_dir.resolve()
    out_dir: Path = args.out.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    extracted_path = work_dir / "extracted.json"
    findings_path = work_dir / "findings.json"

    if not extracted_path.exists():
        print(f"ERROR: extracted.json 없음: {extracted_path}", file=sys.stderr)
        return 2
    if not findings_path.exists():
        print(f"ERROR: findings.json 없음: {findings_path}", file=sys.stderr)
        return 2

    extracted = json.loads(extracted_path.read_text(encoding="utf-8"))
    findings = json.loads(findings_path.read_text(encoding="utf-8"))

    md_path = render_md(findings, extracted, out_dir / "review.md")
    html_path = render_html(findings, extracted, out_dir)

    print(f"Markdown: {md_path}")
    print(f"HTML: {html_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 3: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/integration/test_cli_render_report.py -v
```

Expected: PASS

- [ ] **Step 4: commit**

```bash
git add skill_src/render_report.py skill_src/tests/integration/test_cli_render_report.py
git commit -m "feat(cli): render_report.py — extracted+findings → MD/HTML"
```

---

### Task 7.1: SKILL.md 워크플로 가이드 작성

**Files:**
- Create: `skill_src/SKILL.md`

SKILL.md는 Claude(메인)가 따를 워크플로를 마크다운으로 명시하는 가이드 파일. 테스트 어려움 — 파일 존재 + frontmatter + 핵심 워크플로 단계 명시 검증.

- [ ] **Step 1: SKILL.md 작성**

`skill_src/SKILL.md`:
```markdown
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
  - prompt에 슬라이드별 입력 파일 경로 + 이미지 경로 포함:
    ```
    슬라이드 i를 분석해주세요.
    - 슬라이드 데이터 JSON: <work_dir>/slide_inputs/slide_<i:03d>.json
    - 슬라이드 이미지: <work_dir>/slides/slide_<i:03d>.jpg
    - 임베디드 이미지(있으면): <work_dir>/slides/embedded/...
    Read 도구로 위 파일을 읽고, 자기 시스템 프롬프트에 정의된 JSON 스키마로 응답하세요.
    ```
- 각 SA 응답에서 JSON 코드 블록 추출 → 누적
- 모든 결과를 `review_ws/slide_summaries.json`에 저장:
  ```json
  {"slides": [<slide_summary>, <slide_summary>, ...]}
  ```

### Step 3: 2단계 SA 동시 dispatch (6개 카테고리)
- 6개 SA를 동시 dispatch:
  - `report-reviewer-typo` (ID prefix `T`)
  - `report-reviewer-terminology` (ID prefix `TM`)
  - `report-reviewer-data` (ID prefix `D`)
  - `report-reviewer-conclusion` (ID prefix `C`)
  - `report-reviewer-improvement` (ID prefix `I`)
  - `report-reviewer-logic` (ID prefix `L`)
- 각 SA prompt:
  ```
  자기 카테고리에 해당하는 검토를 수행하세요.
  - 슬라이드 요약: <work_dir>/slide_summaries.json
  - 원본 추출: <work_dir>/extracted.json (필요 시 참조)
  - shape_id로 도형 위치 조회 후 finding의 position_pct에 복사
  - ID는 자기 카테고리 prefix를 사용 (예: T001, T002, ...)
  - 응답은 JSON 배열 (findings 스키마)
  ```
- 6개 SA 응답을 단순 concat하여 `findings[]` 생성 (ID prefix 다르므로 충돌 없음)
- 심각도·카테고리 통계 집계 후 `review_ws/findings.json` 저장:
  ```json
  {
    "summary": {
      "total_issues": <합계>,
      "by_severity": {"critical": ..., "warning": ..., "info": ...},
      "by_category": {"typo": ..., ..., "logic": ...}
    },
    "findings": [...]
  }
  ```

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
- **Subagent 실패**: 1회 재시도, 그래도 실패 시 placeholder 결과로 채우고 리포트에 명시

## 옵션 처리
- `--resume`: extracted.json 있으면 Step 1 건너뜀, slide_summaries.json 있으면 Step 2도 건너뜀
- `--rerun-stage2`: extracted.json + slide_summaries.json 사용, Step 3부터
- `--rerender`: findings.json 사용, Step 4만
```

- [ ] **Step 2: 단순 검증 테스트**

`skill_src/tests/unit/test_skill_md.py`:
```python
from pathlib import Path
import re


def test_skill_md_has_frontmatter():
    skill_md = Path(__file__).parent.parent.parent / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    assert text.startswith("---\n")
    # frontmatter 추출
    end = text.find("\n---\n", 4)
    assert end > 0, "frontmatter 종료 마커 없음"
    fm = text[4:end]
    assert "name: report-reviewer" in fm
    assert "description:" in fm


def test_skill_md_workflow_steps_present():
    skill_md = Path(__file__).parent.parent.parent / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    for step_heading in [
        "### Step 0", "### Step 1", "### Step 2",
        "### Step 3", "### Step 4", "### Step 5",
    ]:
        assert step_heading in text, f"누락: {step_heading}"


def test_skill_md_subagent_names():
    skill_md = Path(__file__).parent.parent.parent / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    for sa in [
        "report-reviewer-slide-analyzer",
        "report-reviewer-typo",
        "report-reviewer-terminology",
        "report-reviewer-data",
        "report-reviewer-conclusion",
        "report-reviewer-improvement",
        "report-reviewer-logic",
    ]:
        assert sa in text, f"SKILL.md에 {sa} 참조 없음"
```

- [ ] **Step 3: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_skill_md.py -v
```

Expected: PASS

- [ ] **Step 4: commit**

```bash
git add skill_src/SKILL.md skill_src/tests/unit/test_skill_md.py
git commit -m "feat(skill): SKILL.md 워크플로 가이드 (Step 0~5 + 옵션 처리)"
```

---

### Task 7.2: Subagent 정의 — slide-analyzer (1단계)

**Files:**
- Create: `agents_src/report-reviewer-slide-analyzer.md`
- Create: `skill_src/tests/unit/test_agents_md.py`

- [ ] **Step 1: subagent 정의 작성**

`agents_src/report-reviewer-slide-analyzer.md`:
```markdown
---
name: report-reviewer-slide-analyzer
description: PPTX 슬라이드 1장의 텍스트·표·이미지를 multimodal 분석하여 핵심 내용·주장·데이터 포인트를 구조화. report-reviewer skill에서만 사용.
tools: Read
---

# 역할
너는 보고서 슬라이드 분석 전문가다. 입력으로 받은 슬라이드의 (1) 추출된 텍스트·표 데이터(JSON 파일) (2) 슬라이드 PNG 이미지 (3) 임베디드 이미지(있으면)를 multimodal로 함께 보고 구조화된 분석 결과를 출력한다.

# 입력
- 슬라이드 데이터 JSON 경로 (Read로 읽기)
- 슬라이드 이미지 경로 (Read로 읽기 — multimodal 입력)
- 임베디드 이미지 경로(있으면)

# 분석 항목
1. **슬라이드 핵심 메시지** (1~2문장 요약)
2. **핵심 주장(claims)** — 작성자가 이 슬라이드에서 주장하는 것 리스트
3. **핵심 데이터 포인트** — `(값, 단위, 맥락)` 튜플 리스트
4. **시각 정보 관찰** — 그림·차트·컨투어 등 이미지에서만 보이는 정보
5. **발표자 노트 요약** (있는 경우)
6. **다른 슬라이드와의 잠재적 연결점 힌트** — 다른 슬라이드에서 검증할 만한 것

# 원칙
- **원문 우선**: 추출된 텍스트(JSON)와 이미지 OCR 결과가 다르면 원문(JSON)을 신뢰. OCR은 보조.
- **그림 안 정보 보강**: JSON에 없는 그림 안 텍스트·차트 라벨·컨투어 hot spot은 이미지에서 추출
- **추측 최소화**: 보이지 않는 것은 만들지 말 것

# 출력 형식
JSON 코드 블록 1개로만 응답:
\```json
{
  "index": <슬라이드 번호>,
  "key_message": "...",
  "claims": ["...", "..."],
  "data_points": [
    {"value": "...", "unit": "...", "context": "..."}
  ],
  "vision_observations": ["...", "..."],
  "notes_summary": "...",
  "cross_slide_hints": ["...", "..."]
}
\```

JSON 외 추가 설명·markdown 헤더는 금지.
```

- [ ] **Step 2: agent 파일 검증 테스트**

`skill_src/tests/unit/test_agents_md.py`:
```python
from pathlib import Path

AGENTS_DIR = Path(__file__).parent.parent.parent.parent / "agents_src"

EXPECTED_AGENTS = [
    "report-reviewer-slide-analyzer",
]


def test_agents_have_frontmatter():
    for name in EXPECTED_AGENTS:
        path = AGENTS_DIR / f"{name}.md"
        assert path.exists(), f"누락: {path}"
        text = path.read_text(encoding="utf-8")
        assert text.startswith("---\n")
        end = text.find("\n---\n", 4)
        assert end > 0
        fm = text[4:end]
        assert f"name: {name}" in fm
        assert "description:" in fm
        assert "tools:" in fm


def test_slide_analyzer_specifics():
    text = (AGENTS_DIR / "report-reviewer-slide-analyzer.md").read_text(encoding="utf-8")
    # tools: Read만 허용
    assert "tools: Read" in text
    # 출력 스키마 키 명시
    for key in ["key_message", "claims", "data_points", "vision_observations"]:
        assert key in text
```

- [ ] **Step 3: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_agents_md.py -v
```

Expected: PASS

- [ ] **Step 4: commit**

```bash
git add agents_src/report-reviewer-slide-analyzer.md skill_src/tests/unit/test_agents_md.py
git commit -m "feat(agent): report-reviewer-slide-analyzer (1단계 multimodal 슬라이드 분석)"
```

---

### Task 7.3: Subagent 정의 — 6개 카테고리 (2단계)

**Files:**
- Create: `agents_src/report-reviewer-typo.md`
- Create: `agents_src/report-reviewer-terminology.md`
- Create: `agents_src/report-reviewer-data.md`
- Create: `agents_src/report-reviewer-conclusion.md`
- Create: `agents_src/report-reviewer-improvement.md`
- Create: `agents_src/report-reviewer-logic.md`
- Modify: `skill_src/tests/unit/test_agents_md.py`

각 SA는 동일 구조 (역할·입력·검증 항목·출력 형식)를 따르고, **카테고리·ID prefix·검증 항목만 다르다**.

- [ ] **Step 1: 공통 템플릿 + 6개 SA 파일 작성**

각 SA는 다음 구조를 갖는다:

```markdown
---
name: report-reviewer-<카테고리>
description: <카테고리 한국어 설명>. report-reviewer skill에서만 사용.
tools: Read
---

# 역할
너는 보고서의 <카테고리> 검토 전문가다. 1단계 슬라이드 분석 결과(`slide_summaries.json`)와 필요 시 원본 추출(`extracted.json`)을 보고 자기 카테고리에 해당하는 이슈를 검출한다.

# 입력
- `slide_summaries.json` 경로 (Read로 읽기)
- `extracted.json` 경로 (필요 시 Read로 인용·위치 정보 조회)

# 검증 항목 (카테고리별)
<항목 리스트>

# 원칙
- **원문 인용 필수**: `quoted_text`에 정확한 원문(짧게)
- **위치 명시**: shape_id로 extracted.json 도형 조회 → `position_pct` 복사
- **개선 제안 구체적**: "수정하세요"가 아니라 "X를 Y로 수정"
- **심각도 기준**: critical(반드시 수정)·warning(검토 권장)·info(참고)
- **ID prefix**: <카테고리별 prefix>를 사용 (예: T001, T002, ...)

# 출력 형식
JSON 코드 블록 1개:
\```json
[
  {
    "id": "<prefix>001",
    "category": "<카테고리>",
    "severity": "critical|warning|info",
    "slide_index": <int>,
    "shape_id": "s<i>_sh<j>",
    "position_hint": "<한국어>",
    "position_pct": {"left": <0~1>, "top": <0~1>, "width": <0~1>, "height": <0~1>},
    "quoted_text": "<원문 짧게>",
    "issue": "<문제 설명>",
    "suggestion": "<개선 제안 구체>",
    "evidence": "<근거>"
  }
]
\```

이슈 없으면 빈 배열 `[]`로 응답.
```

#### 6개 SA의 카테고리·prefix·검증 항목

| 파일 | category 값 | prefix | 검증 항목 |
|---|---|---|---|
| `report-reviewer-typo.md` | `typo` | `T` | 오타·맞춤법·표기 오류 (한글·영문). 원문 텍스트 기반. 인식 모호한 것은 info 처리 |
| `report-reviewer-terminology.md` | `terminology` | `TM` | 용어·약어 통일성. 한·영 혼용 일관성 (예: "스트레스 vs 응력"). 약어 첫 등장 시 풀어쓰기 여부 |
| `report-reviewer-data.md` | `data` | `D` | 표·본문 수치 정합성. 단위 누락. 인용 일치성. 자릿수 일관성 |
| `report-reviewer-conclusion.md` | `conclusion` | `C` | 결론이 데이터·자료로 뒷받침되는지. 근거 슬라이드 부재 검출 |
| `report-reviewer-improvement.md` | `improvement` | `I` | 정보 전달·결론 뒷받침 개선 제안. 누락된 가정·한계, 명확성 부족 |
| `report-reviewer-logic.md` | `logic` | `L` | overclaim 검출 (단정적 표현 vs 근거). 일반화 오류 (1샘플 → 전체). 가정·한계 명시 여부 |

각 파일을 위 구조 + 위 표의 카테고리/prefix/검증 항목으로 채워서 작성.

- [ ] **Step 2: 테스트 확장 — 6개 SA 파일 검증**

`skill_src/tests/unit/test_agents_md.py`의 `EXPECTED_AGENTS` 리스트 + 새 테스트:
```python
EXPECTED_AGENTS = [
    "report-reviewer-slide-analyzer",
    "report-reviewer-typo",
    "report-reviewer-terminology",
    "report-reviewer-data",
    "report-reviewer-conclusion",
    "report-reviewer-improvement",
    "report-reviewer-logic",
]


CATEGORY_PREFIX = {
    "typo": "T",
    "terminology": "TM",
    "data": "D",
    "conclusion": "C",
    "improvement": "I",
    "logic": "L",
}


def test_category_subagents_have_correct_prefix():
    for cat, prefix in CATEGORY_PREFIX.items():
        path = AGENTS_DIR / f"report-reviewer-{cat}.md"
        text = path.read_text(encoding="utf-8")
        # 출력 예시 ID에 정확한 prefix 등장 (예: T001, TM001 등)
        assert f"{prefix}001" in text, f"{cat} SA에 ID 예시 '{prefix}001' 누락"
        # category 값 명시
        assert f'"{cat}"' in text or f"`{cat}`" in text


def test_category_subagents_have_output_schema():
    for cat in CATEGORY_PREFIX:
        path = AGENTS_DIR / f"report-reviewer-{cat}.md"
        text = path.read_text(encoding="utf-8")
        # findings 스키마 핵심 키
        for key in ["id", "severity", "slide_index", "shape_id", "position_pct", "quoted_text", "issue", "suggestion"]:
            assert key in text, f"{cat} SA에 {key} 누락"
```

- [ ] **Step 3: 테스트 통과 확인**

```bash
.venv/Scripts/pytest tests/unit/test_agents_md.py -v
```

Expected: 모든 테스트 PASS

- [ ] **Step 4: commit**

```bash
git add agents_src/report-reviewer-typo.md agents_src/report-reviewer-terminology.md agents_src/report-reviewer-data.md agents_src/report-reviewer-conclusion.md agents_src/report-reviewer-improvement.md agents_src/report-reviewer-logic.md skill_src/tests/unit/test_agents_md.py
git commit -m "feat(agents): 6개 카테고리 subagent (typo/terminology/data/conclusion/improvement/logic)"
```

---

### Task 7.4: 글로벌 배포 (sync 스크립트 + 첫 배포)

**Files:**
- Create: `skill_src/deploy.ps1` (PowerShell — Windows 환경 가정)
- Create: `skill_src/README.md`

- [ ] **Step 1: deploy.ps1 작성**

`skill_src/deploy.ps1`:
```powershell
# report-reviewer 글로벌 배포 스크립트
# 사용법: PowerShell에서 실행
#   cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
#   ./deploy.ps1
#
# 동작:
#   - skill_src 전체를 ~/.claude/skills/report-reviewer/로 복사 (기존 덮어쓰기)
#   - agents_src/*.md를 ~/.claude/agents/로 복사
#   - .venv/는 제외 (배포 대상 환경에서 별도 생성)

$ErrorActionPreference = "Stop"

$SourceSkill = $PSScriptRoot
$SourceAgents = Join-Path (Split-Path -Parent $PSScriptRoot) "agents_src"
$DestSkill = Join-Path $env:USERPROFILE ".claude/skills/report-reviewer"
$DestAgents = Join-Path $env:USERPROFILE ".claude/agents"

Write-Host "=== report-reviewer 배포 ==="
Write-Host "Source skill : $SourceSkill"
Write-Host "Source agents: $SourceAgents"
Write-Host "Dest skill   : $DestSkill"
Write-Host "Dest agents  : $DestAgents"

# 1) Skill 본체 복사 (.venv, tests, __pycache__, *.pyc 제외)
if (Test-Path $DestSkill) {
    Write-Host "기존 $DestSkill 제거 중 (확인됨)..."
    Remove-Item -Recurse -Force $DestSkill
}
New-Item -ItemType Directory -Force -Path $DestSkill | Out-Null

robocopy $SourceSkill $DestSkill /E /XD .venv tests __pycache__ .pytest_cache /XF *.pyc deploy.ps1 | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy 실패: $LASTEXITCODE" }

# 2) Subagent 정의 복사
if (-not (Test-Path $DestAgents)) {
    New-Item -ItemType Directory -Force -Path $DestAgents | Out-Null
}
Copy-Item -Path "$SourceAgents/report-reviewer-*.md" -Destination $DestAgents -Force

Write-Host "`n=== 배포 완료 ==="
Write-Host "Skill files:"
Get-ChildItem -Recurse $DestSkill | Select-Object FullName | Format-Table -AutoSize
Write-Host "Agent files:"
Get-ChildItem "$DestAgents/report-reviewer-*.md" | Select-Object Name | Format-Table -AutoSize

Write-Host "`n다음 단계: 배포 위치에서 가상환경 생성 + 의존성 설치"
Write-Host "  cd `"$DestSkill`""
Write-Host "  python -m venv .venv"
Write-Host "  .venv\Scripts\pip install -e ."
```

- [ ] **Step 2: README.md 작성**

`skill_src/README.md`:
```markdown
# report-reviewer

사내 PPTX 보고서를 작성자가 자가점검하는 Claude Code Skill.

## 환경 요구사항

- Windows + MS PowerPoint 설치
- Python 3.10+
- Claude Code CLI
- (선택) Poppler for Windows — 임베디드 PDF 처리 시

## 설치

### 개발 환경

```bash
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
python -m venv .venv
.venv/Scripts/pip install -e ".[dev]"
```

### 글로벌 배포 (Skill로 등록)

PowerShell에서:
```powershell
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
./deploy.ps1
```

배포 후 글로벌 위치(`~/.claude/skills/report-reviewer/`)에서 가상환경 별도 생성:
```bash
cd "$env:USERPROFILE/.claude/skills/report-reviewer"
python -m venv .venv
.venv/Scripts/pip install -e .
```

## 사용

Claude Code에서:
```
/report-reviewer C:/path/to/report.pptx
```

옵션:
- `--resume`: 중간 산출물 보존되어 있으면 이어서
- `--rerun-stage2`: 1단계 재사용, 2단계만
- `--rerender`: findings.json 재사용, 리포트만

## 산출물

입력 PPTX와 같은 폴더에:
- `review_ws/` — 중간 산출물 (extracted.json, slide_summaries.json, findings.json, 슬라이드 이미지·썸네일)
- `review_output/review.md` — 마크다운 리포트
- `review_output/review.html` — HTML 리포트 (썸네일 + 위치 박스 오버레이)

## 테스트

```bash
cd skill_src
.venv/Scripts/pytest tests/unit/ -v          # 단위 테스트만 (CI 가능)
.venv/Scripts/pytest -m real_com -v          # 실제 PowerPoint 통합 테스트 (Windows + MS Office 필요)
```
```

- [ ] **Step 3: 첫 배포 실행 (수동, PowerShell에서)**

```powershell
cd "C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src"
./deploy.ps1
```

**배포 후 검증 체크리스트:**
- [ ] `~/.claude/skills/report-reviewer/SKILL.md` 존재
- [ ] `~/.claude/skills/report-reviewer/extract_and_render.py` 존재
- [ ] `~/.claude/skills/report-reviewer/render_report.py` 존재
- [ ] `~/.claude/skills/report-reviewer/src/`에 6개 .py 파일
- [ ] `~/.claude/skills/report-reviewer/templates/`에 .j2 + .css
- [ ] `~/.claude/agents/report-reviewer-*.md` 7개 파일
- [ ] `~/.claude/skills/report-reviewer/.venv/` 미존재 (배포 시 제외 확인)

- [ ] **Step 4: 글로벌 위치에 가상환경 + 의존성 설치**

```powershell
cd "$env:USERPROFILE/.claude/skills/report-reviewer"
python -m venv .venv
.venv/Scripts/pip install -e .
```

- [ ] **Step 5: Claude Code 재시작 → Skill 인식 확인**

Claude Code CLI에서:
```
/report-reviewer
```
- 명령어가 인식되는지 확인 (자동완성 등)
- subagent들이 `~/.claude/agents/`에서 자동 등록되는지 확인 (Claude Code의 subagent 목록에 `report-reviewer-*` 7개 노출)

- [ ] **Step 6: end-to-end 수동 테스트**

샘플 PPTX 1개로:
```
/report-reviewer C:/_PYTHON/0_CK_Project/8. Reviewer/skill_src/tests/fixtures/sample_text_only.pptx
```

**검증:**
- [ ] Step 1~5 워크플로 정상 진행 (Claude가 안내)
- [ ] `<pptx>/review_ws/` 생성, 모든 산출물 존재
- [ ] `<pptx>/review_output/review.md`, `review.html` 생성
- [ ] Critical 이슈 요약 메시지가 합리적

- [ ] **Step 7: commit**

```bash
git add skill_src/deploy.ps1 skill_src/README.md
git commit -m "feat(deploy): PowerShell 배포 스크립트 + README"
```

---

## Chunk 4 완료 시점

이 시점에:
- `extract_and_render.py`, `render_report.py` 두 CLI 완성
- `SKILL.md` 워크플로 가이드 (Step 0~5 + 옵션 처리)
- 7개 Subagent 정의 (`report-reviewer-slide-analyzer` + 6개 카테고리)
- `deploy.ps1` 배포 스크립트
- 글로벌 배포 완료 + Claude Code Skill로 인식
- end-to-end 수동 테스트 통과

**향후 확장 메모 (Chunk 5 또는 그 이후):**
- 골든 파일 정확도 메트릭 측정 (Precision/Recall) — 사용자가 샘플 PPTX 추가한 후
- CAE 도메인 특화 SA 추가 (단위계, 좌표계, 모델 메타정보, Pass/Fail) — spec §11
- PowerPoint 임베디드 별도 추출 (subprocess spawn)
- 사내 표준 양식 자동 점검표
- 이전 버전 보고서 자동 diff

---

# Plan 전체 완료

이 plan을 모두 실행하면:
- `~/.claude/skills/report-reviewer/`에 완성된 Skill 배포
- `~/.claude/agents/report-reviewer-*.md` 7개 SA 등록
- `/report-reviewer <pptx>` 명령으로 사내 PPTX 보고서 자가점검 가능

**Chunk 5 (선택)** — 사용자가 샘플 PPTX 1~2개를 추가한 시점에 골든 파일 정확도 메트릭 task로 확장. 별도 plan 또는 본 plan에 추가.
