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

배포 후 글로벌 위치에서 가상환경 별도 생성:
```powershell
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
- `review_ws/` — 중간 산출물 (extracted.json, slide_summaries.json, findings.json, 이미지)
- `review_output/review.md` — 마크다운 리포트
- `review_output/review.html` — HTML 리포트 (썸네일 + 위치 박스 오버레이)

## 테스트

```bash
cd skill_src
.venv/Scripts/pytest tests/unit/ -v          # 단위 테스트 (CI 가능)
.venv/Scripts/pytest -m real_com -v          # PowerPoint 통합 테스트 (수동)
```
