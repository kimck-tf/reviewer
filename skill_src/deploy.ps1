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
    Write-Host "기존 $DestSkill 제거 중..."
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
