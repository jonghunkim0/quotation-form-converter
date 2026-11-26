# ═══════════════════════════════════════════════════════════════
#  견적 표준 양식 변환기 — 워치독 작업 스케줄러 등록
#
#  [사용법] (서버 PC에서 1회만 실행)
#    powershell -ExecutionPolicy Bypass -File .\setup_watchdog.ps1
#
#  [효과]
#    - 로그인 시 워치독 자동 시작 (서버도 자동 시작됨)
#    - 5분마다 워치독 생존 확인 (워치독이 죽어도 자동 복구)
#    - 워치독이 30초마다 서버 health 체크 → 죽으면 exe 자동 재시작
# ═══════════════════════════════════════════════════════════════

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = "견적변환기_Watchdog"

if (-not (Test-Path (Join-Path $root "watchdog.ps1"))) {
    Write-Host "❌ watchdog.ps1을 찾을 수 없습니다. 같은 폴더에 있어야 합니다." -ForegroundColor Red
    exit 1
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$root\watchdog.ps1`""

# 트리거 1: 로그인 시
$triggerLogon = New-ScheduledTaskTrigger -AtLogOn

# 트리거 2: 지금부터 5분마다 반복 (워치독 생존 확인용)
$triggerRepeat = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

Register-ScheduledTask -TaskName $taskName -Action $action `
    -Trigger @($triggerLogon, $triggerRepeat) -Force | Out-Null

Write-Host ""
Write-Host "✅ 워치독 작업 스케줄러 등록 완료: $taskName" -ForegroundColor Green
Write-Host ""
Write-Host "  - 로그인 시 자동 시작 (서버도 함께 자동 시작)"
Write-Host "  - 5분마다 워치독 생존 확인"
Write-Host "  - 워치독이 30초마다 서버 상태 확인 → 죽으면 자동 재시작"
Write-Host ""
Write-Host "  [확인] Get-ScheduledTask -TaskName '$taskName'"
Write-Host "  [제거] Unregister-ScheduledTask -TaskName '$taskName'"
Write-Host ""
