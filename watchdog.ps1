# ═══════════════════════════════════════════════════════════════
#  견적 표준 양식 변환기 — 워치독 (자동 재시작)
#
#  [동작]
#    - 30초마다 서버 health 체크 (http://localhost:8501/_stcore/health)
#    - 3회 연속 실패 시 기존 프로세스 종료 후 exe 자동 재시작
#    - 워치독이 이미 실행 중이면 즉시 종료 (단일 인스턴스 보장)
#      → 작업 스케줄러가 5분마다 반복 실행해도 중복 실행되지 않음
#
#  [설치] (서버 PC에서 1회만 실행)
#    powershell -ExecutionPolicy Bypass -File .\setup_watchdog.ps1
#
#  [파일 위치]
#    exe와 같은 폴더에 watchdog.ps1을 놓으면 됨
# ═══════════════════════════════════════════════════════════════

$ErrorActionPreference = "SilentlyContinue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# exe 찾기 (같은 폴더 → dist/ 순서로, 버전 무관)
$exe = Get-ChildItem (Join-Path $root "견적표준양식변환기_v*.exe") -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $exe) {
    $exe = Get-ChildItem (Join-Path $root "dist\견적표준양식변환기_v*.exe") -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $exe) { exit 1 }

# ── 단일 인스턴스 확인 ─────────────────────────────────────────
$pidFile = Join-Path $root "watchdog.pid"
if (Test-Path $pidFile) {
    $oldPid = (Get-Content $pidFile -Raw).Trim()
    if ($oldPid -match '^\d+$') {
        $p = Get-Process -Id ([int]$oldPid) -ErrorAction SilentlyContinue
        if ($p) { exit 0 }  # 이미 실행 중 → 종료
    }
}
$PID | Out-File -FilePath $pidFile -Encoding ASCII

$logFile = Join-Path $root "watchdog.log"
function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Write-Log "=== 워치독 시작 (PID $PID) ==="

$failCount = 0
while ($true) {
    $ok = $false
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:8501/_stcore/health" -UseBasicParsing -TimeoutSec 5
        if ($r.StatusCode -eq 200) { $ok = $true }
    } catch { }

    if ($ok) {
        if ($failCount -gt 0) { Write-Log "서버 복구됨" }
        $failCount = 0
    } else {
        $failCount++
        Write-Log "health 체크 실패 ($failCount/3)"
        if ($failCount -ge 3) {
            Write-Log "서버 재시작 중..."
            Get-Process | Where-Object { $_.ProcessName -like "견적표준양식변환기*" } | Stop-Process -Force
            Start-Sleep -Seconds 2
            Start-Process -FilePath $exe.FullName -WorkingDirectory $root -WindowStyle Hidden
            $failCount = 0
            Start-Sleep -Seconds 15
        }
    }
    Start-Sleep -Seconds 30
}
