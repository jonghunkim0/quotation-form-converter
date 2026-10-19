# 견적 표준 양식 변환기 - Streamlit 서버 자동 시작 스크립트
# 이 스크립트는 항상 프로젝트 .venv의 Python으로 서버를 시작합니다.
# quotation_analyzer / plotly 등 의존성이 venv에만 설치되어 있기 때문에
# 시스템 Python으로 실행하면 "No module named 'quotation_analyzer'" 오류가 납니다.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venvPython = Join-Path $root ".venv\Scripts\python.exe"

# 1) venv가 없으면 생성
if (-not (Test-Path $venvPython)) {
    Write-Host "[start_server] .venv를 찾을 수 없어 생성합니다..." -ForegroundColor Yellow
    & python -m venv .venv
    if (-not (Test-Path $venvPython)) {
        Write-Host "[start_server] venv 생성에 실패했습니다." -ForegroundColor Red
        exit 1
    }
}

# 2) 의존성 자동 설치 (없을 때만)
Write-Host "[start_server] 의존성 확인 중..." -ForegroundColor Cyan
& $venvPython -m pip install --quiet --disable-pip-version-check -r requirements.txt
& $venvPython -m pip install --quiet --disable-pip-version-check -e "C:\AI_Agent\quotation-analyzer"
& $venvPython -m pip install --quiet --disable-pip-version-check plotly

# 3) 이미 실행 중인 서버(8501)가 있으면 종료
$existing = Get-NetTCPConnection -LocalPort 8501 -State Listen -ErrorAction SilentlyContinue
if ($existing) {
    $procIds = $existing | Select-Object -ExpandProperty OwningProcess -Unique
    foreach ($procId in $procIds) {
        Write-Host "[start_server] 기존 서버(PID $procId) 종료 중..." -ForegroundColor Yellow
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
}

# 4) venv Python으로 서버 시작 (로그를 server.log에 캡처)
#    cmd /c 리다이렉트로 stdout+stderr를 함께 캡처 (PowerShell 2>&1의
#    NativeCommandError 충돌 회피). 서버는 이 프로세스와 함께 유지됨.
Write-Host "[start_server] venv Python으로 Streamlit 서버 시작: $venvPython" -ForegroundColor Green
$logFile = Join-Path $root "server.log"
cmd /c "`"$venvPython`" -m streamlit run app.py --server.port 8501 > `"$logFile`" 2>&1"
