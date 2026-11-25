# ═══════════════════════════════════════════════════════════════
#  견적 표준 양식 변환기 — 서버 PC 배포 스크립트
#
#  [사용법]
#  1. 이 프로젝트 폴더(quotation-form-converter)를 서버 PC에 복사
#  2. PowerShell에서 이 스크립트 실행:
#       powershell -ExecutionPolicy Bypass -File .\deploy_server.ps1
#  3. 화면에 표시되는 URL(http://서버IP:8501)을 사용자에게 전달
#
#  [자동 처리 내용]
#  - Python 확인
#  - quotation-analyzer 패키지 클론/설치 (없을 때만)
#  - .venv 생성 + 의존성 설치
#  - 자동 저장 폴더(D:\견적분석 Agent 파일) 생성
#  - Windows 방화벽 8501 포트 허용 (관리자 권한 필요)
#  - 서버 시작 (LAN 접속 가능: 0.0.0.0)
# ═══════════════════════════════════════════════════════════════

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$port = 8501
$analyzerDir = "C:\AI_Agent\quotation-analyzer"
$analyzerRepo = "https://github.com/jonghunkim0/quotation-analyzer.git"
$saveDir = "D:\견적분석 Agent 파일"

Write-Host ""
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  견적 표준 양식 변환기 — 서버 배포" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ── 1. Python 확인 ─────────────────────────────────────────────
Write-Host "[1/6] Python 확인 중..." -ForegroundColor Yellow
$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) {
    Write-Host "  ❌ Python이 설치되어 있지 않습니다." -ForegroundColor Red
    Write-Host "     https://www.python.org/downloads/ 에서 설치 후 다시 실행하세요." -ForegroundColor Red
    Write-Host "     (설치 시 'Add python.exe to PATH' 체크 필수)" -ForegroundColor Red
    exit 1
}
$pyVer = & python --version 2>&1
Write-Host "  $pyVer" -ForegroundColor Green

# ── 2. quotation-analyzer 패키지 확인/클론 ─────────────────────
Write-Host "[2/6] quotation-analyzer 패키지 확인 중..." -ForegroundColor Yellow
if (-not (Test-Path $analyzerDir)) {
    Write-Host "  $analyzerDir 이 없어 클론합니다..." -ForegroundColor Yellow
    git clone $analyzerRepo $analyzerDir
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ❌ 클론 실패. git이 설치되어 있는지 확인하세요." -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "  기존 폴더 사용: $analyzerDir" -ForegroundColor Green
}

# ── 3. venv + 의존성 설치 ──────────────────────────────────────
Write-Host "[3/6] Python 환경(.venv) + 의존성 설치 중..." -ForegroundColor Yellow
$venvPython = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    Write-Host "  .venv 생성 중..." -ForegroundColor Yellow
    & python -m venv .venv
}
Write-Host "  의존성 설치 중 (최초 실행 시 수 분 소요)..." -ForegroundColor Yellow
& $venvPython -m pip install --quiet --disable-pip-version-check -r requirements.txt
& $venvPython -m pip install --quiet --disable-pip-version-check -e $analyzerDir
& $venvPython -m pip install --quiet --disable-pip-version-check plotly
Write-Host "  의존성 설치 완료" -ForegroundColor Green

# ── 4. 자동 저장 폴더 생성 ─────────────────────────────────────
Write-Host "[4/6] 자동 저장 폴더 확인 중..." -ForegroundColor Yellow
if (-not (Test-Path $saveDir)) {
    New-Item -ItemType Directory -Path $saveDir -Force | Out-Null
    Write-Host "  생성: $saveDir" -ForegroundColor Green
} else {
    Write-Host "  기존 폴더 사용: $saveDir" -ForegroundColor Green
}

# ── 5. 방화벽 포트 허용 (관리자 권한 필요) ─────────────────────
Write-Host "[5/6] Windows 방화벽 $port 포트 허용..." -ForegroundColor Yellow
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    $rule = "QuotationConverter $port"
    if (-not (Get-NetFirewallRule -DisplayName $rule -ErrorAction SilentlyContinue)) {
        New-NetFirewallRule -DisplayName $rule -Direction Inbound -Protocol TCP `
            -LocalPort $port -Action Allow -Profile Any | Out-Null
        Write-Host "  방화벽 규칙 추가: $rule" -ForegroundColor Green
    } else {
        Write-Host "  방화벽 규칙 이미 존재" -ForegroundColor Green
    }
} else {
    Write-Host "  ⚠️  관리자 권한이 없어 방화벽 규칙을 추가하지 못했습니다." -ForegroundColor Yellow
    Write-Host "     다른 PC에서 접속이 안 되면, 관리자 PowerShell에서 아래 명령을 실행하세요:" -ForegroundColor Yellow
    Write-Host "     New-NetFirewallRule -DisplayName 'QuotationConverter $port' -Direction Inbound -Protocol TCP -LocalPort $port -Action Allow -Profile Any" -ForegroundColor Yellow
}

# ── 6. 서버 시작 (LAN 접속 가능) ───────────────────────────────
Write-Host "[6/6] 서버 시작 중 (LAN 접속 가능)..." -ForegroundColor Yellow

# 기존 서버가 있으면 종료
$existing = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($existing) {
    $procIds = $existing | Select-Object -ExpandProperty OwningProcess -Unique
    foreach ($procId in $procIds) {
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
}

# Start-Process 로 독립 프로세스 생성 → 스크립트 종료 후에도 서버 유지
Start-Process -FilePath $venvPython `
    -ArgumentList @("-m", "streamlit", "run", "app.py", "--server.port", "$port", "--server.address", "0.0.0.0") `
    -WorkingDirectory $root -WindowStyle Hidden
Start-Sleep -Seconds 8

# 서버 응답 확인
$ok = $false
try {
    $r = Invoke-WebRequest -Uri "http://localhost:$port" -UseBasicParsing -TimeoutSec 10
    if ($r.StatusCode -eq 200) { $ok = $true }
} catch { }

# LAN IP 조회
$lanIp = (Get-NetIPAddress -AddressFamily IPv4 |
    Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.*' } |
    Select-Object -First 1).IPAddress

Write-Host ""
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
if ($ok) {
    Write-Host "  ✅ 서버 실행 성공!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  사용자에게 전달할 주소:" -ForegroundColor Cyan
    Write-Host "  ┌─────────────────────────────────────┐" -ForegroundColor Cyan
    Write-Host "  │  http://${lanIp}:$port              │" -ForegroundColor White
    Write-Host "  └─────────────────────────────────────┘" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  사용자는 브라우저에서 위 주소로 접속하면 됩니다." -ForegroundColor Green
    Write-Host "  (사용자 PC에는 아무것도 설치할 필요 없음)" -ForegroundColor Green
} else {
    Write-Host "  ⚠️  서버 응답 확인에 실패했습니다. server.log 를 확인하세요." -ForegroundColor Yellow
}
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  [참고] 서버를 다시 시작하려면: 이 스크립트를 다시 실행하면 됩니다." -ForegroundColor Gray
Write-Host "         (기존 서버는 자동 종료 후 재시작)" -ForegroundColor Gray
Write-Host ""
