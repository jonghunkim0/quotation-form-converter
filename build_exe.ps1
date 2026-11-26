# ═══════════════════════════════════════════════════════════════
#  견적 표준 양식 변환기 — exe 빌드 스크립트 (PyInstaller)
#
#  [사용법]
#    powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
#
#  [결과]
#    dist\견적표준양식변환기_v{버전}.exe
#    (버전은 app.py의 APP_VERSION에서 자동 읽음)
#
#  [배포]
#    exe 파일 하나만 서버 PC에 복사 → 실행 → 접속 주소 표시
#    (Python 설치 불필요)
# ═══════════════════════════════════════════════════════════════

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$analyzerDir = "C:\AI_Agent\quotation-analyzer"
$venvPython = Join-Path $root ".venv\Scripts\python.exe"

Write-Host ""
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  견적 표준 양식 변환기 — exe 빌드" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ── 1. 버전 확인 (app.py의 APP_VERSION) ────────────────────────
$verLine = Select-String -Path (Join-Path $root "app.py") -Pattern '^APP_VERSION\s*=\s*"([^"]+)"' | Select-Object -First 1
if (-not $verLine) {
    Write-Host "  ❌ app.py에서 APP_VERSION을 찾을 수 없습니다." -ForegroundColor Red
    exit 1
}
$version = $verLine.Matches[0].Groups[1].Value
$exeName = "견적표준양식변환기_v$version"
Write-Host "[1/5] 버전: v$version" -ForegroundColor Yellow

# ── 2. PyInstaller 확인/설치 ───────────────────────────────────
Write-Host "[2/5] PyInstaller 확인 중..." -ForegroundColor Yellow
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& $venvPython -c "import PyInstaller" 2>&1 | Out-Null
$pyiOk = ($LASTEXITCODE -eq 0)
$ErrorActionPreference = $prevEAP
if (-not $pyiOk) {
    Write-Host "  PyInstaller 설치 중..." -ForegroundColor Yellow
    & $venvPython -m pip install --quiet --disable-pip-version-check pyinstaller
}
Write-Host "  PyInstaller 준비 완료" -ForegroundColor Green

# ── 3. data 폴더 확인 (사용자 파일 보존, 그대로 번들) ──────────
Write-Host "[3/5] data 폴더 확인 중..." -ForegroundColor Yellow
$dataDir = Join-Path $root "data"
if (-not (Test-Path $dataDir)) {
    New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
}
Write-Host "  data/ 폴더 유지 (사용자 파일 보존)" -ForegroundColor Green

# ── 4. PyInstaller 빌드 ────────────────────────────────────────
Write-Host "[4/5] PyInstaller 빌드 중 (수 분 소요)..." -ForegroundColor Yellow

$specFile = Join-Path $root "build\${exeName}.spec"
$distDir = Join-Path $root "dist"

# --add-data: 런처가 런타임에 복사할 앱 소스/리소스 (Windows: src;dst)
# 디렉터리는 dst에 디렉터리명을 명시해야 _MEIPASS/<이름> 으로 번들됨
$addDataArgs = @(
    "--add-data", "$root\app.py;."
    "--add-data", "$root\config.py;."
    "--add-data", "$root\run_conversion.py;."
    "--add-data", "$root\parsers;parsers"
    "--add-data", "$root\converters;converters"
    "--add-data", "$root\classification;classification"
    "--add-data", "$root\templates;templates"
    "--add-data", "$root\data;data"
)
# templates xlsx 파일 명시적 추가 (한글 파일명 디렉터리 복사 누락 방지)
Get-ChildItem (Join-Path $root "templates") -Filter "*.xlsx" -File | ForEach-Object {
    $addDataArgs += @("--add-data", "$($_.FullName);templates")
}

$prevEAP = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --console `
    --name $exeName `
    --paths $analyzerDir `
    --collect-all quotation_analyzer `
    --collect-all streamlit `
    --copy-metadata altair `
    --copy-metadata plotly `
    --copy-metadata pandas `
    --copy-metadata numpy `
    --copy-metadata openpyxl `
    --copy-metadata pdfplumber `
    @addDataArgs `
    --hidden-import parsers `
    --hidden-import parsers.excel_parser `
    --hidden-import parsers.pdf_parser `
    --hidden-import converters `
    --hidden-import converters.rule_mapper `
    --hidden-import converters.standard_exporter `
    --hidden-import converters.ratio_adjuster `
    --hidden-import converters.standard_price `
    --hidden-import classification `
    --hidden-import classification.auto_classifier `
    --hidden-import streamlit.web.cli `
    --hidden-import streamlit.web.bootstrap `
    --hidden-import plotly `
    --hidden-import pdfplumber `
    --exclude-module tkinter `
    --exclude-module matplotlib `
    --exclude-module numpy.testing `
    --distpath $distDir `
    --workpath (Join-Path $root "build") `
    --specpath (Join-Path $root "build") `
    (Join-Path $root "launcher.py")
$buildOk = ($LASTEXITCODE -eq 0)
$ErrorActionPreference = $prevEAP

if (-not $buildOk) {
    Write-Host "  ❌ 빌드 실패" -ForegroundColor Red
    exit 1
}

# ── 5. 결과 확인 ───────────────────────────────────────────────
$exePath = Join-Path $distDir "${exeName}.exe"
Write-Host "[5/5] 빌드 결과 확인 중..." -ForegroundColor Yellow
if (Test-Path $exePath) {
    $sizeMB = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
    Write-Host ""
    Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  ✅ 빌드 성공!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  파일: $exePath" -ForegroundColor White
    Write-Host "  크기: ${sizeMB}MB" -ForegroundColor White
    Write-Host ""
    Write-Host "  [배포 방법]" -ForegroundColor Cyan
    Write-Host "  1. 위 exe 파일을 서버 PC에 복사" -ForegroundColor White
    Write-Host "  2. 더블클릭 실행 (Python 설치 불필요)" -ForegroundColor White
    Write-Host "  3. 콘솔에 표시되는 접속 주소를 사용자에게 전달" -ForegroundColor White
    Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
} else {
    Write-Host "  ❌ exe 파일을 찾을 수 없습니다: $exePath" -ForegroundColor Red
    exit 1
}
Write-Host ""
