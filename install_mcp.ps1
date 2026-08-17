# GPCM을 Claude 데스크톱에 로컬 MCP로 등록하는 스크립트 (Windows)
#
#   1) 이 파일이 있는 폴더에서 주소창에 powershell 입력 → 엔터
#   2) 아래 한 줄 실행
#        powershell -ExecutionPolicy Bypass -File .\install_mcp.ps1
#
# 인증키를 물어본다. 이미 mydart를 설치했다면 그 키를 자동으로 재사용한다.
# 미리 주려면:
#        powershell -ExecutionPolicy Bypass -File .\install_mcp.ps1 -ApiKey "발급받은키"

param([string]$ApiKey = "")

$ErrorActionPreference = "Stop"
$root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

function Step($text) { Write-Host "`n== $text" -ForegroundColor Cyan }
function Ok($text)   { Write-Host "   OK  $text" -ForegroundColor Green }
function Warn($text) { Write-Host "   !!  $text" -ForegroundColor Yellow }

Write-Host "gpcm-mcp 등록 (GPCM을 Claude 데스크톱에서 실행)" -ForegroundColor White
Write-Host "폴더: $root"

# --- 0. 파일 확인 --------------------------------------------------------------
Step "파일 확인"
foreach ($f in @("gpcm_mcp.py", "gpcm_kr.py", "requirements-mcp.txt")) {
    if (-not (Test-Path (Join-Path $root $f))) {
        throw "$f 가 없습니다. 압축을 푼 Finance 폴더에서 실행하세요."
    }
}
Ok "gpcm_mcp.py / gpcm_kr.py / requirements-mcp.txt"

# --- 1. Claude 버전 확인 --------------------------------------------------------
# Store 버전은 설정을 격리된 폴더에서 읽어 이 등록을 못 본다. 그러나 정식 버전의
# 설치 위치는 PC마다 달라 자동 감지가 빗나갈 수 있다. 실제로 정식 버전 사용자를
# 잘못 막은 적이 있으므로, 여기서는 어떤 경우에도 멈추지 않는다 — 상황만 알려주고
# 계속 진행한다. (표준 위치에 설정을 쓰는 것 자체는 해가 없다.)
Step "Claude 버전 확인"
$officialPaths = @(
    (Join-Path $env:LOCALAPPDATA "AnthropicClaude"),
    (Join-Path $env:LOCALAPPDATA "Programs\Claude"),
    (Join-Path $env:LOCALAPPDATA "Programs\claude-desktop"),
    (Join-Path $env:APPDATA "Claude")
)
$official = $false
foreach ($p in $officialPaths) { if (Test-Path $p) { $official = $true; break } }
$store = $false
if (Get-Command Get-AppxPackage -ErrorAction SilentlyContinue) {
    $store = [bool](Get-AppxPackage *Claude* -ErrorAction SilentlyContinue)
}
if ($official) {
    if ($store) {
        Warn "Microsoft Store 버전 흔적도 있습니다. Claude는 정식 버전으로 실행하세요."
        Warn "(잔재 제거는 선택:  Get-AppxPackage *Claude* | Remove-AppxPackage)"
    }
    Ok "정식 버전 확인"
} else {
    Warn "정식 버전 설치 흔적을 찾지 못했습니다 — 감지가 빗나갔을 수 있어 그대로 진행합니다."
    if ($store) {
        Warn "만약 Store 버전만 쓰고 계시다면 이 등록을 Claude가 읽지 못합니다."
        Warn "그 경우 https://claude.ai/download 에서 정식 버전을 설치하세요."
    }
}

# --- 2. 인증키 -----------------------------------------------------------------
# mydart나 기존 gpcm 설정에 이미 있으면 그대로 쓴다. 재입력하지 않아도 된다.
Step "OpenDART 인증키"
$configPath = Join-Path $env:APPDATA "Claude\claude_desktop_config.json"
if (-not $ApiKey -and (Test-Path $configPath)) {
    try {
        $existing = Get-Content $configPath -Raw | ConvertFrom-Json
        foreach ($name in @("gpcm", "mydart")) {
            $candidate = $existing.mcpServers.$name.env.DART_API_KEY
            if ($candidate) { $ApiKey = $candidate; Ok "기존 $name 설정에서 인증키를 찾았습니다."; break }
        }
    } catch { $ApiKey = "" }
}
if (-not $ApiKey) {
    Write-Host "   https://opendart.fss.or.kr 에서 무료로 발급받을 수 있습니다 (40자리)."
    $ApiKey = (Read-Host "   인증키를 붙여넣고 엔터").Trim()
}
if (-not $ApiKey) { throw "인증키가 없으면 등록해도 조회가 되지 않습니다." }
Ok "$($ApiKey.Length)자 확인 (화면에 키는 표시하지 않습니다)"

# --- 3. uv 준비 ----------------------------------------------------------------
Step "uv 준비"
$uv = Join-Path $env:USERPROFILE ".local\bin\uv.exe"
if (-not (Test-Path $uv)) {
    $found = (Get-Command uv -ErrorAction SilentlyContinue).Source
    if ($found) { $uv = $found }
}
if (-not (Test-Path $uv)) {
    Write-Host "   uv가 없어 설치합니다 (1분 내외)..."
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex" | Out-Null
    $uv = Join-Path $env:USERPROFILE ".local\bin\uv.exe"
}
if (-not (Test-Path $uv)) { throw "uv 설치에 실패했습니다. https://astral.sh/uv 를 참고하세요." }
Ok (& $uv --version)

# --- 4. 실행 중인 Claude 종료 ---------------------------------------------------
Step "실행 중인 Claude 종료"
Get-Process Claude -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500
Ok "종료 완료"

# --- 5. Claude 설정에 등록 ------------------------------------------------------
# run_kr.bat과 같은 --no-project 방식이라 설치·빌드가 없다 (백신 문제 회피 경로).
# 기존 서버(mydart·myacc 등)는 건드리지 않고 gpcm만 추가한다.
Step "Claude 설정 등록"
if (Test-Path $configPath) {
    Copy-Item $configPath "$configPath.bak" -Force
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
} else {
    New-Item -ItemType Directory -Force -Path (Split-Path $configPath) | Out-Null
    $config = [pscustomobject]@{}
}
if (-not $config.mcpServers) {
    $config | Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{}) -Force
}
$server = [ordered]@{
    command = $uv
    args    = @("run", "--no-project", "--python", "3.12",
                "--with-requirements", (Join-Path $root "requirements-mcp.txt"),
                "python", (Join-Path $root "gpcm_mcp.py"))
    env     = [ordered]@{ DART_API_KEY = $ApiKey }
}
$config.mcpServers | Add-Member -NotePropertyName gpcm -NotePropertyValue $server -Force
[System.IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json -Depth 30))
Ok $configPath
Ok ("등록된 서버: " + (($config.mcpServers.PSObject.Properties.Name) -join ", "))

# --- 6. 자체점검 ---------------------------------------------------------------
# Claude에 붙이기 전에 여기서 걸러야 원인이 한 곳으로 좁혀진다.
# 첫 실행은 파이썬·패키지를 내려받아 몇 분 걸릴 수 있다.
Step "연결 자체점검 (첫 실행은 몇 분 걸릴 수 있습니다)"
$env:DART_API_KEY = $ApiKey
& $uv run --no-project --python 3.12 `
    --with-requirements (Join-Path $root "requirements-mcp.txt") `
    python (Join-Path $root "gpcm_mcp.py") --selftest
if ($LASTEXITCODE -ne 0) {
    Warn "자체점검이 실패했습니다. 위 메시지를 확인하세요. (등록 자체는 완료됨)"
} else {
    Ok "자체점검 통과"
}

Write-Host "`n끝났습니다." -ForegroundColor Green
Write-Host "Claude Desktop을 실행하고 설정 → 개발자에서 gpcm 을 확인하세요."
Write-Host "대화에서 이렇게 시험해보세요:  run_gpcm으로 005930, 2025.4Q 돌려줘"
Write-Host "결과 엑셀은 문서\GPCM 폴더에 저장됩니다."
