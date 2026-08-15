# 클로드 데스크톱 앱에 gpcm-kr 서버를 등록한다.
#
# run_mcp_setup.bat 이 부른다. 직접 실행할 일은 없다.
#
# 설정 파일에 다른 서버가 이미 있을 수 있으므로 통째로 덮어쓰지 않고 gpcm-kr 항목만
# 넣거나 갈아끼운다. 배치 파일로 JSON 을 다루면 기존 내용을 보존할 수 없어서
# 이 부분만 PowerShell 로 뺐다.

param(
    [Parameter(Mandatory = $true)][string]$PythonPath,
    [string]$ApiKey = ""
)

$ErrorActionPreference = "Stop"

$dir = Join-Path $env:APPDATA "Claude"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$path = Join-Path $dir "claude_desktop_config.json"

# 기존 설정을 읽는다. 파일이 깨져 있으면 빈 것으로 보고 새로 만든다
# (백업을 남기므로 원본이 사라지지는 않는다).
$cfg = $null
if (Test-Path $path) {
    Copy-Item $path "$path.bak" -Force
    try {
        $cfg = Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Warning "기존 설정을 읽지 못해 새로 만듭니다. 원본은 $path.bak 에 있습니다."
        $cfg = $null
    }
}
if ($null -eq $cfg) { $cfg = New-Object PSObject }

if (-not $cfg.PSObject.Properties['mcpServers']) {
    $cfg | Add-Member -NotePropertyName 'mcpServers' -NotePropertyValue (New-Object PSObject) -Force
}

$entry = [ordered]@{
    command = $PythonPath
    args    = @("-m", "gpcm_mcp.server")
}

# 키는 원래 환경변수로만 받게 만들었지만, 앱이 켜져 있던 동안 저장한 환경변수는
# 앱에 전달되지 않는다. 그 실패를 없애려고 여기에도 넣는다. 이 파일은 %APPDATA%
# 안의 개인 설정이라 저장소에 들어가지 않는다.
if ($ApiKey -ne "") {
    $entry['env'] = [ordered]@{ OPENDART_API_KEY = $ApiKey }
}

$cfg.mcpServers | Add-Member -NotePropertyName 'gpcm-kr' `
    -NotePropertyValue ([PSCustomObject]$entry) -Force

# Set-Content -Encoding UTF8 은 윈도우 파워셸 5.1 에서 BOM 을 붙인다. JSON 파서가
# BOM 을 만나면 읽지 못하는 경우가 있어서 BOM 없이 직접 쓴다.
$json = $cfg | ConvertTo-Json -Depth 20
[System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))

Write-Output $path
