# Windows Task Scheduler 가 매주 수·금 10:00 KST 에 호출.
# 이 PC 에 설치된 GitHub CLI (이미 인증됨)로 weekly-nara workflow 강제 trigger.
# GitHub Actions cron 누락에 대한 사용자 PC 측 안전망.

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile   = Join-Path (Split-Path -Parent $ScriptDir) "trigger.log"
$GhExe     = "C:\Program Files\GitHub CLI\gh.exe"
$Repo      = "HyunJongKi/nara-weekly"
$Workflow  = "weekly-nara"
$Stamp     = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"

function Log($msg) { Add-Content -Path $LogFile -Value "[$Stamp] $msg" -Encoding UTF8 }

try {
    if (-not (Test-Path $GhExe)) {
        Log "ERR gh.exe 없음: $GhExe"
        exit 1
    }
    $output = & $GhExe workflow run $Workflow --repo $Repo --ref main 2>&1
    if ($LASTEXITCODE -eq 0) {
        Log "OK $Workflow trigger 성공 — $output"
        exit 0
    } else {
        Log "ERR exit=$LASTEXITCODE — $output"
        exit 1
    }
} catch {
    Log "ERR exception: $_"
    exit 1
}
