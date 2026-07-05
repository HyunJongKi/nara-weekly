# Windows Task Scheduler 가 매일 09:00 KST 에 실행.
# 카카오 refresh_token 을 미리 리프레시해서, 카카오가 새 토큰을 함께 주면
# 로컬 저장 + GitHub Secret 을 자동 갱신 → 다음 워크플로 실행 때 새 토큰으로 안전 사용.
# 이번 7/3 KOE322(expired_or_invalid_refresh_token) 사고의 근본 재발 방지책.

$ErrorActionPreference = 'Stop'
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root       = Split-Path -Parent $ScriptDir
$TokenFile  = Join-Path $Root ".tokens.json"
$LogFile    = Join-Path $Root "rotate.log"
$GhExe      = "C:\Program Files\GitHub CLI\gh.exe"
$Repo       = "HyunJongKi/nara-weekly"
$RestKey    = "e03e0bc0a08b39683ca125d4aac7bcd6"
$ClientSec  = "QgOkaL8zgDa7mAcnSlMh3YbYuMgp1F5y"
$Stamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"

function Log($msg) { Add-Content -Path $LogFile -Value "[$Stamp] $msg" -Encoding UTF8 }

try {
    if (-not (Test-Path $TokenFile)) {
        Log "ERR .tokens.json 없음 - 최초 refresh_token 을 이 파일에 넣어야 함"
        exit 1
    }
    $current = (Get-Content $TokenFile -Raw | ConvertFrom-Json).refresh_token
    if (-not $current) { Log "ERR .tokens.json 에 refresh_token 필드 없음"; exit 1 }

    $body = @{
        grant_type    = "refresh_token"
        client_id     = $RestKey
        refresh_token = $current
        client_secret = $ClientSec
    }
    $resp = Invoke-RestMethod -Uri "https://kauth.kakao.com/oauth/token" `
        -Method POST -Body $body -TimeoutSec 15

    # 카카오는 만료 임박(1개월 이내) 시에만 새 refresh_token 을 함께 반환한다.
    if ($resp.refresh_token -and $resp.refresh_token -ne $current) {
        # 로컬 파일 갱신
        @{ refresh_token = $resp.refresh_token; rotated_at = $Stamp } |
            ConvertTo-Json | Set-Content -Path $TokenFile -Encoding UTF8

        # GitHub Secret 갱신 (이 PC 에 이미 인증된 gh CLI 사용, PAT 불필요)
        $preview = $resp.refresh_token.Substring(0, 12)
        Log "ROTATE 새 refresh_token 감지 → GitHub Secret 갱신 중 ($preview...)"
        $result = & $GhExe secret set KAKAO_REFRESH_TOKEN --repo $Repo --body $resp.refresh_token 2>&1
        if ($LASTEXITCODE -eq 0) {
            Log "OK GitHub Secret KAKAO_REFRESH_TOKEN 갱신 완료"
        } else {
            Log "ERR gh secret set 실패 ($LASTEXITCODE): $result"
            exit 1
        }
    } else {
        Log "OK 갱신 불필요 (카카오가 새 토큰 미반환 = 아직 여유 있음)"
    }
    exit 0
} catch {
    Log "ERR exception: $_"
    exit 1
}
