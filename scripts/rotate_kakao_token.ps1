# 매일 05:00 / 12:00 / 22:00 KST Task Scheduler 실행 (하루 3회).
# 카카오 refresh API 를 자주 호출해, 카카오가 만료 임박 시 응답에 새 토큰을 담아주면
# 그 즉시 잡아서 로컬(.tokens.json) + GitHub Secret 을 sync.
#
# 하루 3회로 자주 돌려 workflow(수/금 10시)가 카카오 응답에서 먼저 새 토큰을
# 받아가버리기 전에 로컬이 잡을 확률을 높인다. (완벽 보장은 PAT 필요 - README 참조)
#
# 이 PC 의 인증된 gh CLI 사용 → 별도 PAT 불필요.

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root      = Split-Path -Parent $ScriptDir
$TokenFile = Join-Path $Root ".tokens.json"
$LogFile   = Join-Path $Root "rotate.log"
$GhExe     = "C:\Program Files\GitHub CLI\gh.exe"
$Repo      = "HyunJongKi/nara-weekly"
$RestKey   = "e03e0bc0a08b39683ca125d4aac7bcd6"
$ClientSec = "QgOkaL8zgDa7mAcnSlMh3YbYuMgp1F5y"
$Stamp     = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"

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
    try {
        $resp = Invoke-RestMethod -Uri "https://kauth.kakao.com/oauth/token" `
            -Method POST -Body $body -TimeoutSec 15
    } catch {
        # 카카오가 400 (KOE322 등) 리턴 - 토큰 만료. 다음 rotation 도 실패할 것이므로
        # 사용자에게 명확히 알림.
        $errMsg = $_.Exception.Message
        Log "ERR 카카오 refresh 실패 (사용자 재인가 필요): $errMsg"
        exit 1
    }

    # 카카오는 만료 임박(1개월 이내) 시에만 새 refresh_token 을 함께 반환한다.
    if ($resp.refresh_token -and $resp.refresh_token -ne $current) {
        @{ refresh_token = $resp.refresh_token; rotated_at = $Stamp } |
            ConvertTo-Json | Set-Content -Path $TokenFile -Encoding UTF8

        $preview = $resp.refresh_token.Substring(0, 12)
        Log "ROTATE 새 refresh_token 감지 → Secret 갱신 중 ($preview...)"
        $result = & $GhExe secret set KAKAO_REFRESH_TOKEN --repo $Repo --body $resp.refresh_token 2>&1
        if ($LASTEXITCODE -eq 0) {
            Log "OK Secret 갱신 완료"
            exit 0
        } else {
            Log "ERR gh secret set 실패 ($LASTEXITCODE): $result"
            exit 1
        }
    }

    Log "OK 갱신 불필요 (카카오가 새 토큰 미반환 = 아직 여유 있음)"
    exit 0
} catch {
    Log "ERR 예외: $_"
    exit 1
}
