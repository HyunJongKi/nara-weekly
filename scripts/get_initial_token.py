"""최초 1회만 로컬에서 실행하는 헬퍼.

Kakao Developers에서 인가 코드를 받아 access_token + refresh_token을 발급한다.
발급된 refresh_token을 GitHub Secrets의 KAKAO_REFRESH_TOKEN에 저장하면 끝.

사용법:
  1) Kakao Developers에 앱을 만들고 REST API 키, Redirect URI를 등록한다.
     (Redirect URI는 https://localhost.test 처럼 도달 불가한 주소여도 OK — 인가 코드만 URL에서 복사하면 됨)
  2) 동의항목에 '카카오톡 메시지 전송 (talk_message)' 을 추가한다.
  3) 아래 URL을 브라우저에서 연다 (REST_KEY, REDIRECT_URI 치환):
     https://kauth.kakao.com/oauth/authorize?response_type=code&client_id={REST_KEY}&redirect_uri={REDIRECT_URI}&scope=talk_message
  4) 로그인 + 동의하면 브라우저가 REDIRECT_URI?code=XXX 로 리디렉트된다. 주소창에서 code 값을 복사.
  5) 이 스크립트 실행:
       python scripts/get_initial_token.py {REST_KEY} {REDIRECT_URI} {code}
  6) 출력된 refresh_token 을 GitHub Secrets에 KAKAO_REFRESH_TOKEN 으로 저장.
"""

from __future__ import annotations

import json
import sys

import httpx


def main(rest_key: str, redirect_uri: str, code: str, client_secret: str | None = None) -> int:
    data = {
        "grant_type": "authorization_code",
        "client_id": rest_key,
        "redirect_uri": redirect_uri,
        "code": code,
    }
    if client_secret:
        data["client_secret"] = client_secret
    resp = httpx.post("https://kauth.kakao.com/oauth/token", data=data, timeout=15.0)
    if resp.status_code != 200:
        print(f"ERROR {resp.status_code}: {resp.text}")
        resp.raise_for_status()
    data = resp.json()
    print(json.dumps(data, ensure_ascii=False, indent=2))
    print("\n=== 다음 단계 ===")
    print("GitHub Secrets에 아래 값들을 등록하세요:")
    print(f"  KAKAO_REST_KEY        = {rest_key}")
    print(f"  KAKAO_REFRESH_TOKEN   = {data.get('refresh_token')}")
    if client_secret:
        print(f"  KAKAO_CLIENT_SECRET   = {client_secret}")
    return 0


if __name__ == "__main__":
    if len(sys.argv) not in (4, 5):
        print("usage: python scripts/get_initial_token.py REST_KEY REDIRECT_URI CODE [CLIENT_SECRET]")
        sys.exit(2)
    cs = sys.argv[4] if len(sys.argv) == 5 else None
    sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3], cs))
