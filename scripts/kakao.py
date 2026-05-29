"""카카오톡 '나에게 보내기' 메시지 발송.

토큰 흐름:
1. GitHub Secrets에 KAKAO_REST_KEY, KAKAO_REFRESH_TOKEN 저장.
2. 매 실행마다 refresh_token으로 새 access_token 발급.
3. 카카오가 refresh_token도 갱신해서 응답으로 돌려주는 경우(만료 1개월 전부터)가 있으므로,
   응답에 새 refresh_token이 있으면 GitHub Actions output으로 노출 → 워크플로가 Secret을 갱신.

엔드포인트: https://kapi.kakao.com/v2/api/talk/memo/default/send
템플릿: feed (제목 + 본문 + 링크 버튼)
메시지 한도: 약 2KB(template_object JSON 기준). 신규 건이 많으면 상위 N건만 노출하고 나머지는 대시보드 링크로.
"""

from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

import httpx

logger = logging.getLogger(__name__)

ROOT = Path(__file__).resolve().parent.parent
NEW_ITEMS_FILE = ROOT / "data" / "new_items.json"
TOKEN_OUTPUT_FILE = ROOT / ".tokens.json"   # 새 refresh_token이 발급되면 여기에 기록

MAX_ITEMS_IN_MSG = 5  # 카톡 메시지 한도(2KB) 안에 들어가도록


def refresh_access_token(rest_key: str, refresh_token: str,
                         client_secret: str | None = None) -> dict[str, Any]:
    """refresh_token으로 access_token 갱신.
    응답에 refresh_token이 포함되어 있을 수도(만료 1개월 전), 없을 수도 있다.
    앱에 '클라이언트 시크릿'이 활성화되어 있으면 client_secret 도 함께 보내야 한다(미전송 시 401).
    """
    data = {
        "grant_type": "refresh_token",
        "client_id": rest_key,
        "refresh_token": refresh_token,
    }
    if client_secret:
        data["client_secret"] = client_secret
    resp = httpx.post(
        "https://kauth.kakao.com/oauth/token",
        data=data,
        timeout=15.0,
    )
    if resp.status_code != 200:
        logger.error("토큰 갱신 실패 %s: %s", resp.status_code, resp.text)
        resp.raise_for_status()
    return resp.json()


def send_message(access_token: str, dashboard_url: str, new_items: list[dict[str, Any]],
                 stats: dict[str, Any]) -> None:
    if not new_items:
        # 신규 건 0이어도 1회 핑은 보내서 실행이 잘 됐는지 확인 가능
        text = f"이번 주 새 입찰 건수 0건. (누적 {stats.get('total', 0)}건)\n대시보드: {dashboard_url}"
        template = {
            "object_type": "text",
            "text": text,
            "link": {"web_url": dashboard_url, "mobile_web_url": dashboard_url},
            "button_title": "대시보드 보기",
        }
    else:
        top = new_items[:MAX_ITEMS_IN_MSG]
        lines = [f"이번 주 신규 {len(new_items)}건 (상위 {len(top)}건 표시)"]
        for it in top:
            kws = ", ".join(it.get("matched_keywords", [])[:3])
            agency = it.get("agency") or "기관미상"
            budget = it.get("budget_amount")
            budget_str = f" / {budget:,}원" if isinstance(budget, int) else ""
            lines.append(f"• [{kws}] {it['title']}\n  {agency}{budget_str}")
        if len(new_items) > MAX_ITEMS_IN_MSG:
            lines.append(f"\n외 {len(new_items) - MAX_ITEMS_IN_MSG}건 → 대시보드")
        text = "\n".join(lines)
        # 카카오 text 템플릿은 최대 200자가 권장이지만 실제로는 더 들어감.
        # 안전하게 1900자로 truncate.
        if len(text) > 1900:
            text = text[:1880] + "\n…(잘림)"
        template = {
            "object_type": "text",
            "text": text,
            "link": {"web_url": dashboard_url, "mobile_web_url": dashboard_url},
            "button_title": "대시보드에서 전체 보기",
        }

    resp = httpx.post(
        "https://kapi.kakao.com/v2/api/talk/memo/default/send",
        headers={"Authorization": f"Bearer {access_token}"},
        data={"template_object": json.dumps(template, ensure_ascii=False)},
        timeout=15.0,
    )
    if resp.status_code != 200:
        logger.error("카톡 발송 실패 %s: %s", resp.status_code, resp.text)
        resp.raise_for_status()
    logger.info("카톡 발송 완료: %s", resp.json())


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    rest_key = os.environ.get("KAKAO_REST_KEY", "").strip()
    refresh_token = os.environ.get("KAKAO_REFRESH_TOKEN", "").strip()
    client_secret = os.environ.get("KAKAO_CLIENT_SECRET", "").strip() or None
    dashboard_url = os.environ.get("DASHBOARD_URL", "").strip() or "https://example.github.io/"

    if not (rest_key and refresh_token):
        logger.error("KAKAO_REST_KEY / KAKAO_REFRESH_TOKEN 환경변수가 비어 있습니다.")
        return 1

    token_data = refresh_access_token(rest_key, refresh_token, client_secret)
    access_token = token_data["access_token"]
    new_refresh = token_data.get("refresh_token")
    if new_refresh and new_refresh != refresh_token:
        logger.info("새 refresh_token 발급됨 — Secret 갱신 필요")
        TOKEN_OUTPUT_FILE.write_text(
            json.dumps({"refresh_token": new_refresh}, ensure_ascii=False),
            encoding="utf-8",
        )
        # GitHub Actions output (workflow가 이걸로 Secret 갱신할 수 있음)
        gh_output = os.environ.get("GITHUB_OUTPUT")
        if gh_output:
            with open(gh_output, "a", encoding="utf-8") as f:
                f.write(f"new_refresh_token={new_refresh}\n")

    # 신규 건 로드
    new_items: list[dict[str, Any]] = []
    stats: dict[str, Any] = {}
    if NEW_ITEMS_FILE.exists():
        payload = json.loads(NEW_ITEMS_FILE.read_text(encoding="utf-8"))
        new_items = payload.get("items", [])
        stats = payload.get("generated_at", {}) if isinstance(payload.get("generated_at"), dict) else {}

    send_message(access_token, dashboard_url, new_items, stats)
    return 0


if __name__ == "__main__":
    sys.exit(main())
