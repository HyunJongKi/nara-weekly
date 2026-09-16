"""한국임업진흥원(kofpi.or.kr) 알림/PR - 공지사항 목록 수집.

목록 URL: https://www.kofpi.or.kr/notice/notice_01.do
페이지네이션: POST body {target=total, cPage=N, searchValue=""}
행 셀렉터: table tbody tr (td 5개: 번호·제목·첨부·조회수·작성일)
카테고리 배지: .badge-enotice(긴급) .badge-notice(공지) .badge-edu(교육/행사) .badge-bid(입찰/공모)
상세 URL: /notice/notice_01view.do?bb_seq={bb_seq}   ← 목록 링크의 fnGoView('bb_seq')
"""

from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

KST = timezone(timedelta(hours=9))
BASE = "https://www.kofpi.or.kr"
LIST_URL = f"{BASE}/notice/notice_01.do"
DETAIL_URL_TMPL = f"{BASE}/notice/notice_01view.do?bb_seq={{seq}}"

BADGE_LABEL = {
    "badge-enotice": "긴급",
    "badge-notice": "공지",
    "badge-edu": "교육/행사",
    "badge-bid": "입찰/공모",
}

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept-Language": "ko,en-US;q=0.7,en;q=0.3",
}


def _fetch_page(client: httpx.Client, cpage: int) -> str:
    if cpage == 1:
        resp = client.get(LIST_URL, headers=DEFAULT_HEADERS, timeout=20)
    else:
        resp = client.post(
            LIST_URL,
            data={"target": "total", "cPage": str(cpage), "searchValue": ""},
            headers=DEFAULT_HEADERS,
            timeout=20,
        )
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return resp.text


def _parse_deadline(title: str) -> str | None:
    """제목 안의 (~9.30) / (~2026-09-30) 같은 마감 표기 추출."""
    if not title:
        return None
    now_year = datetime.now(KST).year
    m = re.search(r"~\s*(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})", title)
    if m:
        y, mo, d = m.groups()
    else:
        m = re.search(r"~\s*(\d{1,2})[.\-/](\d{1,2})(?!\d)", title)
        if not m:
            return None
        mo, d = m.groups()
        y = str(now_year)
    try:
        dt = datetime(int(y), int(mo), int(d), 23, 59, tzinfo=KST)
        return dt.isoformat()
    except ValueError:
        return None


def _parse_row(tr) -> dict[str, Any] | None:
    tds = tr.find_all("td")
    if len(tds) < 5:
        return None
    num_txt = tds[0].get_text(strip=True)
    if not num_txt.isdigit():
        return None
    title_td = tds[1]
    a = title_td.find("a")
    onclick = (a.get("onclick") or "") if a else ""
    m = re.search(r"fnGoView\('([^']+)'\)", onclick)
    bb_seq = m.group(1) if m else ""
    badges: list[str] = []
    for b in title_td.select(".badge"):
        for cls in b.get("class", []):
            if cls.startswith("badge-") and cls in BADGE_LABEL:
                badges.append(BADGE_LABEL[cls])
        b.decompose()
    for extra in title_td.select(".mo_con"):
        extra.decompose()
    title = (a.get_text(strip=True) if a else title_td.get_text(strip=True)).strip()
    date_str = tds[4].get_text(strip=True)
    try:
        posted_dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=KST)
        posted = posted_dt.isoformat()
    except ValueError:
        posted = None
    return {
        "num": num_txt,
        "bb_seq": bb_seq,
        "title": title,
        "badges": badges,
        "posted": posted,
    }


def _to_item(row: dict[str, Any]) -> dict[str, Any]:
    seq = row["bb_seq"]
    detail = DETAIL_URL_TMPL.format(seq=seq) if seq else LIST_URL
    deadline = _parse_deadline(row["title"])
    category = "/".join(row["badges"]) if row["badges"] else None
    description_parts = [f"카테고리: {category}"] if category else []
    description_parts.append("출처: 한국임업진흥원")
    return {
        "source_type": "kofpi_bid",
        "external_id": f"kofpi_bid::{seq or row['num']}",
        "title": row["title"],
        "agency": "한국임업진흥원",
        "agency_dept": None,
        "agency_type": "공공기관",
        "contract_method": None,
        "bsns_div": category,  # 배지 카테고리 (입찰/공모 등)
        "budget_amount": None,
        "order_planned_date": row["posted"],
        "deadline": deadline,
        "region": "중앙/전국",
        "description": "\n".join(description_parts),
        "attachments": [],
        "url": detail,
        "ref_no": row["num"],
        "officer": None,
        "officer_tel": None,
        "source_label": "🌲 임업진흥원",
    }


def collect(max_pages: int = 5, lookback_days: int = 90) -> list[dict[str, Any]]:
    """kofpi 목록에서 최근 lookback_days 이내 항목만 정규화 반환."""
    cutoff = datetime.now(KST) - timedelta(days=lookback_days)
    items: list[dict[str, Any]] = []
    with httpx.Client() as client:
        for page in range(1, max_pages + 1):
            try:
                html = _fetch_page(client, page)
            except Exception as exc:
                logger.warning("kofpi fetch fail page=%d: %s", page, exc)
                break
            soup = BeautifulSoup(html, "html.parser")
            tbody = soup.select_one("table tbody")
            rows = tbody.find_all("tr") if tbody else []
            if not rows:
                break
            saw_old = 0
            page_items = 0
            for tr in rows:
                r = _parse_row(tr)
                if not r:
                    continue
                page_items += 1
                if r["posted"]:
                    posted_dt = datetime.fromisoformat(r["posted"])
                    if posted_dt < cutoff:
                        saw_old += 1
                        continue
                items.append(_to_item(r))
            if page_items == 0:
                break
            # 페이지 전체가 cutoff 이전이면 더 이상 조회할 필요 없음
            if saw_old >= page_items:
                break
    logger.info("kofpi collected=%d (max_pages=%d, lookback=%d)", len(items), max_pages, lookback_days)
    return items


if __name__ == "__main__":  # pragma: no cover
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    got = collect()
    import json as _json
    print(_json.dumps(got[:5], ensure_ascii=False, indent=2))
    print(f"total: {len(got)}")
