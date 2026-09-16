"""기업마당(bizinfo.go.kr) 지원사업 공고 목록 수집.

목록 URL: /web/lay1/bbs/S1T122C128/AS/74/list.do?cpage=N&rows=15
     → 302 redirect → /sii/siia/selectSIIA200View.do?...
컬럼: 번호 · 지원분야 · 지원사업명(a href=selectSIIA200Detail.do?pblancId=X) · 신청기간 · 소관부처 · 사업수행기관 · 등록일 · 조회수
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
BASE = "https://www.bizinfo.go.kr"
LIST_URL = f"{BASE}/web/lay1/bbs/S1T122C128/AS/74/list.do"
DETAIL_URL_TMPL = f"{BASE}/sii/siia/selectSIIA200Detail.do?pblancId={{pid}}"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept-Language": "ko,en-US;q=0.7,en;q=0.3",
}

# 농림·산림·연구 관련성이 있어보이는 지원분야만 통과.
# 기업마당은 창업/금융/수출 등 대부분 INK 업무와 무관 → 여기서 1차 필터.
FIELD_ALLOW = ("경영", "기술", "인력", "창업", "기타")


def _parse_period_end(text: str) -> str | None:
    """'2026-09-15 ~ 2026-10-06' 형태에서 종료일 추출."""
    if not text:
        return None
    m = re.search(r"~\s*(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", text)
    if not m:
        return None
    y, mo, d = m.groups()
    try:
        return datetime(int(y), int(mo), int(d), 23, 59, tzinfo=KST).isoformat()
    except ValueError:
        return None


def _parse_date(text: str) -> str | None:
    if not text:
        return None
    m = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", text)
    if not m:
        return None
    y, mo, d = m.groups()
    try:
        return datetime(int(y), int(mo), int(d), tzinfo=KST).isoformat()
    except ValueError:
        return None


def _parse_row(tr) -> dict[str, Any] | None:
    tds = tr.find_all("td")
    if len(tds) < 7:
        return None
    a = tr.find("a")
    href = a.get("href", "") if a else ""
    m = re.search(r"pblancId=(PBLN_\d+)", href)
    pblanc_id = m.group(1) if m else ""
    if not pblanc_id:
        return None
    title = (a.get_text(strip=True) if a else tds[2].get_text(strip=True)).strip()
    field = tds[1].get_text(strip=True)
    period = tds[3].get_text(" ", strip=True)
    ministry = tds[4].get_text(strip=True)
    exec_agency = tds[5].get_text(strip=True)
    reg_date_txt = tds[6].get_text(strip=True)
    return {
        "no": tds[0].get_text(strip=True),
        "pblanc_id": pblanc_id,
        "title": title,
        "field": field,
        "period": period,
        "ministry": ministry,
        "exec_agency": exec_agency,
        "reg_date": _parse_date(reg_date_txt),
        "deadline": _parse_period_end(period),
    }


def _classify_region(text: str) -> str:
    if not text:
        return "중앙/전국"
    for tok in ("서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
                "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"):
        if tok in text:
            return tok if tok not in ("전북", "전남", "경북", "경남") else tok
    return "중앙/전국"


def _to_item(row: dict[str, Any]) -> dict[str, Any]:
    detail = DETAIL_URL_TMPL.format(pid=row["pblanc_id"])
    agency = row["ministry"] or row["exec_agency"] or "기업마당"
    region_src = f"{row['ministry']} {row['exec_agency']} {row['title']}"
    description = "\n".join([
        f"지원분야: {row['field']}",
        f"신청기간: {row['period']}",
        f"소관부처·지자체: {row['ministry']}",
        f"사업수행기관: {row['exec_agency']}",
        "출처: 기업마당(bizinfo)",
    ])
    return {
        "source_type": "bizinfo_support",
        "external_id": f"bizinfo_support::{row['pblanc_id']}",
        "title": row["title"],
        "agency": agency,
        "agency_dept": row["exec_agency"] or None,
        "agency_type": "지자체" if any(k in row["ministry"] for k in ("도", "시", "군", "구")) else "공공기관",
        "contract_method": None,
        "bsns_div": row["field"],  # 지원분야 (경영/기술/창업 등)
        "budget_amount": None,
        "order_planned_date": row["reg_date"],
        "deadline": row["deadline"],
        "region": _classify_region(region_src),
        "description": description,
        "attachments": [],
        "url": detail,
        "ref_no": row["pblanc_id"],
        "officer": None,
        "officer_tel": None,
        "source_label": "💼 기업마당",
    }


def collect(max_pages: int = 6, lookback_days: int = 60) -> list[dict[str, Any]]:
    cutoff = datetime.now(KST) - timedelta(days=lookback_days)
    items: list[dict[str, Any]] = []
    with httpx.Client(follow_redirects=True) as client:
        for page in range(1, max_pages + 1):
            try:
                resp = client.get(
                    LIST_URL, params={"cpage": page, "rows": 15},
                    headers=DEFAULT_HEADERS, timeout=20,
                )
                resp.raise_for_status()
            except Exception as exc:
                logger.warning("bizinfo fetch fail page=%d: %s", page, exc)
                break
            soup = BeautifulSoup(resp.content, "html.parser")
            tbody = soup.find("tbody")
            rows = tbody.find_all("tr", recursive=False) if tbody else []
            if not rows:
                break
            saw_old = 0
            page_items = 0
            for tr in rows:
                r = _parse_row(tr)
                if not r:
                    continue
                page_items += 1
                # 지원분야 화이트리스트 (수출/내수/금융/자금은 INK와 무관)
                if r["field"] and r["field"] not in FIELD_ALLOW:
                    continue
                if r["reg_date"]:
                    reg_dt = datetime.fromisoformat(r["reg_date"])
                    if reg_dt < cutoff:
                        saw_old += 1
                        continue
                items.append(_to_item(r))
            if page_items == 0:
                break
            if saw_old >= page_items:
                break
    logger.info("bizinfo collected=%d (max_pages=%d, lookback=%d)", len(items), max_pages, lookback_days)
    return items


if __name__ == "__main__":  # pragma: no cover
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    got = collect()
    import json as _json
    print(_json.dumps(got[:5], ensure_ascii=False, indent=2))
    print(f"total: {len(got)}")
