"""과학기술정보통신부 사업공고 OpenAPI 수집기.

공공데이터포털: https://www.data.go.kr/data/15074634/openapi.do
Endpoint: https://apis.data.go.kr/1721000/msitannouncementinfo/businessAnnouncMentList
Response: XML 전용 (JSON 지원 없음)

응답 필드:
  subject, viewUrl(nttSeqNo=... 가 unique), deptName, managerName, managerTel,
  pressDt (YYYY-MM-DD), files/file/(fileName, fileUrl)

인증키: 환경변수 IRIS_SERVICE_KEY (없으면 G2B_SERVICE_KEY 로 폴백).
       공공데이터포털에서 데이터별로 다른 키가 발급되므로 별도 관리를 권장.
"""

from __future__ import annotations

import logging
import os
import re
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

KST = timezone(timedelta(hours=9))
ENDPOINT = "https://apis.data.go.kr/1721000/msitannouncementinfo/businessAnnouncMentList"


def _get_key() -> str:
    return (
        os.environ.get("IRIS_SERVICE_KEY", "").strip()
        or os.environ.get("G2B_SERVICE_KEY", "").strip()
    )


def _parse_ntt_seq(view_url: str) -> str:
    m = re.search(r"nttSeqNo=(\d+)", view_url or "")
    return m.group(1) if m else ""


def _parse_date(raw: str | None) -> str | None:
    if not raw:
        return None
    raw = raw.strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(raw, fmt).replace(tzinfo=KST).isoformat()
        except ValueError:
            continue
    return None


def _to_item(node) -> dict[str, Any] | None:
    subject = (node.find("subject").text if node.find("subject") else "").strip()
    view_url = (node.find("viewUrl").text if node.find("viewUrl") else "").strip()
    ntt_seq = _parse_ntt_seq(view_url)
    if not (subject and ntt_seq):
        return None
    dept = (node.find("deptName").text if node.find("deptName") else "").strip() or None
    manager = (node.find("managerName").text if node.find("managerName") else "").strip() or None
    tel = (node.find("managerTel").text if node.find("managerTel") else "").strip() or None
    posted = _parse_date(node.find("pressDt").text if node.find("pressDt") else None)

    attachments: list[dict[str, str]] = []
    seen_names: set[str] = set()
    for f in node.find_all("file"):
        name = (f.find("fileName").text if f.find("fileName") else "").strip()
        url = (f.find("fileUrl").text if f.find("fileUrl") else "").strip()
        if not (name and url):
            continue
        # 같은 문서의 .hwp/.hwpx/.odt 중복은 hwpx 우선으로 하나만 유지
        base = re.sub(r"\.(hwp|hwpx|odt|pdf|docx)$", "", name, flags=re.IGNORECASE)
        if base in seen_names and name.lower().endswith((".hwp", ".odt")):
            continue
        seen_names.add(base)
        attachments.append({"name": name, "url": url})

    return {
        "source_type": "iris_rnd",
        "external_id": f"iris_rnd::{ntt_seq}",
        "title": subject,
        "agency": "과학기술정보통신부",
        "agency_dept": dept,
        "agency_type": "중앙부처",
        "contract_method": None,
        "bsns_div": "사업공고",
        "budget_amount": None,
        "order_planned_date": posted,
        "deadline": None,  # 목록 API 에는 접수마감 필드 없음 (상세 페이지에서만 확인 가능)
        "region": "중앙/전국",
        "description": "\n".join(p for p in [
            f"소관부처: 과학기술정보통신부",
            f"담당부서: {dept}" if dept else None,
            "출처: 범부처통합연구지원시스템(IRIS)",
        ] if p),
        "attachments": attachments,
        "url": view_url or None,
        "ref_no": ntt_seq,
        "officer": manager,
        "officer_tel": tel,
        "source_label": "🔬 IRIS",
    }


def collect(max_pages: int = 5, num_of_rows: int = 100, lookback_days: int = 90) -> list[dict[str, Any]]:
    key = _get_key()
    if not key:
        logger.info("iris skip: neither IRIS_SERVICE_KEY nor G2B_SERVICE_KEY set")
        return []
    cutoff = datetime.now(KST) - timedelta(days=lookback_days)
    items: list[dict[str, Any]] = []
    with httpx.Client(timeout=30.0) as client:
        for page in range(1, max_pages + 1):
            params = {"serviceKey": key, "pageNo": page, "numOfRows": num_of_rows}
            try:
                resp = client.get(ENDPOINT, params=params)
                # 활용신청 미승인 케이스 = 400 + errMsg=NO_OPENAPI_SERVICE_ERROR 또는 SERVICE_KEY_IS_NOT_REGISTERED_ERROR
                if resp.status_code in (401, 403):
                    logger.info("iris auth pending (status=%d) → skip", resp.status_code)
                    return []
                resp.raise_for_status()
            except Exception as exc:
                logger.warning("iris fetch fail page=%d: %s", page, exc)
                return items
            body = resp.text
            if "NO_OPENAPI_SERVICE_ERROR" in body or "SERVICE_KEY_IS_NOT_REGISTERED_ERROR" in body:
                logger.info("iris service not registered yet → skip")
                return []
            soup = BeautifulSoup(body, "xml")
            code = soup.find("resultCode")
            if code and code.text.strip() not in ("00", "0"):
                logger.warning("iris result err: %s / %s", code.text,
                               (soup.find("resultMsg").text if soup.find("resultMsg") else ""))
                return items
            nodes = soup.find_all("item")
            if not nodes:
                break
            saw_old = 0
            for n in nodes:
                item = _to_item(n)
                if not item:
                    continue
                if item["order_planned_date"]:
                    dt = datetime.fromisoformat(item["order_planned_date"])
                    if dt < cutoff:
                        saw_old += 1
                        continue
                items.append(item)
            if saw_old >= len(nodes):
                break
            total = int((soup.find("totalCount").text if soup.find("totalCount") else "0") or 0)
            if page * num_of_rows >= total:
                break
    logger.info("iris collected=%d (max_pages=%d, lookback=%d)", len(items), max_pages, lookback_days)
    return items


if __name__ == "__main__":  # pragma: no cover
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    got = collect()
    import json as _json
    print(_json.dumps(got[:3], ensure_ascii=False, indent=2))
    print(f"total: {len(got)}")
