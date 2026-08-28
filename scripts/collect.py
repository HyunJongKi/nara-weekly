"""나라장터 OpenAPI에서 지난 N일치 발주계획·사전규격을 수집해서
키워드 매칭 결과를 data/tenders.json에 누적 저장한다.

기존 Nara Radar 프로젝트(C:\\Users\\CEO\\nara-radar)에서 실측 검증된
엔드포인트·파라미터 매핑을 그대로 사용한다.
"""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable

import httpx
import yaml
from tenacity import retry, stop_after_attempt, wait_exponential

logger = logging.getLogger(__name__)

KST = timezone(timedelta(hours=9))
BASE_URL = "https://apis.data.go.kr/1230000"

ROOT = Path(__file__).resolve().parent.parent
DATA_FILE = ROOT / "data" / "tenders.json"
DOCS_DATA_FILE = ROOT / "docs" / "tenders.json"
KEYWORDS_FILE = ROOT / "keywords.yml"


@dataclass(frozen=True)
class Endpoint:
    path: str
    operation: str
    src_type: str          # "order_plan" | "pre_spec"
    date_param: str        # "order_ym" | "inqry_dt"


ENDPOINTS: list[Endpoint] = [
    Endpoint("/ao/OrderPlanSttusService",     "getOrderPlanSttusListServc",  "order_plan", "order_ym"),
    Endpoint("/ao/HrcspSsstndrdInfoService",  "getPublicPrcureThngInfoServc", "pre_spec",   "inqry_dt"),
]

# ─── 학술연구용역 필터 (재설계) ───────────────────────────────────
# 나라장터 업무구분(bsnsDivNm)을 최우선 판정 기준으로 삼는다.
# - '용역' 카테고리(일반용역/기술용역/학술연구/학술용역)면 원칙적으로 통과
# - '공사', '물품', '외자' 카테고리는 하드 제외
# - 그 위에서 제목·설명에 명백한 비-용역 지시어(공사/구매/급식/청소 등)면 최종 제외
#
# 종전 필터가 너무 엄격해 나라장터 웹 검색 결과 대비 유효 항목 다수를 놓치던 문제 해결.

# 명확한 학술·조사·연구·계획·평가·타당성 표현 (있으면 다른 조건 무관 통과)
RESEARCH_STRONG = re.compile(
    r"(연구\s*용역|학술\s*연구|학술용역|조사\s*용역|조사연구|기초\s*연구|기획\s*연구|"
    r"정책\s*연구|타당성\s*(조사|분석|연구|검토)|실태\s*조사|모니터링|연구개발|R&D|"
    r"마스터플랜|기본\s*계획\s*수립|중장기\s*계획|성과\s*평가|성과평가|"
    r"역량강화|용역\s*연구|실증\s*연구|시범\s*사업|설계기준|평가\s*연구|"
    r"영향\s*조사|영향조사|실무활용\s*정립|정립\s*연구|중심지\s*활성화|"
    r"공간\s*정비\s*사업|공간정비|재구조화|기술\s*수요|수요\s*분석|기획\s*용역)"
)

# 하드 제외 - 무조건 배제. bsnsDivNm 카테고리와 제목에서 이 표현이 나오면 용역 아님.
RESEARCH_HARD_EXCLUDE = re.compile(
    r"(신축\s*공사|건축\s*공사|토목\s*공사|시공|설치\s*공사|시설\s*공사|"
    r"구매\s*계약|물품\s*구매|물품\s*납품|기자재\s*구매|장비\s*구매|"
    r"임대\s*차|리스\s*계약|급식|청소|경비\s*용역|환경미화|"
    r"보수\s*공사|보강\s*공사|리모델링|철거)"
)

# bsnsDivNm 카테고리별 판정
BD_ALLOW = ("일반용역", "기술용역", "학술연구", "학술용역", "연구용역", "용역")
BD_DENY  = ("공사", "물품", "외자", "리스", "임대")


def _is_research(item: dict[str, Any]) -> bool:
    title_desc = " ".join(filter(None, [item.get("title"), item.get("description")]))
    if not title_desc:
        return False
    bd = (item.get("bsns_div") or "").replace(" ", "")

    # 1) 하드 제외: 신축공사·물품구매·급식 등 명백히 용역 아님
    if RESEARCH_HARD_EXCLUDE.search(title_desc):
        return False

    # 2) bsnsDivNm 이 명시적 공사/물품/외자면 제외
    if any(x in bd for x in BD_DENY):
        return False

    # 3) 명확한 연구·조사·평가 표현 있으면 통과 (모든 다른 조건 무관)
    if RESEARCH_STRONG.search(title_desc):
        return True

    # 4) bsnsDivNm 이 용역 카테고리이고, 제목이 연구·조사 성격 있으면 통과
    if any(x in bd for x in BD_ALLOW):
        # 용역 카테고리 내 하위 필터 - 연구·조사·분석·평가·계획·기획·전략·개발 등
        if re.search(r"연구|조사|분석|평가|계획|기획|전략|정책|타당성|용역|"
                     r"컨설팅|자문|진단|검토|수립|개발|설계|기준|기본구상|"
                     r"활성화|정비|육성|경영체|기술\s*보급|보급|육종|품종|"
                     r"인재\s*양성|양성|교육\s*훈련|프로그램", title_desc):
            return True

    return False


# ─── 지역·발주처 분류 ──────────────────────────────────────────────

REGION_TOKENS = [
    ("서울", "서울"), ("부산", "부산"), ("대구", "대구"), ("인천", "인천"),
    ("광주", "광주"), ("대전", "대전"), ("울산", "울산"), ("세종", "세종"),
    ("경기", "경기"), ("강원", "강원"), ("충북", "충북"), ("충남", "충남"),
    ("전북", "전북"), ("전라북도", "전북"), ("전남", "전남"), ("전라남도", "전남"),
    ("경북", "경북"), ("경상북도", "경북"), ("경남", "경남"), ("경상남도", "경남"),
    ("제주", "제주"),
]


def _extract_region(*texts: str | None) -> str:
    blob = " ".join(t for t in texts if t)
    for tok, label in REGION_TOKENS:
        if tok in blob:
            return label
    return "중앙/전국"


def _classify_agency(name: str | None) -> str:
    """발주기관명을 중앙부처/지자체/공공기관/교육기관/기타로 분류.
    이름의 첫 토큰(상위 조직)을 기준으로 판정하므로
    '농촌진흥청 식량과학원' 같이 하부조직이 붙어도 '중앙부처'로 잡힌다.
    """
    if not name:
        return "기타"
    n = name.strip()
    head = n.split()[0] if n.split() else n
    n_nospace = n.replace(" ", "")
    # 교육
    if any(k in n_nospace for k in ("교육청", "교육지원청", "대학교", "교육원", "학교법인")) \
       or head.endswith(("초등학교", "중학교", "고등학교", "학교")):
        return "교육기관"
    # 중앙부처 (행정조직). '위원회'/'..원회'도 포함.
    if head.endswith(("부", "처", "청")) or head.endswith("위원회") or head.endswith("원회"):
        return "중앙부처"
    # 지자체
    if any(k in n_nospace for k in ("도청", "시청", "군청", "구청", "도의회", "시의회")) \
       or head.endswith(("도", "시", "군", "구")):
        return "지자체"
    # 공공기관 (공사/공단/재단/진흥원/연구원/과학원/센터/협회/사업단 등 + '..원/..관'로 끝나는 단어)
    if any(k in head for k in ("공사", "공단", "재단", "진흥원", "연구원", "과학원",
                                "센터", "협회", "사업단", "박물관", "미술관", "수목원", "도서관")) \
       or head.endswith(("원", "관", "단")):
        return "공공기관"
    return "기타"


# ─── 유틸 ──────────────────────────────────────────────────────

def _nonempty(*values: Any) -> str | None:
    for v in values:
        if v not in (None, ""):
            s = str(v).strip()
            if s:
                return s
    return None


def _to_int(value: Any) -> int | None:
    if value in (None, "", "null"):
        return None
    try:
        return int(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _parse_dt(value: str | None) -> str | None:
    """다양한 포맷의 날짜를 ISO 8601(KST) 문자열로."""
    if not value:
        return None
    value = str(value).strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y%m%d%H%M", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).replace(tzinfo=KST).isoformat()
        except ValueError:
            continue
    return None


# ─── OpenAPI 호출 ──────────────────────────────────────────────

@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=1, max=10))
def _fetch(url: str, params: dict[str, Any]) -> dict[str, Any]:
    with httpx.Client(timeout=30.0) as client:
        resp = client.get(url, params=params)
        resp.raise_for_status()
        return resp.json()


def _iter_pages(ep: Endpoint, service_key: str, lookback: int, lookahead: int) -> Iterable[dict[str, Any]]:
    url = f"{BASE_URL}{ep.path}/{ep.operation}"
    now = datetime.now(KST)
    if ep.date_param == "order_ym":
        date_params = {
            "orderBgnYm": (now - timedelta(days=lookback)).strftime("%Y%m"),
            "orderEndYm": (now + timedelta(days=lookahead)).strftime("%Y%m"),
        }
    else:
        date_params = {
            "inqryBgnDt": (now - timedelta(days=lookback)).strftime("%Y%m%d") + "0000",
            "inqryEndDt": (now + timedelta(days=lookahead)).strftime("%Y%m%d") + "2359",
        }

    page = 1
    rows = 100
    while True:
        params = {
            "serviceKey": service_key,
            "type": "json",
            "pageNo": page,
            "numOfRows": rows,
            "inqryDiv": "1",
            **date_params,
        }
        try:
            data = _fetch(url, params)
        except Exception as exc:
            logger.warning("fetch failed page=%d op=%s: %s", page, ep.operation, exc)
            break

        if isinstance(data, dict) and "nkoneps.com.response.ResponseError" in data:
            err = data["nkoneps.com.response.ResponseError"].get("header", {})
            logger.error("API error %s: %s", err.get("resultCode"), err.get("resultMsg"))
            break

        body = data.get("response", {}).get("body") or data.get("body") or {}
        items_node = body.get("items") or {}
        if isinstance(items_node, dict):
            raw_items = items_node.get("item") or []
        elif isinstance(items_node, list):
            raw_items = items_node
        else:
            raw_items = []
        if isinstance(raw_items, dict):
            raw_items = [raw_items]

        if not raw_items:
            break

        for it in raw_items:
            yield it

        total = _to_int(body.get("totalCount")) or 0
        if page * rows >= total:
            break
        page += 1
        if page > 50:
            logger.warning("page guard at %s", ep.operation)
            break


# ─── 정규화 ────────────────────────────────────────────────────

def _normalize_order_plan(raw: dict[str, Any]) -> dict[str, Any] | None:
    title = _nonempty(raw.get("bizNm"))
    ext_id = _nonempty(
        raw.get("orderPlanUntyNo"),
        "-".join(
            str(raw.get(k, "")).strip()
            for k in ("orderInsttCd", "orderYear", "orderMnth", "orderPlanSno")
            if raw.get(k) not in (None, "")
        ),
    )
    if not (title and ext_id):
        return None

    spec_items = " ".join(
        _nonempty(raw.get(f"specItemNm{i}"), raw.get(f"specItemCntnts{i}")) or ""
        for i in range(1, 6)
    ).strip()
    description = "\n".join(p for p in [
        _nonempty(raw.get("usgCntnts")),
        _nonempty(raw.get("specCntnts")),
        spec_items or None,
        _nonempty(raw.get("rmrkCntnts")),
    ] if p) or None

    agency = _nonempty(raw.get("orderInsttNm"), raw.get("totlmngInsttNm"))
    raw_region = _nonempty(raw.get("cnstwkRgnNm"), raw.get("jrsdctnDivNm"))
    return {
        "source_type": "order_plan",
        "external_id": f"order_plan::{ext_id}",
        "title": title,
        "agency": agency,
        "agency_dept": _nonempty(raw.get("deptNm")),
        "agency_type": _classify_agency(agency),
        "contract_method": _nonempty(raw.get("cntrctMthdNm")),
        "bsns_div": _nonempty(raw.get("bsnsDivNm")),
        "budget_amount": _to_int(raw.get("sumOrderAmt")),
        "order_planned_date": _parse_dt(raw.get("nticeDt")),
        "deadline": None,   # 발주계획은 입찰마감 개념 없음(예정) → 경과는 last_seen 기준으로 판정
        "region": _extract_region(raw_region, agency),
        "description": description,
        "attachments": [],  # 발주계획 단계에는 첨부 없음 (입찰공고 단계에서 생김)
        "url": None,
        "ref_no": _nonempty(raw.get("orderPlanUntyNo")),
        "officer": _nonempty(raw.get("ofclNm")),
        "officer_tel": _nonempty(raw.get("telNo")),
    }


def _normalize_pre_spec(raw: dict[str, Any]) -> dict[str, Any] | None:
    title = _nonempty(raw.get("prdctClsfcNoNm"))
    ext_id = _nonempty(raw.get("bfSpecRgstNo"), raw.get("refNo"))
    if not (title and ext_id):
        return None

    # 첨부파일 1~5 (과업지시서/제안요청서/규격서 등). 파일명 필드가 응답에 있을 때만 채움.
    attachments: list[dict[str, str]] = []
    for i in range(1, 6):
        url = _nonempty(raw.get(f"specDocFileUrl{i}"))
        if not url:
            continue
        name = _nonempty(
            raw.get(f"specDocFileNm{i}"),
            raw.get(f"specDocFlNm{i}"),
            raw.get(f"specDocFileNm0{i}"),
        )
        attachments.append({"url": url, "name": name or f"첨부{i}"})

    bsns_div = _nonempty(raw.get("bsnsDivNm"))
    description = "\n".join(p for p in [
        bsns_div,
        f"참조번호: {raw.get('refNo')}" if raw.get("refNo") else None,
        _nonempty(raw.get("prdctDtlList")),
    ] if p) or None

    agency = _nonempty(raw.get("rlDminsttNm"), raw.get("orderInsttNm"))
    return {
        "source_type": "pre_spec",
        "external_id": f"pre_spec::{ext_id}",
        "title": title,
        "agency": agency,
        "agency_dept": None,
        "agency_type": _classify_agency(agency),
        "contract_method": None,
        "bsns_div": bsns_div,
        "budget_amount": _to_int(raw.get("asignBdgtAmt")),
        "order_planned_date": _parse_dt(raw.get("rcptDt") or raw.get("rgstDt")),
        # 의견 등록 마감일 — 이 시각이 지나면 곧 입찰공고 단계로 넘어감(참고 시한 만료).
        # 카카오 raw 필드명은 opninRgstClseDt (Rgst=등록). 이전엔 오타로 채워지지 않았음.
        "deadline": _parse_dt(raw.get("opninRgstClseDt") or raw.get("opninRgtClseDt") or raw.get("opninRgtClsDt")),
        "region": _extract_region(agency),
        "description": description,
        "attachments": attachments,
        "url": attachments[0]["url"] if attachments else None,
        "ref_no": _nonempty(raw.get("refNo")),
        "officer": _nonempty(raw.get("ofclNm")),
        "officer_tel": _nonempty(raw.get("ofclTelNo")),
    }


NORMALIZERS = {"order_plan": _normalize_order_plan, "pre_spec": _normalize_pre_spec}


# ─── 키워드 매칭 ────────────────────────────────────────────────

def _load_keywords() -> tuple[list[dict[str, Any]], int]:
    with KEYWORDS_FILE.open(encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return cfg.get("keywords", []), int(cfg.get("min_score", 0))


def _score_item(item: dict[str, Any], keywords: list[dict[str, Any]]) -> tuple[int, list[str]]:
    haystack = " ".join(filter(None, [
        item.get("title"), item.get("description"), item.get("agency"), item.get("agency_dept"),
    ])).lower()
    score = 0
    matched: list[str] = []
    for kw in keywords:
        term = kw["term"].lower()
        if term in haystack:
            score += int(kw.get("weight", 1))
            matched.append(kw["term"])
    return score, matched


# ─── 메인 ──────────────────────────────────────────────────────

def load_existing() -> dict[str, dict[str, Any]]:
    if not DATA_FILE.exists():
        return {}
    try:
        with DATA_FILE.open(encoding="utf-8") as f:
            payload = json.load(f)
        return {it["external_id"]: it for it in payload.get("items", [])}
    except Exception as exc:
        logger.warning("기존 데이터 로드 실패, 비우고 시작: %s", exc)
        return {}


def save(payload: dict[str, Any]) -> None:
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DOCS_DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=False)
    DATA_FILE.write_text(text, encoding="utf-8")
    DOCS_DATA_FILE.write_text(text, encoding="utf-8")


def collect(service_key: str, lookback: int = 10, lookahead: int = 60,
            research_only: bool = True) -> dict[str, Any]:
    """수집 → 정규화 → 학술연구용역 필터 → 키워드 매칭 → 누적 데이터 갱신 후 payload 반환."""
    keywords, min_score = _load_keywords()
    existing = load_existing()
    run_started = datetime.now(KST).isoformat()
    # 연도말(12/31) 갱신: 저장된 데이터의 연도가 올해와 다르면 누적을 초기화하고 새해부터 다시 축적.
    cur_year = datetime.now(KST).year
    try:
        _prev = json.loads(DATA_FILE.read_text(encoding="utf-8")) if DATA_FILE.exists() else {}
        _prev_year = int((_prev.get("generated_at") or "0000")[:4] or 0)
    except Exception:
        _prev_year = 0
    if _prev_year and _prev_year != cur_year:
        logger.info("연도 변경(%s→%s) 감지 → 누적 데이터 초기화(연말 갱신)", _prev_year, cur_year)
        existing = {}

    fetched = 0
    research_passed = 0
    new_count = 0
    new_items: list[dict[str, Any]] = []

    for ep in ENDPOINTS:
        for raw in _iter_pages(ep, service_key, lookback, lookahead):
            fetched += 1
            normalizer = NORMALIZERS[ep.src_type]
            item = normalizer(raw)
            if not item:
                continue
            if research_only and not _is_research(item):
                continue
            research_passed += 1
            score, matched = _score_item(item, keywords)
            if score < min_score:
                continue
            item["score"] = score
            item["matched_keywords"] = matched
            item["first_seen_at"] = existing.get(item["external_id"], {}).get("first_seen_at", run_started)
            item["last_seen_at"] = run_started
            is_new = item["external_id"] not in existing
            if is_new:
                new_count += 1
                new_items.append(item)
            existing[item["external_id"]] = item

    all_items = sorted(
        existing.values(),
        key=lambda x: (x.get("score", 0), x.get("last_seen_at") or ""),
        reverse=True,
    )

    payload = {
        "generated_at": run_started,
        "stats": {
            "total": len(all_items),
            "fetched_this_run": fetched,
            "research_passed_this_run": research_passed,
            "new_this_run": new_count,
            "min_score": min_score,
            "research_only": research_only,
        },
        "items": all_items,
    }
    save(payload)
    return {"payload": payload, "new_items": new_items}


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    service_key = os.environ.get("G2B_SERVICE_KEY", "").strip()
    if not service_key:
        logger.error("G2B_SERVICE_KEY 환경변수가 비어 있습니다.")
        return 1
    lookback = int(os.environ.get("LOOKBACK_DAYS", "10"))
    lookahead = int(os.environ.get("LOOKAHEAD_DAYS", "60"))
    research_only = os.environ.get("RESEARCH_ONLY", "1").strip() not in ("", "0", "false", "False")
    result = collect(service_key, lookback=lookback, lookahead=lookahead, research_only=research_only)
    stats = result["payload"]["stats"]
    logger.info("done: total=%d new=%d fetched=%d research_passed=%d",
                stats["total"], stats["new_this_run"], stats["fetched_this_run"],
                stats.get("research_passed_this_run", 0))
    # 다음 단계(kakao.py)가 읽을 수 있도록 신규 건만 별도 파일로
    new_path = ROOT / "data" / "new_items.json"
    new_path.write_text(
        json.dumps({"items": result["new_items"], "generated_at": stats},
                   ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
