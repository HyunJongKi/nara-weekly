"""docs/index.html 을 생성한다.
사이드 메뉴바 + Chart.js 시각화 + 엑셀 다운로드(SheetJS) + 임시 키워드 설정(localStorage).
정적 HTML 단일 파일. 데이터는 같은 폴더의 tenders.json 을 fetch 한다.
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "docs" / "index.html"
KEYWORDS_SRC = ROOT / "keywords.yml"
KEYWORDS_DST = ROOT / "docs" / "keywords.yml"

HTML = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Nara Radar — 농림업 학술연구용역 모니터</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {
    --bg: #f4f6fb; --surface: #ffffff; --text: #0f172a; --muted: #64748b;
    --border: #e2e8f0; --accent: #2563eb; --accent-soft: #dbeafe;
    --chip: #eff6ff; --chip-text: #1d4ed8;
    --pos: #16a34a; --neg: #dc2626; --warn: #d97706;
    --sidebar: #0f172a; --sidebar-text: #cbd5e1; --sidebar-active: #2563eb;
  }
  * { box-sizing: border-box; }
  html, body { margin:0; padding:0; height:100%; }
  body {
    font-family: -apple-system, "Segoe UI", "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
    background: var(--bg); color: var(--text);
    display: grid; grid-template-columns: 220px 1fr; grid-template-rows: 100vh;
  }
  aside.sidebar { background: var(--sidebar); color: var(--sidebar-text); padding: 18px 0; overflow-y: auto; display: flex; flex-direction: column; }
  .brand { padding: 4px 20px 16px; border-bottom: 1px solid #1e293b; }
  .brand h1 { margin:0; font-size:17px; color:white; }
  .brand .sub { color: #94a3b8; font-size:11.5px; margin-top:3px; }
  .nav { padding: 12px 0; flex: 1; }
  .nav button { width:100%; text-align:left; background:transparent; border:0; color: var(--sidebar-text); padding: 10px 20px; font-size:14px; cursor:pointer; display:flex; align-items:center; gap:10px; }
  .nav button:hover { background:#1e293b; color:white; }
  .nav button.active { background: var(--sidebar-active); color:white; border-left:3px solid #93c5fd; padding-left:17px; }
  .nav .icon { width:18px; display:inline-block; text-align:center; opacity:0.8; }
  .sidebar-foot { padding: 12px 20px; font-size:11.5px; color:#64748b; border-top:1px solid #1e293b; }
  .sidebar-foot a { color:#93c5fd; text-decoration:none; }

  main.main { overflow-y: auto; padding: 22px 28px 60px; }
  header.page-head { display:flex; align-items:flex-end; justify-content:space-between; gap:16px; flex-wrap:wrap; margin-bottom:14px; }
  header.page-head h2 { margin:0; font-size:22px; }
  header.page-head .meta { color: var(--muted); font-size:13px; }

  .filters { display:flex; flex-wrap:wrap; gap:8px; padding:12px 14px; background:var(--surface); border:1px solid var(--border); border-radius:12px; margin-bottom:14px; }
  .filters input, .filters select { padding:8px 10px; border:1px solid var(--border); border-radius:8px; font-size:13.5px; background:white; }
  .filters input[type=search] { flex:1; min-width:220px; }
  .filters .btn { border:1px solid var(--border); border-radius:8px; padding:8px 12px; cursor:pointer; font-size:13.5px; background:white; }
  .filters .btn:hover { background:#f1f5f9; }
  .filters .btn.primary { background: var(--accent); color:white; border-color: var(--accent); }
  .filters .btn.primary:hover { background:#1d4ed8; }
  .exp-toggle { display:flex; align-items:center; gap:5px; font-size:12.5px; color:#475569; cursor:pointer; white-space:nowrap; padding:0 4px; }
  .exp-toggle input { width:15px; height:15px; accent-color: var(--accent); cursor:pointer; }
  .list-note { font-size:12.5px; color:#475569; background:#f1f5f9; border:1px solid var(--border); border-radius:8px; padding:7px 11px; margin-bottom:10px; }
  .list-note b { color:#0f172a; }

  .kpis { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:12px; margin-bottom:14px; }
  .kpi { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:14px 16px; }
  .kpi .label { color: var(--muted); font-size:12.5px; margin-bottom:4px; }
  .kpi .value { font-size:24px; font-weight:700; letter-spacing:-0.5px; }
  .kpi .sub { color: var(--muted); font-size:11.5px; margin-top:3px; }

  .charts { display:grid; grid-template-columns: repeat(12, 1fr); gap:14px; }
  .panel { background: var(--surface); border:1px solid var(--border); border-radius:12px; padding: 14px 16px; }
  .panel h3 { margin: 0 0 10px; font-size:14.5px; color: var(--text); }
  .panel h3 .hint { color: var(--muted); font-size:11.5px; font-weight:normal; margin-left:6px; }
  .panel.col-6 { grid-column: span 6; }
  .panel.col-4 { grid-column: span 4; }
  .panel.col-8 { grid-column: span 8; }
  .panel.col-12 { grid-column: span 12; }
  .chart-box { position: relative; height: 240px; }
  .chart-box.tall { height: 320px; }

  .kwcloud { display:flex; flex-wrap:wrap; gap:6px 10px; align-items:baseline; padding: 8px 4px; }
  .kwcloud span { color: var(--accent); font-weight:600; }
  .kwcloud .c0 { color:#94a3b8; font-size:13px; }
  .kwcloud .c1 { color:#60a5fa; font-size:15px; }
  .kwcloud .c2 { color:#3b82f6; font-size:17px; }
  .kwcloud .c3 { color:#2563eb; font-size:20px; }
  .kwcloud .c4 { color:#1d4ed8; font-size:24px; font-weight:800; }

  .listwrap { display:flex; flex-direction:column; gap:10px; }
  .card { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:14px 16px; }
  .card h4 { margin:0 0 6px; font-size:15px; }
  .card-head { display:flex; align-items:flex-start; gap:12px; justify-content:space-between; margin-bottom:6px; }
  .card-title { flex:1; margin:0; font-size:15px; }
  .card-prob { display:flex; align-items:center; gap:8px; white-space:nowrap; }
  .card-prob b { font-size:14px; }
  .card-prob .pct-high { color:#16a34a; }
  .card-prob .pct-mid  { color:#d97706; }
  .card-prob .pct-low  { color:#64748b; }
  .chip.class-a { background:#dbeafe; color:#1d4ed8; font-weight:600; }
  .chip.class-b { background:#dcfce7; color:#166534; font-weight:600; }
  /* 찜하기 (즐겨찾기) */
  .card.favorite { background: linear-gradient(to right, #fef9c3, #fffbeb 40%); border: 2px solid #f59e0b; box-shadow: 0 2px 8px rgba(245, 158, 11, 0.15); }
  .card.favorite::before { content: "⭐ 찜한 공고"; position: absolute; top: -10px; left: 16px; background: #f59e0b; color: white; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 700; }
  .card { position: relative; }
  .fav-btn { background: none; border: 1.5px solid #cbd5e1; border-radius: 999px; padding: 4px 10px; cursor: pointer; font-size: 15px; color: #94a3b8; transition: all 0.15s; display:inline-flex; align-items:center; gap:4px; font-weight:600; }
  .fav-btn:hover { border-color: #f59e0b; color: #d97706; background: #fef9c3; }
  .fav-btn.active { border-color: #f59e0b; color: #d97706; background: #fef3c7; }
  .fav-btn.active:hover { background: #fef9c3; }
  /* 찜 필터 토글 */
  .fav-toggle { display:flex; align-items:center; gap:5px; font-size:12.5px; color:#78350f; cursor:pointer; white-space:nowrap; padding:5px 10px; background:#fef9c3; border:1px solid #fbbf24; border-radius:8px; font-weight:600; }
  .fav-toggle input { width:15px; height:15px; accent-color: #f59e0b; cursor:pointer; }
  /* 임시 키워드 필터바 인라인 */
  .tempkw-inline { display:flex; align-items:center; gap:5px; flex-wrap:wrap; padding:4px 8px; border:1px dashed #ca8a04; border-radius:8px; background:#fef9c3; }
  .tempkw-inline label { font-size:11px; color:#854d0e; font-weight:600; }
  .tempkw-inline input.tk-in { padding:5px 8px; border:1px solid #eab308; border-radius:6px; font-size:13px; width:120px; background:white; }
  .tempkw-inline button.tk-add { background:#eab308; color:white; border:0; border-radius:6px; padding:5px 9px; cursor:pointer; font-weight:600; font-size:12px; }
  .tempkw-inline button.tk-add:hover { background:#ca8a04; }
  .tempkw-inline .tk-chip { background:#fde68a; color:#854d0e; padding:3px 7px 3px 10px; border-radius:999px; font-size:12px; display:inline-flex; align-items:center; gap:4px; }
  .tempkw-inline .tk-chip button { background:none; border:0; color:#854d0e; font-weight:bold; cursor:pointer; padding:0 3px; font-size:14px; }
  .card .agency { color: var(--muted); font-size:12.5px; margin-bottom:6px; }
  .chips { display:flex; flex-wrap:wrap; gap:4px; margin:6px 0; }
  .chip { background:var(--chip); color:var(--chip-text); padding:2px 8px; border-radius:999px; font-size:11.5px; }
  .chip.score { background:#fef3c7; color:#92400e; }
  .chip.type { background:#dcfce7; color:#166534; }
  .chip.region { background:#f3e8ff; color:#7e22ce; }
  .chip.agency_type { background:#fee2e2; color:#b91c1c; }
  .chip.temp { background:#fef9c3; color:#854d0e; border:1px dashed #ca8a04; }
  .attachments { display:flex; flex-wrap:wrap; gap:6px; margin-top:6px; }
  .att-link { background:#eef2ff; color:#3730a3; padding:4px 9px; border-radius:6px; font-size:12px; text-decoration:none; border:1px solid #c7d2fe; }
  .att-link:hover { background:#e0e7ff; }
  details { margin-top:4px; }
  details summary { cursor:pointer; color: var(--accent); font-size:12.5px; }
  details pre { white-space:pre-wrap; word-break:break-word; background:#f1f5f9; padding:10px; border-radius:8px; font-size:12.5px; margin:6px 0 0; }

  /* 키워드 설정 */
  .kwset { display:grid; grid-template-columns: 1fr 1fr; gap:14px; }
  .kwset .panel { padding: 16px 18px; }
  .kwset input.kw-term { padding:8px 10px; border:1px solid var(--border); border-radius:8px; font-size:14px; min-width:140px; }
  .kwset input.kw-weight { padding:8px 10px; border:1px solid var(--border); border-radius:8px; font-size:14px; width:70px; }
  .kwset .row { display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-bottom:10px; }
  .kwset .chip-edit { background:#fef9c3; color:#854d0e; padding:4px 8px 4px 12px; border-radius:999px; font-size:13px; display:inline-flex; align-items:center; gap:6px; border:1px dashed #ca8a04; }
  .kwset .chip-edit button { background:none; border:0; color:#854d0e; font-weight:bold; cursor:pointer; font-size:14px; padding:0 4px; }
  .kwset pre.yml { background:#0f172a; color:#e2e8f0; padding:12px; border-radius:8px; font-size:12px; overflow:auto; max-height:300px; }
  .kwset .notice { background:#fef3c7; border:1px solid #fde68a; color:#854d0e; padding:10px 12px; border-radius:8px; font-size:13px; margin-bottom:10px; line-height:1.6; }
  .kwset label { font-size:12.5px; color: var(--muted); display:block; margin: 6px 0 4px; }
  .kwset .repo-input { width:100%; padding:8px 10px; border:1px solid var(--border); border-radius:8px; font-size:13px; font-family: monospace; }
  .kwset a.btn-link { display:inline-block; background: var(--accent); color:white; padding:8px 14px; border-radius:8px; text-decoration:none; font-size:13.5px; margin-top:8px; }
  .kwset a.btn-link:hover { background:#1d4ed8; }
  .kwset .new-yml { background:#ecfdf5; border:1px solid #a7f3d0; color:#065f46; padding:10px 12px; border-radius:8px; font-family: monospace; font-size:12.5px; margin-top:8px; white-space: pre-wrap; }

  /* INK 수주 가능성 테이블 */
  .ink-table-wrap { overflow-x:auto; }
  .ink-table { width:100%; border-collapse:collapse; font-size:13px; }
  .ink-table th, .ink-table td { padding:8px 10px; border-bottom:1px solid var(--border); text-align:left; vertical-align:top; }
  .ink-table th { background:#f8fafc; color:var(--muted); font-weight:600; font-size:12px; position:sticky; top:0; }
  .ink-table .rank { font-weight:800; color:var(--accent); text-align:center; width:38px; }
  .ink-table .rank .medal { font-size:15px; }
  .ink-table .title { font-weight:600; max-width:340px; }
  .ink-table .title a { color:var(--text); }
  .ink-table small { color:var(--muted); }
  .ink-table .prob { white-space:nowrap; min-width:130px; }
  .ink-table .prob b { margin-left:6px; font-size:13.5px; }
  .ink-table .reasons { color:var(--muted); font-size:12px; max-width:260px; }
  .ink-bar { display:inline-block; width:72px; height:9px; background:#e2e8f0; border-radius:999px; overflow:hidden; vertical-align:middle; }
  .ink-bar span { display:block; height:100%; border-radius:999px; }
  .ink-bar.high span { background:#16a34a; }
  .ink-bar.mid  span { background:#d97706; }
  .ink-bar.low  span { background:#94a3b8; }
  .ink-note { color:var(--muted); font-size:11.5px; margin-top:8px; line-height:1.6; }
  .empty { color: var(--muted); text-align:center; padding: 40px; }
  .footnote { color: var(--muted); font-size:11.5px; padding: 10px 4px; }

  @media (max-width: 920px) {
    body { grid-template-columns: 60px 1fr; }
    .brand h1, .brand .sub, .nav button span.label, .sidebar-foot { display:none; }
    .nav button { justify-content:center; padding: 12px 0; }
    .nav button.active { border-left:0; padding: 12px 0; }
    .kwset { grid-template-columns: 1fr; }
  }
  @media (max-width: 700px) {
    .panel.col-4, .panel.col-6, .panel.col-8 { grid-column: span 12; }
  }
</style>
</head>
<body>
  <aside class="sidebar">
    <div class="brand">
      <h1>🌾 Nara Radar</h1>
      <div class="sub">농림업 학술연구용역 모니터</div>
    </div>
    <nav class="nav">
      <button class="active" data-view="dashboard"><span class="icon">📊</span><span class="label">대시보드</span></button>
      <button data-view="list"><span class="icon">📋</span><span class="label">공고 목록</span></button>
      <button data-view="keyword"><span class="icon">🧭</span><span class="label">사업 분류 분석</span></button>
      <button data-view="region"><span class="icon">🗺️</span><span class="label">지역 분석</span></button>
      <button data-view="settings"><span class="icon">⚙️</span><span class="label">키워드 설정</span></button>
      <button data-view="about"><span class="icon">ℹ️</span><span class="label">설명</span></button>
    </nav>
    <div class="sidebar-foot">
      <div id="generated">로딩 중…</div>
      <div style="margin-top:6px"><a href="https://www.g2b.go.kr" target="_blank">나라장터 ↗</a></div>
    </div>
  </aside>

  <main class="main">
    <header class="page-head">
      <div>
        <h2 id="page-title">대시보드</h2>
        <div class="meta" id="page-sub">검색어 매칭 발주계획·사전규격 통합 통계</div>
      </div>
      <div class="meta" id="page-stat"></div>
    </header>

    <div class="filters" id="filter-bar">
      <input type="search" id="q" placeholder="제목·기관·내용 검색…">
      <select id="type">
        <option value="">전체 유형</option>
        <option value="order_plan">발주계획</option>
        <option value="pre_spec">사전규격</option>
      </select>
      <select id="kw"><option value="">전체 키워드</option></select>
      <select id="region"><option value="">전체 지역</option></select>
      <select id="agency_type"><option value="">전체 발주처</option></select>
      <select id="sort">
        <option value="recent">최신순</option>
        <option value="score">점수순</option>
        <option value="budget">예산순</option>
      </select>
      <label class="exp-toggle" title="입찰공고기간이 경과(마감)했거나 최근 수집창에서 사라진 공고를 목록에 포함합니다. (차트는 1년 누적 전체 기준)"><input type="checkbox" id="showExpired"> 경과 공고 포함</label>
      <label class="fav-toggle" title="⭐ 찜한 공고만 목록에 표시합니다."><input type="checkbox" id="favOnly"> ⭐ 찜한 것만</label>
      <button class="btn" id="reset">초기화</button>
      <button class="btn primary" id="xlsx">⬇ 엑셀 다운로드</button>
      <!-- 임시 키워드 인라인 (본인 브라우저에만 저장) -->
      <div class="tempkw-inline" title="본인 브라우저에만 저장되는 임시 키워드. 즉시 차트·필터에 반영됩니다.">
        <label>+ 임시 KW</label>
        <input type="text" id="tk-in" class="tk-in" placeholder="예: 축산" />
        <button class="tk-add" id="tk-add">추가</button>
        <span id="tk-list"></span>
      </div>
    </div>

    <!-- 대시보드 -->
    <section data-view-pane="dashboard">
      <div class="kpis" id="kpis"></div>
      <div class="charts">
        <div class="panel col-8">
          <h3>일자별 공고건수 추이 <span class="hint">최근 갱신일 기준</span></h3>
          <div class="chart-box tall"><canvas id="ch-timeline"></canvas></div>
        </div>
        <div class="panel col-4">
          <h3>공고유형 비중</h3>
          <div class="chart-box tall"><canvas id="ch-type"></canvas></div>
        </div>
        <div class="panel col-6">
          <h3>금액규모별 분포 <span class="hint">예산금액 구간</span></h3>
          <div class="chart-box"><canvas id="ch-budget"></canvas></div>
        </div>
        <div class="panel col-6">
          <h3>발주처 유형별</h3>
          <div class="chart-box"><canvas id="ch-agency-type"></canvas></div>
        </div>
        <div class="panel col-6">
          <h3>지역별 공고건수</h3>
          <div class="chart-box"><canvas id="ch-region"></canvas></div>
        </div>
        <div class="panel col-6">
          <h3>발주기관 상위 10</h3>
          <div class="chart-box"><canvas id="ch-agency"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>키워드별 언급 빈도</h3>
          <div class="chart-box"><canvas id="ch-keyword"></canvas></div>
          <div class="kwcloud" id="kwcloud"></div>
        </div>
        <div class="panel col-12">
          <h3>🎯 아이엔케이(주) 수주 가능성 <span class="hint">추정 모델 · 최신 공고일자순 정렬 (동점 시 가능성순)</span></h3>
          <div class="ink-table-wrap" id="ink-ranking"></div>
          <div class="ink-note">
            ※ 아이엔케이(주) <b>실적 프로파일</b> 반영 모델 — 주력분야(농업·식량 정책/제도 + 스마트팜·원예 R&D),
            적정 규모(<b>1억 미만 소형</b>), 주력 고객(공공기관·지자체).
            <b>분야 적합도(30)</b> + <b>용역 성격(28)</b> + <b>예산 규모(22·소형 선호)</b> + <b>발주처 친화도(20)</b> 가중 합산.
            <b>추정치</b>이며 실제 입찰 경쟁·평가위원 성향·컨소시엄 구성에 따라 달라집니다.
            상단 필터를 적용하면 순위도 함께 갱신됩니다.
          </div>
        </div>
      </div>
      <div class="footnote">※ 본 통계는 나라장터에 게시된 발주계획·사전규격 공고를 기반으로 분석되었으며, 실제 계약과는 차이가 있을 수 있습니다.</div>
    </section>

    <!-- 목록 -->
    <section data-view-pane="list" style="display:none">
      <div id="list" class="listwrap"><div class="empty">데이터 로딩 중…</div></div>
    </section>

    <!-- 사업 분류 분석 (3축: 성격 A × 분야 B × 발주주체 C) -->
    <section data-view-pane="keyword" style="display:none">
      <div class="charts">
        <div class="panel col-6">
          <h3>🎯 사업 성격 (A축)  <span class="hint">정책·계획·타당성·조사·평가·위탁</span></h3>
          <div class="chart-box tall"><canvas id="ch-classA"></canvas></div>
        </div>
        <div class="panel col-6">
          <h3>🌾 대상 분야 (B축)  <span class="hint">농업·농촌·축산·스마트·산림·수자원</span></h3>
          <div class="chart-box tall"><canvas id="ch-classB"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>🔀 성격 × 분야 히트맵  <span class="hint">교차 분석: 어떤 분야에서 어떤 성격 사업이 많은지</span></h3>
          <div class="chart-box tall"><canvas id="ch-classAB"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>🏛️ 발주주체 × 분야 (C×B 스택바)  <span class="hint">누가 어떤 분야를 발주하는지</span></h3>
          <div class="chart-box tall"><canvas id="ch-classCB"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>🔍 (참고) 원 키워드별 매칭 빈도</h3>
          <div class="chart-box"><canvas id="ch-keyword2"></canvas></div>
        </div>
      </div>
    </section>

    <!-- 지역 분석 (지역 × 분야·성격) -->
    <section data-view-pane="region" style="display:none">
      <div class="charts">
        <div class="panel col-12">
          <h3>🗺️ 지역별 공고건수</h3>
          <div class="chart-box tall"><canvas id="ch-region2"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>🌾 지역 × 분야(B) 스택 바  <span class="hint">각 지역에서 어떤 분야가 강한지</span></h3>
          <div class="chart-box tall"><canvas id="ch-region-domain"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>🎯 지역 × 사업 성격(A) 스택 바  <span class="hint">각 지역에서 어떤 성격 사업이 많은지</span></h3>
          <div class="chart-box tall"><canvas id="ch-region-nature"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>공고유형(발주계획·사전규격) 지역 분포</h3>
          <div class="chart-box"><canvas id="ch-region-type"></canvas></div>
        </div>
      </div>
    </section>

    <!-- 키워드 설정 -->
    <section data-view-pane="settings" style="display:none">
      <div class="kwset">
        <div class="panel">
          <h3>임시 키워드 추가 <span class="hint">내 브라우저에만 저장 (localStorage)</span></h3>
          <div class="notice">
            이 키워드는 <b>본인 브라우저</b>에만 저장되며 즉시 대시보드 필터·차트에 반영됩니다.
            카카오톡 알림이나 다른 사람의 대시보드에는 영향이 없습니다.
            영구적으로 추가하려면 오른쪽 패널에서 <b>keywords.yml</b>을 편집하세요.
          </div>
          <div class="row">
            <input class="kw-term"   id="new-term"   placeholder="새 키워드 (예: 축산)" />
            <input class="kw-weight" id="new-weight" type="number" min="1" max="10" value="5" />
            <button class="btn primary" id="add-kw">＋ 추가</button>
            <button class="btn" id="clear-kw">모두 초기화</button>
          </div>
          <div id="temp-kw-list" style="margin-top:10px"></div>
          <label style="margin-top:14px">📁 현재 keywords.yml (서버 기본값, 카톡 알림 기준)</label>
          <pre class="yml" id="server-yml">로딩 중…</pre>
        </div>

        <div class="panel">
          <h3>영구 반영 (keywords.yml 직접 편집)</h3>
          <div class="notice">
            <b>keywords.yml</b> 을 수정하면 다음 GitHub Actions 실행(매주 수·금 10시)부터
            카톡 알림과 모든 사용자의 대시보드에 반영됩니다.
          </div>
          <label>본인 GitHub repo URL (예: <code>https://github.com/your-id/nara-weekly</code>)</label>
          <input class="repo-input" id="repo-url" placeholder="https://github.com/your-id/nara-weekly" />
          <a id="github-edit" class="btn-link" target="_blank">✏ GitHub 에서 keywords.yml 편집 ↗</a>

          <label style="margin-top:16px">📌 추가/삭제용 YAML 한 줄 자동 생성</label>
          <div class="row" style="margin-top:4px">
            <input class="kw-term"   id="yml-term"   placeholder="키워드" />
            <input class="kw-weight" id="yml-weight" type="number" min="1" max="10" value="5" />
            <input class="kw-term"   id="yml-group"  placeholder="그룹 (예: agriculture)" value="agriculture" />
          </div>
          <div class="new-yml" id="new-yml-line"># 키워드와 그룹을 입력하면 여기에 자동 생성됩니다.</div>

          <label style="margin-top:16px">📒 사용법 요약</label>
          <ol style="font-size:13px; line-height:1.7; padding-left:18px; color:#334155;">
            <li>위 박스에서 새 키워드 한 줄을 복사.</li>
            <li>"GitHub 에서 편집" 버튼으로 keywords.yml 을 GitHub 웹에서 열기.</li>
            <li><code>keywords:</code> 아래에 한 줄 붙여넣고 commit.</li>
            <li>다음 수·금 10시 자동 실행 후 카톡과 대시보드에 반영. 즉시 보고 싶으면 Actions → Run workflow.</li>
          </ol>
        </div>
      </div>
    </section>

    <!-- 설명 -->
    <section data-view-pane="about" style="display:none">
      <div class="panel col-12" style="grid-column: span 12;">
        <h3>이 대시보드는?</h3>
        <p style="font-size:14px; line-height:1.7;">
          나라장터(조달청) <b>발주계획</b> 및 <b>사전규격</b> 공고 중 회사 관심 검색어
          (<b>농업·농촌·식량·원예·스마트팜·노지 스마트·산림·임업</b>) 에 매칭되는 <b>학술연구용역</b> 만
          자동으로 수집·정리한 페이지입니다.
        </p>
        <p style="font-size:14px; line-height:1.7;">
          GitHub Actions cron 으로 <b>매주 수요일·금요일 오전 10시(KST)</b> 에 자동 실행되며,
          신규 공고는 <b>카카오톡 '나에게 보내기'</b> 로도 함께 발송됩니다.
        </p>
        <p style="font-size:14px; line-height:1.7; color: var(--muted);">
          데이터 출처 · <a href="https://www.data.go.kr" target="_blank">공공데이터포털</a> /
          <a href="https://www.g2b.go.kr" target="_blank">나라장터</a>.
          호스팅 · GitHub Pages (무료).
        </p>
      </div>
    </section>
  </main>

<script>
const TYPE_LABEL = { order_plan: "발주계획", pre_spec: "사전규격" };
const BUDGET_BUCKETS = [
  { label: "5천만 미만", max: 5e7 },
  { label: "0.5~1억",   max: 1e8 },
  { label: "1~2억",     max: 2e8 },
  { label: "2~3억",     max: 3e8 },
  { label: "3~4억",     max: 4e8 },
  { label: "4~5억",     max: 5e8 },
  { label: "5억 이상",  max: Infinity },
];
const COLORS = ["#2563eb","#16a34a","#dc2626","#d97706","#7e22ce","#0891b2","#db2777","#65a30d","#475569","#0ea5e9","#a16207","#be123c"];
const LS_TEMP_KW = "nw_temp_kw_v1";
const LS_REPO_URL = "nw_repo_url_v1";

let ALL = [];            // 원본
let AUGMENTED = [];      // 임시 키워드 적용본
let CHARTS = {};
let CURRENT_VIEW = "dashboard";
let TEMP_KW = JSON.parse(localStorage.getItem(LS_TEMP_KW) || "[]");

// ── 찜하기 (즐겨찾기) — external_id 기준으로 localStorage 저장 ──
const LS_FAVS = "nw_favorites_v1";
let FAVORITES = new Set(JSON.parse(localStorage.getItem(LS_FAVS) || "[]"));
function isFav(it) { return FAVORITES.has(it.external_id); }
function toggleFav(extId) {
  if (FAVORITES.has(extId)) FAVORITES.delete(extId);
  else FAVORITES.add(extId);
  localStorage.setItem(LS_FAVS, JSON.stringify([...FAVORITES]));
}

// ── 유틸 ──
function fmtMoney(n) {
  if (typeof n !== "number" || !isFinite(n)) return "";
  if (n >= 1e8) return (n/1e8).toFixed(1).replace(/\.0$/,"") + "억원";
  return n.toLocaleString("ko-KR") + "원";
}
function fmtDate(s) { if (!s) return ""; try { return s.replace("T"," ").slice(0,16); } catch { return s; } }
function bucketOf(amount) {
  if (typeof amount !== "number" || amount <= 0) return null;
  for (const b of BUDGET_BUCKETS) if (amount < b.max) return b.label;
  return BUDGET_BUCKETS[BUDGET_BUCKETS.length-1].label;
}
function destroy(name) { if (CHARTS[name]) { CHARTS[name].destroy(); delete CHARTS[name]; } }

// ── 중복 제거: 같은 사업이 발주계획+사전규격에 동시에 있으면 사전규격(pre_spec) 우선 ──
function dedupe(items) {
  const norm = s => (s || "").toLowerCase().replace(/[\s()\[\]\-_,./·:、，]/g, "");
  const map = new Map();
  for (const it of items) {
    const key = norm(it.title) || it.external_id;
    const cur = map.get(key);
    if (!cur) { map.set(key, it); continue; }
    // 이미 있는 항목과 충돌: 사전규격을 우선 채택
    if (it.source_type === "pre_spec" && cur.source_type !== "pre_spec") {
      // first_seen_at 은 더 이른 값 보존
      const merged = {...it};
      if (cur.first_seen_at && (!merged.first_seen_at || cur.first_seen_at < merged.first_seen_at))
        merged.first_seen_at = cur.first_seen_at;
      map.set(key, merged);
    }
    // 둘 다 같은 유형이거나 기존이 이미 사전규격이면 기존 유지
  }
  return [...map.values()];
}

// ── 임시 키워드 적용 ──
function applyTempKeywords(items) {
  if (!TEMP_KW.length) return items.map(x => ({...x, _temp_matched: []}));
  return items.map(it => {
    const hay = [it.title, it.description, it.agency, it.agency_dept].filter(Boolean).join(" ").toLowerCase();
    const tempMatched = [];
    let bonus = 0;
    for (const k of TEMP_KW) {
      const t = k.term.toLowerCase();
      if (t && hay.includes(t) && !(it.matched_keywords||[]).includes(k.term)) {
        tempMatched.push(k.term);
        bonus += Number(k.weight) || 5;
      }
    }
    return {
      ...it,
      matched_keywords: [...(it.matched_keywords||[]), ...tempMatched],
      score: (it.score||0) + bonus,
      _temp_matched: tempMatched,
    };
  });
}
function rebuildAugmented() {
  AUGMENTED = applyTempKeywords(ALL);
  // 키워드 셀렉터도 갱신
  const sel = document.getElementById("kw");
  const cur = sel.value;
  const kwSet = new Set();
  AUGMENTED.forEach(it => (it.matched_keywords||[]).forEach(k => kwSet.add(k)));
  sel.innerHTML = '<option value="">전체 키워드</option>' +
    [...kwSet].sort().map(k => `<option value="${k}">${k}</option>`).join("");
  if (kwSet.has(cur)) sel.value = cur;
}

// ── 필터링 ──
function getFilters() {
  return {
    q: document.getElementById("q").value.trim().toLowerCase(),
    t: document.getElementById("type").value,
    k: document.getElementById("kw").value,
    r: document.getElementById("region").value,
    at: document.getElementById("agency_type").value,
    sort: document.getElementById("sort").value,
  };
}
// 입찰공고기간 경과 판정: 명시 마감(사전규격 의견제출마감)이 지났거나,
// 최근 수집창에서 14일 이상 사라진(=공고기간 종료) 건. 목록에서만 기본 숨김(차트는 1년 누적 전체).
function isExpired(it) {
  // deadline(사전규격 의견제출 마감일) 이 실제로 지난 것만 만료로 판정.
  // last_seen_at 기준(수집창에서 사라짐)은 폐기 — workflow lookback 이 7일이라
  // 옛 항목이 재수집 안 될 뿐 실제로는 아직 진행 중일 수 있어 과도한 필터였음.
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (it.deadline) {
    const d = new Date(it.deadline);
    if (!isNaN(d) && d < today) return true;
  }
  return false;
}
function applyFilter(items, f) {
  return items.filter(it => {
    if (f.t && it.source_type !== f.t) return false;
    if (f.k && !(it.matched_keywords || []).includes(f.k)) return false;
    if (f.r && (it.region||"중앙/전국") !== f.r) return false;
    if (f.at && (it.agency_type||"기타") !== f.at) return false;
    if (f.q) {
      const hay = [it.title, it.agency, it.agency_dept, it.description].filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(f.q)) return false;
    }
    return true;
  });
}

// ── 집계 ──
function aggregate(items) {
  const byDate = {}, byType = {order_plan:0, pre_spec:0}, byBudget = {}, byRegion = {}, byAgency = {}, byAgencyType = {}, byKw = {};
  const byKwType = {}, byRegionType = {};
  let budgetSum = 0, budgetCnt = 0;
  for (const it of items) {
    const d = (it.order_planned_date || it.last_seen_at || "").slice(0,10);
    if (d) byDate[d] = (byDate[d]||0) + 1;
    byType[it.source_type] = (byType[it.source_type]||0) + 1;
    const b = bucketOf(it.budget_amount);
    if (b) byBudget[b] = (byBudget[b]||0) + 1;
    const r = it.region || "중앙/전국";
    byRegion[r] = (byRegion[r]||0) + 1;
    if (!byRegionType[r]) byRegionType[r] = {order_plan:0, pre_spec:0};
    byRegionType[r][it.source_type] = (byRegionType[r][it.source_type]||0) + 1;
    if (it.agency) byAgency[it.agency] = (byAgency[it.agency]||0) + 1;
    const at = it.agency_type || "기타";
    byAgencyType[at] = (byAgencyType[at]||0) + 1;
    (it.matched_keywords||[]).forEach(k => {
      byKw[k] = (byKw[k]||0) + 1;
      if (!byKwType[k]) byKwType[k] = {order_plan:0, pre_spec:0};
      byKwType[k][it.source_type] = (byKwType[k][it.source_type]||0) + 1;
    });
    if (typeof it.budget_amount === "number" && it.budget_amount > 0) { budgetSum += it.budget_amount; budgetCnt += 1; }
  }
  return { byDate, byType, byBudget, byRegion, byAgency, byAgencyType, byKw, byKwType, byRegionType,
           budgetSum, budgetAvg: budgetCnt ? budgetSum/budgetCnt : 0 };
}

// ── 차트 헬퍼 ──
function barChart(canvasId, labels, data, label, opts={}) {
  destroy(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  CHARTS[canvasId] = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ label, data, backgroundColor: opts.colors || "#2563eb" }] },
    options: { responsive:true, maintainAspectRatio:false, indexAxis: opts.horizontal?"y":"x",
      plugins:{ legend:{ display:false } },
      scales:{ x:{ beginAtZero:true, ticks:{ precision:0 } }, y:{ beginAtZero:true, ticks:{ precision:0 } } } },
  });
}
function doughnutChart(canvasId, labels, data) {
  destroy(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  CHARTS[canvasId] = new Chart(ctx, {
    type: "doughnut",
    data: { labels, datasets: [{ data, backgroundColor: COLORS }] },
    options: { responsive:true, maintainAspectRatio:false, plugins:{ legend:{ position:"bottom" } } },
  });
}
function stackedBar(canvasId, labels, datasets, horizontal=false) {
  destroy(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  CHARTS[canvasId] = new Chart(ctx, { type:"bar", data:{ labels, datasets },
    options:{ responsive:true, maintainAspectRatio:false, indexAxis: horizontal?"y":"x",
      plugins:{ legend:{ position:"bottom" } },
      scales:{ x:{ stacked:true, beginAtZero:true, ticks:{ precision:0 } }, y:{ stacked:true, beginAtZero:true, ticks:{ precision:0 } } } } });
}
function timelineChart(canvasId, items) {
  destroy(canvasId);
  const op = {}, ps = {}, tot = {};
  for (const it of items) {
    const d = (it.order_planned_date || it.last_seen_at || "").slice(0,10);
    if (!d) continue;
    tot[d] = (tot[d]||0) + 1;
    if (it.source_type === "order_plan") op[d] = (op[d]||0) + 1;
    else if (it.source_type === "pre_spec") ps[d] = (ps[d]||0) + 1;
  }
  const dates = Object.keys(tot).sort();
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  CHARTS[canvasId] = new Chart(ctx, { type:"line",
    data:{ labels: dates, datasets:[
      { label:"발주계획", data: dates.map(d=>op[d]||0), borderColor:"#2563eb", backgroundColor:"rgba(37,99,235,0.15)", fill:true, tension:0.25 },
      { label:"사전규격", data: dates.map(d=>ps[d]||0), borderColor:"#16a34a", backgroundColor:"rgba(22,163,74,0.15)", fill:true, tension:0.25 },
      { label:"전체",     data: dates.map(d=>tot[d]||0), borderColor:"#0f172a", borderDash:[4,4], fill:false, tension:0.25 },
    ]}, options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ position:"bottom" } }, scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } } } } });
}

// ── 사업 자동 분류 (3축: 성격 A × 분야 B × 발주주체 C) ──
// 사용자 요청: A6(R&D), A7(설계·규격), A9(안전·환경조사) 는 수집단계 제외라 분류에도 없음
const CLASSIFY_A = [
  { code:"A1", label:"정책·제도 연구",  re:/정책\s*연구|제도\s*개선|법제|개정|개선\s*방안|제도\s*수립|규제/ },
  { code:"A2", label:"계획 수립",       re:/기본\s*계획|종합\s*계획|마스터플랜|중장기|시행\s*계획|기본\s*구상|전략\s*수립|재생\s*기본/ },
  { code:"A3", label:"타당성 조사",     re:/타당성|예비\s*타당성|사업성|기본구상/ },
  { code:"A4", label:"실태·현황 조사",   re:/실태\s*조사|현황\s*조사|통계\s*조사|기초\s*조사|수요\s*조사|기술\s*수요|수요\s*분석/ },
  { code:"A5", label:"성과·평가",       re:/성과\s*평가|영향\s*평가|효과\s*분석|모니터링|진단|성과분석|평가\s*연구/ },
  { code:"A8", label:"위탁·지원",       re:/위탁\s*조사|위탁\s*운영|지원\s*사업|컨설팅|자문|용역\s*지원/ },
];
const CLASSIFY_B = [
  { code:"B1", label:"농업 생산·기술",   re:/재배|영농|기술\s*보급|농기계|병해충|경작|농법/ },
  { code:"B2", label:"농촌 공간·정비",   re:/농촌\s*공간|정비\s*사업|재구조화|재생|중심지\s*활성화|농촌\s*마을|공간\s*재구조화/ },
  { code:"B3", label:"축산·경영체",     re:/축산|낙농|젖소|경영체|양돈|양계|가축/ },
  { code:"B4", label:"스마트농업",      re:/스마트팜|노지\s*스마트|노지스마트|IoT|빅데이터|자동화|디지털\s*농업/ },
  { code:"B5", label:"식량·안보",       re:/식량\s*안보|식량\s*수급|자급률|식량\s*정책/ },
  { code:"B6", label:"원예·특작",       re:/원예|화훼|인삼|흑삼|채소|과수|특작|약용/ },
  { code:"B7", label:"산림 자원·경영",  re:/산림\s*자원|임업|산림\s*복지|탄소\s*흡수|숲가꾸기|입목|산림\s*경영|목재/ },
  { code:"B8", label:"산림 시설·환경",  re:/임도|산사태|산림\s*재해|재선충|방제|산림\s*안전|산림\s*환경/ },
  { code:"B9", label:"농업기반·수자원", re:/저수지|관정|지하수|수리\s*시설|간척지|용수|배수/ },
];
const CLASSIFY_C = [  // agency_type + agency 조합
  { code:"C1", label:"중앙부처",       match: it => (it.agency_type||"") === "중앙부처" },
  { code:"C2", label:"공공기관",       match: it => (it.agency_type||"") === "공공기관" },
  { code:"C3", label:"광역지자체",     match: it => (it.agency_type||"") === "지자체" && /특별시|광역시|특별자치도|특별자치시|^[가-힣]{2}도\b/.test(it.agency||"") },
  { code:"C4", label:"기초지자체",     match: it => (it.agency_type||"") === "지자체" && !/특별시|광역시|특별자치도|특별자치시|^[가-힣]{2}도\b/.test(it.agency||"") },
  { code:"C5", label:"교육·연구기관",  match: it => (it.agency_type||"") === "교육기관" },
];
const CLASSIFY_A_COLORS = {A1:"#2563eb",A2:"#0891b2",A3:"#7e22ce",A4:"#16a34a",A5:"#d97706",A8:"#64748b",A0:"#cbd5e1"};
const CLASSIFY_B_COLORS = {B1:"#059669",B2:"#0284c7",B3:"#dc2626",B4:"#7c3aed",B5:"#ca8a04",B6:"#db2777",B7:"#65a30d",B8:"#78350f",B9:"#0d9488",B0:"#cbd5e1"};

function classifyA(it) {
  const t = (it.title||"") + " " + (it.description||"");
  for (const c of CLASSIFY_A) if (c.re.test(t)) return c;
  return { code:"A0", label:"기타" };
}
function classifyB(it) {
  const t = (it.title||"") + " " + (it.description||"");
  for (const c of CLASSIFY_B) if (c.re.test(t)) return c;
  return { code:"B0", label:"기타" };
}
function classifyC(it) {
  for (const c of CLASSIFY_C) if (c.match(it)) return c;
  return { code:"C0", label:"기타" };
}

// ── 아이엔케이(주) 수주 가능성 추정 (실적 프로파일 반영) ──
// INK 프로파일: ①주력=농업·식량 정책/제도 + 스마트팜·원예 기술R&D  ②적정규모=1억 미만 소형
//               ③주력고객=공공기관·광역/기초 지자체
// 배점: 분야 적합도(30) + 용역 성격(28) + 예산 규모(22, 소형 선호) + 발주처 친화도(20)
const INK_CORE_KW = ["농업","식량","농촌","스마트팜","노지 스마트","노지스마트","원예"]; // 주력 키워드
function winProbability(it) {
  const text = [it.title, it.description, it.bsns_div].filter(Boolean).join(" ");
  const reasons = [];

  // 1) 분야 적합도 (30) — 키워드 가중치 기본 + INK 주력 키워드 보너스
  const mk = it.matched_keywords || [];
  const coreHits = mk.filter(k => INK_CORE_KW.includes(k)).length;
  const baseFit = Math.min(20, Math.round((it.score || 0) * 2));
  const coreBonus = Math.min(10, coreHits * 5);
  const fit = Math.min(30, baseFit + coreBonus);
  if (mk.length) reasons.push((coreHits ? "주력분야 " : "분야 ") + mk.slice(0,3).join("·"));

  // 2) 용역 성격 (28) — INK 강점: 정책·제도 연구 + 기술 R&D
  let nature = 10, natLabel = "일반 용역";
  if (/정책\s*연구|제도\s*개선|기본\s*계획|마스터플랜|중장기|종합\s*계획|전략\s*수립|타당성|개선\s*방안|개정\s*방향/.test(text)) { nature = 28; natLabel = "정책·제도·계획"; }
  else if (/학술|연구\s*용역|기초\s*연구|기획\s*연구|R&D|연구개발|개발\s*연구|마커|육종|품종/.test(text))                       { nature = 25; natLabel = "기술·연구개발"; }
  else if (/실태\s*조사|조사\s*연구|모니터링|분석|진단|평가/.test(text))                                                       { nature = 15; natLabel = "조사·분석"; }
  else if (/위탁\s*조사|안전\s*점검|점검|측량|이행|단순/.test(text))                                                           { nature = 6;  natLabel = "단순·위탁"; }
  reasons.push(natLabel);

  // 3) 예산 규모 (22) — INK는 1억 미만 소형이 스위트스팟 (클수록 감점)
  let bScore = 12, bLabel = "예산 미공개";
  const b = it.budget_amount;
  if (typeof b === "number" && b > 0) {
    if (b < 3e7)            { bScore = 16; bLabel = "초소형(<3천만)"; }
    else if (b < 1e8)       { bScore = 22; bLabel = "소형(3천만~1억) 적정"; }
    else if (b < 3e8)       { bScore = 16; bLabel = "중형(1~3억)"; }
    else if (b < 5e8)       { bScore = 10; bLabel = "중대형(3~5억)"; }
    else                    { bScore = 5;  bLabel = "대형(5억~)"; }
  }
  reasons.push(bLabel);

  // 4) 발주처 친화도 (20) — INK 주력: 공공기관·지자체(광역/기초)
  let aScore = 7, aLabel = "";
  const at = it.agency_type || "";
  if (at === "공공기관")      { aScore = 20; aLabel = "공공기관"; }
  else if (at === "지자체")   { aScore = 20; aLabel = "지자체"; }
  else if (at === "중앙부처") { aScore = 13; aLabel = "중앙부처"; }
  else if (at === "교육기관") { aScore = 9;  aLabel = "교육기관"; }
  else                        { aScore = 7;  aLabel = "기타"; }

  let pct = fit + nature + bScore + aScore;
  pct = Math.max(5, Math.min(98, pct));
  return { pct, reasons };
}

function renderInkRanking(items) {
  const el = document.getElementById("ink-ranking");
  if (!el) return;
  // 최신 공고일자 우선, 동점일 땐 수주 가능성(%)이 높은 순으로 tiebreak
  const ranked = items.map(it => ({ it, ...winProbability(it) })).sort((a,b) => {
    const ka = a.it.order_planned_date || a.it.last_seen_at || "";
    const kb = b.it.order_planned_date || b.it.last_seen_at || "";
    const dateCmp = kb.localeCompare(ka);
    if (dateCmp !== 0) return dateCmp;
    return b.pct - a.pct;
  });
  if (!ranked.length) { el.innerHTML = '<div class="empty">조건에 맞는 공고가 없습니다.</div>'; return; }
  const medal = i => i===0?"🥇":i===1?"🥈":i===2?"🥉":(i+1);
  el.innerHTML = `<table class="ink-table">
    <thead><tr><th>순위</th><th>공고명</th><th>발주처</th><th>예산</th><th>수주 가능성</th><th>주요 근거</th></tr></thead>
    <tbody>` + ranked.map((r,i) => {
      const grade = r.pct>=75?"high":r.pct>=55?"mid":"low";
      const titleHtml = r.it.url ? `<a href="${r.it.url}" target="_blank" rel="noopener">${r.it.title||""}</a>` : (r.it.title||"");
      return `<tr>
        <td class="rank"><span class="medal">${medal(i)}</span></td>
        <td class="title">${titleHtml}<br><small>${TYPE_LABEL[r.it.source_type]||""}</small></td>
        <td>${r.it.agency||""}<br><small>${r.it.agency_type||""} · ${r.it.region||""}</small></td>
        <td>${fmtMoney(r.it.budget_amount)||"-"}</td>
        <td class="prob"><span class="ink-bar ${grade}"><span style="width:${r.pct}%"></span></span><b>${r.pct}%</b></td>
        <td class="reasons">${r.reasons.join(" · ")}</td>
      </tr>`;
    }).join("") + `</tbody></table>`;
}

function renderKwCloud(byKw) {
  const el = document.getElementById("kwcloud"); if (!el) return;
  const entries = Object.entries(byKw).sort((a,b)=>b[1]-a[1]);
  if (entries.length === 0) { el.innerHTML = '<span class="c0">매칭된 키워드 없음</span>'; return; }
  const max = entries[0][1], min = entries[entries.length-1][1];
  const range = Math.max(1, max-min);
  el.innerHTML = entries.map(([k,v]) => {
    const t = Math.round(((v-min)/range)*4);
    return `<span class="c${t}">${k} (${v})</span>`;
  }).join("");
}

// ── 메인 렌더 ──
function renderAll() {
  const f = getFilters();
  const items = applyFilter(AUGMENTED, f);
  document.getElementById("page-stat").textContent =
    `필터 결과 ${items.length.toLocaleString()}건 / 전체 ${AUGMENTED.length.toLocaleString()}건${TEMP_KW.length?` · 임시 키워드 ${TEMP_KW.length}개 적용`:""}`;

  const A = aggregate(items);
  // 사업 분류(A·B) 상위 카테고리 산출
  const topAB = (() => {
    const bA={}, bB={};
    for (const it of items) {
      const a=classifyA(it).label, b=classifyB(it).label;
      bA[a]=(bA[a]||0)+1; bB[b]=(bB[b]||0)+1;
    }
    const topA = Object.entries(bA).sort((x,y)=>y[1]-x[1])[0];
    const topB = Object.entries(bB).sort((x,y)=>y[1]-x[1])[0];
    return { topA, topB };
  })();
  // INK 평균 수주가능성
  const avgWinPct = items.length ? Math.round(items.reduce((s,it)=>s+winProbability(it).pct, 0) / items.length) : 0;
  const highWinCount = items.filter(it => winProbability(it).pct >= 75).length;

  // ⭐ 찜한 공고 개수 (현재 필터된 목록 안)
  const favInFilter = items.filter(it => isFav(it)).length;

  document.getElementById("kpis").innerHTML = [
    {label:"총 공고건수", value: items.length.toLocaleString(), sub:`발주계획 ${A.byType.order_plan||0} · 사전규격 ${A.byType.pre_spec||0}`},
    {label:"⭐ 찜한 공고", value: FAVORITES.size, sub: favInFilter !== FAVORITES.size ? `현재 필터에 ${favInFilter}건` : "관리자 관심 표시"},
    {label:"총 예산금액", value: fmtMoney(A.budgetSum), sub:"금액공개 건 합계"},
    {label:"평균 예산", value: fmtMoney(A.budgetAvg), sub:"금액공개 건 평균"},
    {label:"🎯 최다 성격", value: topAB.topA ? topAB.topA[0] : "-", sub: topAB.topA ? `${topAB.topA[1]}건` : ""},
    {label:"🌾 최다 분야", value: topAB.topB ? topAB.topB[0] : "-", sub: topAB.topB ? `${topAB.topB[1]}건` : ""},
    {label:"평균 수주가능성", value: avgWinPct + "%", sub:`유망(≥75%) ${highWinCount}건`},
    {label:"지역 수", value: Object.keys(A.byRegion).length, sub:"발주기관 소재 기준"},
  ].map(k => `<div class="kpi"><div class="label">${k.label}</div><div class="value">${k.value||"-"}</div><div class="sub">${k.sub||""}</div></div>`).join("");

  if (CURRENT_VIEW === "dashboard") {
    timelineChart("ch-timeline", items);
    doughnutChart("ch-type", ["발주계획","사전규격"], [A.byType.order_plan||0, A.byType.pre_spec||0]);
    barChart("ch-budget", BUDGET_BUCKETS.map(b=>b.label), BUDGET_BUCKETS.map(b=>A.byBudget[b.label]||0), "건수", {colors:"#0891b2"});
    const at = Object.entries(A.byAgencyType).sort((a,b)=>b[1]-a[1]);
    barChart("ch-agency-type", at.map(([k])=>k), at.map(([,v])=>v), "건수", {colors:"#7e22ce"});
    const reg = Object.entries(A.byRegion).sort((a,b)=>b[1]-a[1]);
    barChart("ch-region", reg.map(([k])=>k), reg.map(([,v])=>v), "건수", {colors:"#16a34a"});
    const ag = Object.entries(A.byAgency).sort((a,b)=>b[1]-a[1]).slice(0,10);
    barChart("ch-agency", ag.map(([k])=>k), ag.map(([,v])=>v), "건수", {horizontal:true, colors:"#d97706"});
    const kw = Object.entries(A.byKw).sort((a,b)=>b[1]-a[1]);
    barChart("ch-keyword", kw.map(([k])=>k), kw.map(([,v])=>v), "건수", {colors:"#2563eb"});
    renderKwCloud(A.byKw);
    renderInkRanking(items);
  }
  if (CURRENT_VIEW === "keyword") {
    // === 사업 분류 분석 ===
    // 각 items 를 A / B / C 축으로 태깅
    const byA = {}, byB = {}, byAB = {}, byCB = {};
    for (const it of items) {
      const a = classifyA(it), b = classifyB(it), c = classifyC(it);
      byA[a.label] = (byA[a.label]||0) + 1;
      byB[b.label] = (byB[b.label]||0) + 1;
      const abKey = a.label + "||" + b.label;
      byAB[abKey] = (byAB[abKey]||0) + 1;
      if (!byCB[c.label]) byCB[c.label] = {};
      byCB[c.label][b.label] = (byCB[c.label][b.label]||0) + 1;
    }
    // A축 도넛
    const aEntries = Object.entries(byA).sort((x,y)=>y[1]-x[1]);
    doughnutChart("ch-classA", aEntries.map(x=>x[0]), aEntries.map(x=>x[1]));
    // B축 바 (수평)
    const bEntries = Object.entries(byB).sort((x,y)=>y[1]-x[1]);
    barChart("ch-classB", bEntries.map(x=>x[0]), bEntries.map(x=>x[1]), "건수", {horizontal:true, colors:"#059669"});
    // A × B 스택바
    const bLabels2 = bEntries.map(x=>x[0]);
    stackedBar("ch-classAB", bLabels2, aEntries.map(([a],idx) => ({
      label: a,
      data: bLabels2.map(bl => byAB[a+"||"+bl] || 0),
      backgroundColor: CLASSIFY_A_COLORS[(CLASSIFY_A.find(x=>x.label===a)||{}).code] || "#cbd5e1",
    })));
    // C × B 스택바
    const cEntries = Object.entries(byCB).sort((x,y)=>Object.values(y[1]).reduce((a,b)=>a+b,0) - Object.values(x[1]).reduce((a,b)=>a+b,0));
    stackedBar("ch-classCB", cEntries.map(x=>x[0]), bLabels2.map((bl,idx) => ({
      label: bl,
      data: cEntries.map(([c]) => byCB[c][bl] || 0),
      backgroundColor: CLASSIFY_B_COLORS[(CLASSIFY_B.find(x=>x.label===bl)||{}).code] || "#cbd5e1",
    })));
    // (참고) 원 키워드
    const kw = Object.entries(A.byKw).sort((a,b)=>b[1]-a[1]);
    barChart("ch-keyword2", kw.map(([k])=>k), kw.map(([,v])=>v), "건수", {colors:"#2563eb", horizontal:true});
  }
  if (CURRENT_VIEW === "region") {
    const reg = Object.entries(A.byRegion).sort((a,b)=>b[1]-a[1]);
    barChart("ch-region2", reg.map(([k])=>k), reg.map(([,v])=>v), "건수", {colors:"#16a34a"});
    // 지역 × 분야(B)
    const byRegB = {};
    const byRegA = {};
    for (const it of items) {
      const r = it.region || "중앙/전국";
      const b = classifyB(it).label;
      const a = classifyA(it).label;
      if (!byRegB[r]) byRegB[r] = {};
      if (!byRegA[r]) byRegA[r] = {};
      byRegB[r][b] = (byRegB[r][b]||0) + 1;
      byRegA[r][a] = (byRegA[r][a]||0) + 1;
    }
    const bLabels = CLASSIFY_B.map(x=>x.label).concat(["기타"]);
    stackedBar("ch-region-domain", reg.map(([r])=>r), bLabels.map((bl) => ({
      label: bl,
      data: reg.map(([r]) => (byRegB[r]||{})[bl] || 0),
      backgroundColor: CLASSIFY_B_COLORS[(CLASSIFY_B.find(x=>x.label===bl)||{}).code] || "#cbd5e1",
    })).filter(ds => ds.data.some(v => v > 0)));
    const aLabels = CLASSIFY_A.map(x=>x.label).concat(["기타"]);
    stackedBar("ch-region-nature", reg.map(([r])=>r), aLabels.map((al) => ({
      label: al,
      data: reg.map(([r]) => (byRegA[r]||{})[al] || 0),
      backgroundColor: CLASSIFY_A_COLORS[(CLASSIFY_A.find(x=>x.label===al)||{}).code] || "#cbd5e1",
    })).filter(ds => ds.data.some(v => v > 0)));
    const regLabels = reg.map(([k])=>k);
    stackedBar("ch-region-type", regLabels, [
      { label:"발주계획", data: regLabels.map(k=>(A.byRegionType[k]||{}).order_plan||0), backgroundColor:"#2563eb" },
      { label:"사전규격", data: regLabels.map(k=>(A.byRegionType[k]||{}).pre_spec||0),  backgroundColor:"#16a34a" },
    ]);
  }
  if (CURRENT_VIEW === "list") renderList(items, f);
}

function renderList(items, f) {
  let arr = items.slice();
  // 경과(마감/기간종료) 공고는 목록에서 기본 제외 (체크박스로 포함 가능). 차트·KPI는 영향 없음.
  const showExp = (document.getElementById("showExpired") || {}).checked;
  let hidden = 0;
  if (!showExp) {
    const before = arr.length;
    arr = arr.filter(it => !isExpired(it));
    hidden = before - arr.length;
  }
  // "⭐ 찜한 것만" 체크박스가 켜지면 찜한 것만 표시
  const favOnly = (document.getElementById("favOnly") || {}).checked;
  if (favOnly) arr = arr.filter(it => isFav(it));

  // 최신순: 발주예정일 우선, 없으면 갱신일 — 둘 다 ISO/yyyy-mm-dd 형식이라 문자열 정렬 OK
  const dateKey = it => it.order_planned_date || it.last_seen_at || "";
  if (f.sort === "recent") arr.sort((a,b) => dateKey(b).localeCompare(dateKey(a)));
  if (f.sort === "budget") arr.sort((a,b) => (b.budget_amount||0) - (a.budget_amount||0));
  if (f.sort === "score")  arr.sort((a,b) => (b.score||0) - (a.score||0));

  // ⭐ 찜한 것들은 항상 최상단 재배치 (정렬 옵션 무관)
  arr.sort((a,b) => (isFav(b)?1:0) - (isFav(a)?1:0));

  const el = document.getElementById("list");
  const note = hidden ? `<div class="list-note">진행 중 공고 <b>${arr.length}</b>건 · 경과(마감) <b>${hidden}</b>건 숨김 — '경과 공고 포함'으로 전체 보기</div>` : "";
  if (arr.length === 0) { el.innerHTML = note + '<div class="empty">진행 중인(미경과) 공고가 없습니다. \'경과 공고 포함\'을 켜면 전체를 볼 수 있습니다.</div>'; return; }
  el.innerHTML = note + arr.slice(0, 300).map(it => {
    const kws = (it.matched_keywords||[]).map(k => {
      const isTemp = (it._temp_matched||[]).includes(k);
      return `<span class="chip${isTemp?" temp":""}" title="${isTemp?"임시 키워드":"yml 키워드"}">${k}</span>`;
    }).join("");
    const atts = (it.attachments||[]).length
      ? `<div class="attachments">${it.attachments.map((a,i)=>`<a class="att-link" href="${a.url}" target="_blank" rel="noopener">📎 ${a.name||('첨부'+(i+1))}</a>`).join("")}</div>`
      : (it.url ? `<div class="attachments"><a class="att-link" href="${it.url}" target="_blank" rel="noopener">📎 첨부/원문</a></div>` : "");
    const budget = it.budget_amount ? ` · ${fmtMoney(it.budget_amount)}` : "";
    const region = it.region ? ` · ${it.region}` : "";
    const desc = it.description ? `<details><summary>상세</summary><pre>${(it.description||"").replace(/</g,"&lt;")}</pre></details>` : "";
    // 수주 가능성 + 분류 태그
    const wp = winProbability(it);
    const grade = wp.pct>=75?"high":wp.pct>=55?"mid":"low";
    const a = classifyA(it), b = classifyB(it);
    const fav = isFav(it);
    return `<article class="card${fav?" favorite":""}">
      <div class="card-head">
        <h4 class="card-title">${it.title||""}</h4>
        <div class="card-prob" title="수주 가능성 추정 %">
          <button class="fav-btn${fav?" active":""}" data-fav-ext="${it.external_id}" title="${fav?"찜 해제":"찜하기 (상단 재배치)"}">${fav?"★":"☆"} <span style="font-size:11px">${fav?"찜":"찜"}</span></button>
          <span class="ink-bar ${grade}"><span style="width:${wp.pct}%"></span></span>
          <b class="pct-${grade}">${wp.pct}%</b>
        </div>
      </div>
      <div class="agency">${it.agency||""}${it.agency_dept?" / "+it.agency_dept:""}${budget}${region}${it.bsns_div?" · "+it.bsns_div:""}</div>
      <div class="chips">
        <span class="chip type">${TYPE_LABEL[it.source_type]||it.source_type}</span>
        <span class="chip class-a">${a.code} ${a.label}</span>
        <span class="chip class-b">${b.code} ${b.label}</span>
        <span class="chip agency_type">${it.agency_type||"기타"}</span>
        <span class="chip region">${it.region||"중앙/전국"}</span>
        ${kws}
      </div>
      ${atts}
      <div style="font-size:11.5px;color:var(--muted);margin-top:6px">
        최근 갱신 ${fmtDate(it.last_seen_at)} · 최초 발견 ${fmtDate(it.first_seen_at)}${it.officer?" · 담당 "+it.officer:""}${it.officer_tel?" ("+it.officer_tel+")":""}
      </div>
      ${desc}
    </article>`;
  }).join("") + (arr.length > 300 ? `<div class="footnote">상위 300건만 표시됨 (전체 ${arr.length}건). 필터를 좁혀 보세요.</div>` : "");

  // ⭐ 찜하기 버튼 이벤트 delegation
  el.querySelectorAll(".fav-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const extId = btn.dataset.favExt;
      if (extId) {
        toggleFav(extId);
        renderAll();  // 재정렬(찜한 것 최상단)
      }
    });
  });
}

// ── 엑셀 다운로드 ──
function downloadXlsx() {
  const f = getFilters();
  const items = applyFilter(AUGMENTED, f);
  if (items.length === 0) { alert("다운로드할 데이터가 없습니다. 필터를 확인하세요."); return; }
  // 찜한 것 먼저, 그 다음 원본 순서
  const sortedItems = items.slice().sort((a,b) => (isFav(b)?1:0) - (isFav(a)?1:0));
  const rows = sortedItems.map(it => {
    const a = it.attachments || [];
    const wp = winProbability(it);
    const clsA = classifyA(it), clsB = classifyB(it), clsC = classifyC(it);
    const row = {
      "⭐ 찜": isFav(it) ? "★" : "",
      "수주가능성(%)": wp.pct,
      "성격(A)": clsA.code + " " + clsA.label,
      "분야(B)": clsB.code + " " + clsB.label,
      "발주주체(C)": clsC.code + " " + clsC.label,
      "공고유형": TYPE_LABEL[it.source_type] || it.source_type,
      "제목": it.title || "",
      "발주기관": it.agency || "",
      "발주부서": it.agency_dept || "",
      "발주처유형": it.agency_type || "",
      "지역": it.region || "",
      "업무구분": it.bsns_div || "",
      "계약방법": it.contract_method || "",
      "예산금액(원)": typeof it.budget_amount === "number" ? it.budget_amount : "",
      "예산금액(억)": typeof it.budget_amount === "number" ? +(it.budget_amount/1e8).toFixed(2) : "",
      "매칭키워드": (it.matched_keywords||[]).join(", "),
      "점수": it.score || 0,
      "발주예정일": fmtDate(it.order_planned_date),
      "최초발견일": fmtDate(it.first_seen_at),
      "최근갱신일": fmtDate(it.last_seen_at),
      "참조번호": it.ref_no || "",
      "담당자": it.officer || "",
      "연락처": it.officer_tel || "",
    };
    for (let i = 0; i < 5; i++) {
      row[`첨부${i+1} 파일명`] = a[i] ? (a[i].name || "") : "";
      row[`첨부${i+1} URL`]   = a[i] ? (a[i].url  || "") : "";
    }
    row["상세설명"] = it.description || "";
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  // 열 너비 자동
  const headers = Object.keys(rows[0] || {});
  ws["!cols"] = headers.map(h => {
    const maxLen = Math.max(h.length, ...rows.map(r => String(r[h]||"").length));
    return { wch: Math.min(60, Math.max(10, maxLen + 2)) };
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "공고목록");
  const today = new Date().toISOString().slice(0,10).replace(/-/g,"");
  XLSX.writeFile(wb, `nara-radar_${today}_${items.length}건.xlsx`);
}

// ── 키워드 설정 UI ──
function renderTempKwList() {
  // 사이드바 설정 뷰의 상세 관리
  const el = document.getElementById("temp-kw-list");
  if (el) {
    if (TEMP_KW.length === 0) {
      el.innerHTML = '<div style="color:var(--muted);font-size:13px">아직 추가된 임시 키워드가 없습니다.</div>';
    } else {
      el.innerHTML = TEMP_KW.map((k,i) =>
        `<span class="chip-edit">${k.term} <small style="color:#a16207">×${k.weight}</small>
           <button data-idx="${i}" title="삭제">×</button></span>`
      ).join(" ");
      el.querySelectorAll("button").forEach(b => b.addEventListener("click", () => {
        TEMP_KW.splice(+b.dataset.idx, 1);
        localStorage.setItem(LS_TEMP_KW, JSON.stringify(TEMP_KW));
        rebuildAugmented(); renderTempKwList(); renderAll();
      }));
    }
  }
  // 필터바 인라인 목록
  const inline = document.getElementById("tk-list");
  if (inline) {
    inline.innerHTML = TEMP_KW.map((k,i) =>
      `<span class="tk-chip">${k.term}<button data-tki="${i}">×</button></span>`
    ).join(" ");
    inline.querySelectorAll("button").forEach(b => b.addEventListener("click", () => {
      TEMP_KW.splice(+b.dataset.tki, 1);
      localStorage.setItem(LS_TEMP_KW, JSON.stringify(TEMP_KW));
      rebuildAugmented(); renderTempKwList(); renderAll();
    }));
  }
}

function inlineAddTk() {
  const inp = document.getElementById("tk-in");
  if (!inp) return;
  const t = inp.value.trim();
  if (!t) return;
  if (!TEMP_KW.some(k => k.term === t)) TEMP_KW.push({term: t, weight: 5});
  localStorage.setItem(LS_TEMP_KW, JSON.stringify(TEMP_KW));
  inp.value = "";
  rebuildAugmented(); renderTempKwList(); renderAll();
}
function updateGithubEditLink() {
  let u = document.getElementById("repo-url").value.trim();
  if (!u) {
    document.getElementById("github-edit").style.pointerEvents = "none";
    document.getElementById("github-edit").style.opacity = "0.5";
    document.getElementById("github-edit").removeAttribute("href");
    return;
  }
  u = u.replace(/\/$/, "");
  document.getElementById("github-edit").href = `${u}/edit/main/keywords.yml`;
  document.getElementById("github-edit").style.pointerEvents = "";
  document.getElementById("github-edit").style.opacity = "";
  localStorage.setItem(LS_REPO_URL, u);
}
function updateNewYmlLine() {
  const t = document.getElementById("yml-term").value.trim();
  const w = document.getElementById("yml-weight").value || 5;
  const g = document.getElementById("yml-group").value.trim() || "custom";
  const el = document.getElementById("new-yml-line");
  if (!t) { el.textContent = "# 키워드와 그룹을 입력하면 여기에 자동 생성됩니다."; return; }
  el.textContent = `  - { term: "${t}", weight: ${w}, group: "${g}" }`;
}

// ── 뷰 전환 ──
function setView(name) {
  CURRENT_VIEW = name;
  document.querySelectorAll(".nav button").forEach(b => b.classList.toggle("active", b.dataset.view === name));
  document.querySelectorAll("[data-view-pane]").forEach(p => p.style.display = (p.dataset.viewPane === name) ? "" : "none");
  // 필터바는 settings/about 뷰에선 숨김
  document.getElementById("filter-bar").style.display = (name === "settings" || name === "about") ? "none" : "";
  const titles = { dashboard:"대시보드", list:"공고 목록", keyword:"키워드 분석", region:"지역 분석", settings:"키워드 설정", about:"설명" };
  const subs = {
    dashboard:"검색어 매칭 발주계획·사전규격 통합 통계",
    list:"필터·정렬된 공고 목록 (엑셀 다운로드 가능)",
    keyword:"키워드별 매칭 빈도와 유형 분포",
    region:"지역별 공고 분포",
    settings:"임시(localStorage) / 영구(keywords.yml) 키워드 관리",
    about:"이 페이지에 대하여",
  };
  document.getElementById("page-title").textContent = titles[name];
  document.getElementById("page-sub").textContent = subs[name];
  renderAll();
}

// ── UI 이벤트 바인딩 (데이터 로드 성공 여부와 무관하게 항상 실행) ──
const DEFAULT_REPO_URL = "https://github.com/HyunJongKi/nara-weekly";
function initUI() {
  // 필터
  const _se = document.getElementById("showExpired"); if (_se) _se.addEventListener("change", renderAll);
  const _fo = document.getElementById("favOnly"); if (_fo) _fo.addEventListener("change", renderAll);
  ["q","type","kw","region","agency_type","sort"].forEach(id => {
    const el = document.getElementById(id); if (el) el.addEventListener("input", renderAll);
  });
  document.getElementById("reset").addEventListener("click", () => {
    ["q","type","kw","region","agency_type"].forEach(id => document.getElementById(id).value = "");
    document.getElementById("sort").value = "recent";
    renderAll();
  });
  document.getElementById("xlsx").addEventListener("click", downloadXlsx);
  // 필터바 인라인 임시 키워드
  const tkBtn = document.getElementById("tk-add"); if (tkBtn) tkBtn.addEventListener("click", inlineAddTk);
  const tkIn = document.getElementById("tk-in");
  if (tkIn) tkIn.addEventListener("keydown", e => { if (e.key === "Enter") inlineAddTk(); });
  document.querySelectorAll(".nav button").forEach(b =>
    b.addEventListener("click", () => setView(b.dataset.view))
  );

  // 키워드 설정 — 임시 키워드(localStorage)
  document.getElementById("add-kw").addEventListener("click", () => {
    const t = document.getElementById("new-term").value.trim();
    const w = parseInt(document.getElementById("new-weight").value || "5", 10);
    if (!t) return;
    if (!TEMP_KW.some(k => k.term === t)) TEMP_KW.push({term: t, weight: w});
    localStorage.setItem(LS_TEMP_KW, JSON.stringify(TEMP_KW));
    document.getElementById("new-term").value = "";
    rebuildAugmented(); renderTempKwList(); renderAll();
  });
  document.getElementById("clear-kw").addEventListener("click", () => {
    if (!confirm("임시 키워드를 모두 삭제할까요?")) return;
    TEMP_KW = []; localStorage.setItem(LS_TEMP_KW, "[]");
    rebuildAugmented(); renderTempKwList(); renderAll();
  });
  document.getElementById("new-term").addEventListener("keydown", e => { if (e.key === "Enter") document.getElementById("add-kw").click(); });

  // 영구 반영 — GitHub repo URL & yml 한 줄 (repo URL 기본값 자동 입력 → 편집 버튼 즉시 활성)
  const savedRepo = localStorage.getItem(LS_REPO_URL);
  document.getElementById("repo-url").value = savedRepo || DEFAULT_REPO_URL;
  updateGithubEditLink();
  document.getElementById("repo-url").addEventListener("input", updateGithubEditLink);
  ["yml-term","yml-weight","yml-group"].forEach(id => document.getElementById(id).addEventListener("input", updateNewYmlLine));

  renderTempKwList();
  setView("dashboard");
}
initUI();

// ── 데이터 로드 (실패해도 위 UI 컨트롤은 정상 동작) ──
fetch("./tenders.json?ts=" + Date.now())
  .then(r => r.json())
  .then(data => {
    ALL = dedupe(data.items || []);
    document.getElementById("generated").textContent = "최근 수집 " + (fmtDate(data.generated_at) || "—");
    // keywords.yml 미리보기 (render.py 가 docs/keywords.yml 로 미러링한 것)
    fetch("./keywords.yml?ts=" + Date.now()).then(r => r.ok ? r.text() : "").then(t => {
      document.getElementById("server-yml").textContent = t || "(keywords.yml 을 불러오지 못했습니다.)";
    }).catch(() => {});

    rebuildAugmented();

    // 필터 옵션 (지역, 발주처유형)
    const rSet = new Set(), atSet = new Set();
    AUGMENTED.forEach(it => { if (it.region) rSet.add(it.region); if (it.agency_type) atSet.add(it.agency_type); });
    const fill = (id, set) => {
      const el = document.getElementById(id);
      [...set].sort().forEach(v => {
        const o = document.createElement("option"); o.value = v; o.textContent = v; el.appendChild(o);
      });
    };
    fill("region", rSet); fill("agency_type", atSet);
    renderAll();
  })
  .catch(err => {
    const pane = document.querySelector('[data-view-pane="dashboard"]');
    if (pane) pane.innerHTML = '<div class="empty">데이터를 불러오지 못했습니다: ' + err + '</div>';
  });
</script>
</body>
</html>
"""


def main() -> int:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(HTML, encoding="utf-8")
    print(f"wrote {OUTPUT}")
    # docs/ 에 keywords.yml 미러링 (대시보드 키워드 설정 뷰에서 fetch 용)
    if KEYWORDS_SRC.exists():
        KEYWORDS_DST.write_text(KEYWORDS_SRC.read_text(encoding="utf-8"), encoding="utf-8")
        print(f"mirrored {KEYWORDS_DST}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
