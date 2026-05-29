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
<title>Nara Weekly — 농림업 학술연구용역 모니터</title>
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
      <h1>🌾 Nara Weekly</h1>
      <div class="sub">농림업 학술연구용역 모니터</div>
    </div>
    <nav class="nav">
      <button class="active" data-view="dashboard"><span class="icon">📊</span><span class="label">대시보드</span></button>
      <button data-view="list"><span class="icon">📋</span><span class="label">공고 목록</span></button>
      <button data-view="keyword"><span class="icon">🔍</span><span class="label">키워드 분석</span></button>
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
        <option value="recent">최근 갱신순</option>
        <option value="score">점수순</option>
        <option value="budget">예산순</option>
      </select>
      <button class="btn" id="reset">초기화</button>
      <button class="btn primary" id="xlsx">⬇ 엑셀 다운로드</button>
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
          <h3>🎯 아이엔케이(주) 수주 가능성 순위 <span class="hint">추정 모델 · 가능성순 정렬</span></h3>
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

    <!-- 키워드 분석 -->
    <section data-view-pane="keyword" style="display:none">
      <div class="charts">
        <div class="panel col-12">
          <h3>키워드별 공고건수</h3>
          <div class="chart-box tall"><canvas id="ch-keyword2"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>키워드 × 공고유형 (스택 바)</h3>
          <div class="chart-box tall"><canvas id="ch-keyword-type"></canvas></div>
        </div>
      </div>
    </section>

    <!-- 지역 분석 -->
    <section data-view-pane="region" style="display:none">
      <div class="charts">
        <div class="panel col-12">
          <h3>지역별 공고건수</h3>
          <div class="chart-box tall"><canvas id="ch-region2"></canvas></div>
        </div>
        <div class="panel col-12">
          <h3>지역 × 공고유형 (스택 바)</h3>
          <div class="chart-box tall"><canvas id="ch-region-type"></canvas></div>
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
  { label: "1억 미만",   max: 1e8 },
  { label: "1~5억",      max: 5e8 },
  { label: "5~10억",     max: 10e8 },
  { label: "10~50억",    max: 50e8 },
  { label: "50~100억",   max: 100e8 },
  { label: "100억 이상", max: Infinity },
];
const COLORS = ["#2563eb","#16a34a","#dc2626","#d97706","#7e22ce","#0891b2","#db2777","#65a30d","#475569","#0ea5e9","#a16207","#be123c"];
const LS_TEMP_KW = "nw_temp_kw_v1";
const LS_REPO_URL = "nw_repo_url_v1";

let ALL = [];            // 원본
let AUGMENTED = [];      // 임시 키워드 적용본
let CHARTS = {};
let CURRENT_VIEW = "dashboard";
let TEMP_KW = JSON.parse(localStorage.getItem(LS_TEMP_KW) || "[]");

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
  const ranked = items.map(it => ({ it, ...winProbability(it) })).sort((a,b) => b.pct - a.pct);
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
  document.getElementById("kpis").innerHTML = [
    {label:"총 공고건수", value: items.length.toLocaleString(), sub:`발주계획 ${A.byType.order_plan||0} · 사전규격 ${A.byType.pre_spec||0}`},
    {label:"매칭 키워드 수", value: Object.keys(A.byKw).length, sub:"매칭된 회사 관심어"},
    {label:"총 예산금액", value: fmtMoney(A.budgetSum), sub:"금액공개 건 합계"},
    {label:"평균 예산", value: fmtMoney(A.budgetAvg), sub:"금액공개 건 평균"},
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
    const kw = Object.entries(A.byKw).sort((a,b)=>b[1]-a[1]);
    barChart("ch-keyword2", kw.map(([k])=>k), kw.map(([,v])=>v), "건수", {colors:"#2563eb", horizontal:true});
    const kwLabels = kw.map(([k])=>k);
    stackedBar("ch-keyword-type", kwLabels, [
      { label:"발주계획", data: kwLabels.map(k=>(A.byKwType[k]||{}).order_plan||0), backgroundColor:"#2563eb" },
      { label:"사전규격", data: kwLabels.map(k=>(A.byKwType[k]||{}).pre_spec||0),  backgroundColor:"#16a34a" },
    ]);
  }
  if (CURRENT_VIEW === "region") {
    const reg = Object.entries(A.byRegion).sort((a,b)=>b[1]-a[1]);
    barChart("ch-region2", reg.map(([k])=>k), reg.map(([,v])=>v), "건수", {colors:"#16a34a"});
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
  if (f.sort === "recent") arr.sort((a,b) => (b.last_seen_at||"").localeCompare(a.last_seen_at||""));
  if (f.sort === "budget") arr.sort((a,b) => (b.budget_amount||0) - (a.budget_amount||0));
  if (f.sort === "score")  arr.sort((a,b) => (b.score||0) - (a.score||0));

  const el = document.getElementById("list");
  if (arr.length === 0) { el.innerHTML = '<div class="empty">조건에 맞는 결과가 없습니다.</div>'; return; }
  el.innerHTML = arr.slice(0, 300).map(it => {
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
    return `<article class="card">
      <h4>${it.title||""}</h4>
      <div class="agency">${it.agency||""}${it.agency_dept?" / "+it.agency_dept:""}${budget}${region}${it.bsns_div?" · "+it.bsns_div:""}</div>
      <div class="chips">
        <span class="chip type">${TYPE_LABEL[it.source_type]||it.source_type}</span>
        <span class="chip agency_type">${it.agency_type||"기타"}</span>
        <span class="chip region">${it.region||"중앙/전국"}</span>
        <span class="chip score">점수 ${it.score||0}</span>
        ${kws}
      </div>
      ${atts}
      <div style="font-size:11.5px;color:var(--muted);margin-top:6px">
        최근 갱신 ${fmtDate(it.last_seen_at)} · 최초 발견 ${fmtDate(it.first_seen_at)}${it.officer?" · 담당 "+it.officer:""}${it.officer_tel?" ("+it.officer_tel+")":""}
      </div>
      ${desc}
    </article>`;
  }).join("") + (arr.length > 300 ? `<div class="footnote">상위 300건만 표시됨 (전체 ${arr.length}건). 필터를 좁혀 보세요.</div>` : "");
}

// ── 엑셀 다운로드 ──
function downloadXlsx() {
  const f = getFilters();
  const items = applyFilter(AUGMENTED, f);
  if (items.length === 0) { alert("다운로드할 데이터가 없습니다. 필터를 확인하세요."); return; }
  const rows = items.map(it => {
    const a = it.attachments || [];
    const row = {
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
  XLSX.writeFile(wb, `nara-weekly_${today}_${items.length}건.xlsx`);
}

// ── 키워드 설정 UI ──
function renderTempKwList() {
  const el = document.getElementById("temp-kw-list");
  if (!el) return;
  if (TEMP_KW.length === 0) {
    el.innerHTML = '<div style="color:var(--muted);font-size:13px">아직 추가된 임시 키워드가 없습니다.</div>';
    return;
  }
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

// ── 초기화 ──
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

    // 이벤트
    ["q","type","kw","region","agency_type","sort"].forEach(id =>
      document.getElementById(id).addEventListener("input", renderAll)
    );
    document.getElementById("reset").addEventListener("click", () => {
      ["q","type","kw","region","agency_type"].forEach(id => document.getElementById(id).value = "");
      document.getElementById("sort").value = "recent";
      renderAll();
    });
    document.getElementById("xlsx").addEventListener("click", downloadXlsx);

    document.querySelectorAll(".nav button").forEach(b =>
      b.addEventListener("click", () => setView(b.dataset.view))
    );

    // 키워드 설정 이벤트
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

    // 영구 반영 — GitHub repo URL & yml 한 줄
    const savedRepo = localStorage.getItem(LS_REPO_URL);
    if (savedRepo) document.getElementById("repo-url").value = savedRepo;
    updateGithubEditLink();
    document.getElementById("repo-url").addEventListener("input", updateGithubEditLink);
    ["yml-term","yml-weight","yml-group"].forEach(id => document.getElementById(id).addEventListener("input", updateNewYmlLine));

    renderTempKwList();
    setView("dashboard");
  })
  .catch(err => {
    document.querySelector('[data-view-pane="dashboard"]').innerHTML =
      '<div class="empty">데이터를 불러오지 못했습니다: ' + err + '</div>';
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
