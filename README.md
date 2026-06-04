# Nara Radar — 농림업 학술연구용역 모니터

나라장터(조달청) **발주계획**·**사전규격** 공고 중
회사 관심 검색어(농업·농촌·식량·원예·스마트팜·노지 스마트·산림·임업)에 매칭되는
**학술연구용역** 만을 **주 2회**(매주 수·금 10:00 KST) 자동 수집하여,
**카카오톡 '나에게 보내기'** 알림과 **GitHub Pages 정적 대시보드** 로 조회한다.

호스팅 비용 0원. 별도 서버 없음. GitHub Actions가 모든 일을 한다.

```
[매주 수·금 10:00 KST]
  GitHub Actions cron
    ├─ scripts/collect.py   ─ 나라장터 OpenAPI 호출 → 학술연구용역 필터 → 키워드 매칭 → JSON 저장
    ├─ scripts/render.py    ─ docs/index.html (사이드바 + Chart.js 시각화) 재생성
    └─ scripts/kakao.py     ─ 카카오톡 '나에게 보내기' 발송 (신규 공고만)
        ↓
  변경분 git commit & push → GitHub Pages 자동 재배포
```

## 대시보드 구성

- **사이드 메뉴바**: 대시보드 / 공고 목록 / 키워드 분석 / 지역 분석 / **키워드 설정** / 설명
- **KPI 카드**: 총 공고건수, 매칭 키워드 수, 총 예산금액, 평균 예산, 지역 수
- **시각화 (Chart.js)**
  - 일자별 공고건수 추이 (발주계획 vs 사전규격 vs 전체)
  - 공고유형 비중 (도넛)
  - 금액규모별 분포 (6구간: 1억 미만 ~ 100억 이상)
  - 발주처 유형별 (중앙부처/지자체/공공기관/교육기관/기타)
  - 지역별 공고건수
  - 발주기관 상위 10
  - 키워드별 언급 빈도 (바차트 + 간이 워드클라우드)
- **공통 필터**: 검색어, 공고유형, 키워드, 지역, 발주처 유형, 정렬
- **엑셀 다운로드** (`⬇ 엑셀 다운로드` 버튼): 현재 필터된 결과를 `.xlsx` 로 즉시 다운로드.
  컬럼 — 공고유형, 제목, 발주기관/부서/유형, 지역, 업무구분, 계약방법, 예산금액(원/억),
  매칭키워드, 점수, 발주예정일, 최초/최근 일자, 참조번호, 담당자/연락처,
  **첨부1~5 파일명+URL (과업지시서·제안요청서·규격서 등)**, 상세설명.
- **키워드 설정** (사이드바 ⚙ 메뉴):
  - *임시 키워드*: 새 키워드를 입력하면 `localStorage` 에 저장되어 본인 브라우저의 대시보드 필터·차트에 즉시 반영. 카톡 알림에는 영향 없음.
  - *영구 반영*: 본인 GitHub repo URL 을 입력하면 `keywords.yml` 편집 페이지로 바로 이동하는 링크가 생성됨. YAML 한 줄도 자동으로 생성해 줌.

---

## 1. 셋업 (최초 1회)

### 1-1. 이 폴더를 GitHub 리포지토리로 만들기

```powershell
cd C:\Users\CEO\nara-weekly
git init
git add .
git commit -m "init"
# GitHub에 빈 repo (예: nara-weekly) 만든 뒤
git remote add origin https://github.com/<your-id>/nara-weekly.git
git branch -M main
git push -u origin main
```

**public repo로 만들면 GitHub Actions 가 무제한 무료**이고 Pages 설정도 단순.
private repo도 월 2,000분 무료 한도 내에서 동작 (이 워크플로 1회 ≈ 1분).

### 1-2. 나라장터 OpenAPI 키 발급

공공데이터포털에서 아래 두 서비스를 활용신청 (자동승인):

| 서비스 | End Point |
|---|---|
| 조달청\_나라장터 발주계획현황서비스 | `https://apis.data.go.kr/1230000/ao/OrderPlanSttusService` |
| 조달청\_나라장터 사전규격정보서비스 | `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService` |

활용기간 2년, 일일 트래픽 1,000건/오퍼레이션. 인증키(Decoding)를 복사해 둔다.

> 발급한 인증키(Decoding)는 GitHub Secret `G2B_SERVICE_KEY` 에만 저장하세요.
> (public repo 이므로 README 등 코드에 평문으로 적지 않습니다.)

### 1-3. 카카오 '나에게 보내기' 토큰 발급

이 단계가 가장 손이 많이 가지만 **단 한 번**만 하면 됩니다.

1. [Kakao Developers](https://developers.kakao.com) → **내 애플리케이션** → 앱 생성.
2. 앱 → **앱 설정 → 요약 정보** 에서 **REST API 키** 복사 → 이게 `KAKAO_REST_KEY`.
   > REST API 키는 GitHub Secret `KAKAO_REST_KEY` 에만 저장하세요 (코드에 평문 금지).
3. **제품 설정 → 카카오 로그인** → 활성화 ON.
4. **Redirect URI** 등록: `http://localhost:3000/api/auth/kakao/callback`
   (실제 도달하지 않아도 됨 — 인가 코드만 URL에서 빼올 거임)
5. **동의항목** → **카카오톡 메시지 전송 (`talk_message`)** → 사용 설정 ON.
6. 아래 URL을 브라우저에 입력 (`{REST_KEY}` 만 치환):
   ```
   https://kauth.kakao.com/oauth/authorize?response_type=code&client_id={REST_KEY}&redirect_uri=http://localhost:3000/api/auth/kakao/callback&scope=talk_message
   ```
7. 카카오 로그인 → 동의 → 브라우저가
   `http://localhost:3000/api/auth/kakao/callback?code=XXXX` 로 리디렉트되면서
   **연결 실패 화면**이 뜸 (정상). 주소창의 `code=` 뒤 값을 복사.
8. 로컬에서:
   ```powershell
   cd C:\Users\CEO\nara-weekly
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   # 클라이언트 시크릿이 활성화(ON)된 앱이면 마지막에 시크릿도 추가:
   python scripts/get_initial_token.py {REST_KEY} http://localhost:3000/api/auth/kakao/callback {복사한_CODE} {CLIENT_SECRET}
   ```
9. 출력에서 `refresh_token` 값을 복사 → 이게 `KAKAO_REFRESH_TOKEN`.
   > ℹ️ **클라이언트 시크릿**: 카카오 앱 → 카카오 로그인 → 보안 의 Client Secret 이
   > **OFF(현재 설정)면 불필요** — 토큰 교환·갱신에 `client_secret` 없이 동작합니다.
   > 만약 **ON**으로 켜면 토큰 교환·갱신 모두 `client_secret`이 필요해져
   > `KAKAO_CLIENT_SECRET` Secret 등록이 **필수**가 됩니다 (안 하면 자동 실행이 401로 실패).
   > 코드는 두 경우 모두 자동 대응합니다 (`KAKAO_CLIENT_SECRET` 있으면 사용, 없으면 생략).

### 1-4. GitHub Secrets / Variables 등록

GitHub repo → **Settings → Secrets and variables → Actions**.

**Secrets (Repository secrets):**

| 이름 | 값 |
|---|---|
| `G2B_SERVICE_KEY` | 공공데이터포털 인증키 (1-2) |
| `KAKAO_REST_KEY` | Kakao Developers REST API 키 (1-3 step 2) |
| `KAKAO_REFRESH_TOKEN` | 1-3 step 9 에서 얻은 값 |
| `KAKAO_CLIENT_SECRET` | (선택) 카카오 앱의 클라이언트 시크릿이 **ON**일 때만 필요. **현재는 OFF라 등록 불필요.** |

**Variables (Repository variables):**

| 이름 | 값 |
|---|---|
| `DASHBOARD_URL` | `https://<your-id>.github.io/nara-weekly/` |

### 1-5. GitHub Pages 켜기

repo → **Settings → Pages**:
- **Source**: `Deploy from a branch`
- **Branch**: `main`, 폴더 `/docs`

1~2분 지나면 `https://<your-id>.github.io/nara-weekly/` 가 살아납니다.

### 1-6. 첫 실행

repo → **Actions** → `weekly-nara` → **Run workflow**.
- 성공 시: 카톡 알림 + 대시보드 갱신.
- 실패 시: Actions 로그에서 어느 단계인지 확인.

---

## 2. 운영 메모

### 2-1. 스케줄

`.github/workflows/weekly.yml` 에 정의됨.

| 요일 | KST | UTC | cron |
|---|---|---|---|
| 수요일 | 10:00 | 01:00 | `0 1 * * 3` |
| 금요일 | 10:00 | 01:00 | `0 1 * * 5` |

`workflow_dispatch` 로 수동 실행도 가능.

### 2-2. 키워드 수정

**방법 1 — 영구 (카톡 알림 + 모두의 대시보드에 반영)**:
1. 대시보드 → **⚙ 키워드 설정** → 본인 repo URL 입력.
2. "GitHub 에서 keywords.yml 편집" 버튼 → 자동 생성된 YAML 한 줄 복사 → `keywords:` 아래에 붙여넣고 commit.
3. 다음 수·금 10시 자동 실행 시 반영 (즉시 보고 싶으면 Actions → Run workflow).

또는 로컬에서 직접 [`keywords.yml`](keywords.yml) 편집 → 푸시.
- 현재 키워드: 농업·농촌·식량·원예·스마트팜·노지 스마트·노지스마트·산림·임업
- `weight`: 점수 가중치. 여러 키워드가 매칭되면 합산.
- `min_score`: 카톡 발송 최소 점수 (대시보드에는 점수 무관 모두 노출).

**방법 2 — 임시 (본인 브라우저에만)**:
대시보드 → ⚙ 키워드 설정 → 임시 키워드 칸에 입력. 새로고침해도 유지됨 (localStorage).
즉시 차트와 필터에 반영되지만 카톡 알림과 다른 사람 대시보드에는 영향 없음.

### 2-3. 학술연구용역 필터

`scripts/collect.py` 의 `RESEARCH_INCLUDE` / `RESEARCH_EXCLUDE` 정규식이 처리.

- **포함 패턴**: 연구용역, 학술연구, 조사용역, 기초연구, 기획연구, 정책연구,
  타당성 (조사/분석/연구), 실태조사, 모니터링 연구, 연구개발, R&D,
  마스터플랜, 기본계획 수립, 중장기 계획.
- **제외 패턴**: 공사, 건설, 시공, 설치, 구입/구매, 임대/리스,
  유지보수, 정비, 청소, 경비, 급식, 장비 구매, 소프트웨어 개발/시스템 구축, 위탁운영 등.
- `RESEARCH_ONLY=0` 환경변수로 필터 OFF 가능 (워크플로에서 `RESEARCH_ONLY: "1"` 로 기본 켜짐).

### 2-4. 카카오 refresh_token 만료 (약 2개월)

만료 1개월 전부터 카카오 응답에 새 토큰이 포함되며, `scripts/kakao.py` 가 이를
`GITHUB_OUTPUT` 으로 노출함. (수동 갱신 액션 미연결 — 향후 작업)

완전 만료 시 1-3 절차 재실행 → `KAKAO_REFRESH_TOKEN` Secret 교체.

### 2-5. 비용

- GitHub Actions: public repo 무제한 / private repo 월 2,000분.
- GitHub Pages: 무료. 트래픽 100GB/월.
- 공공데이터포털: 일 10,000건 (오퍼레이션당 1,000건) 무료.
- 카카오 메시지 API: '나에게 보내기' 무료.

전체 = **무료**.

### 2-6. 트러블슈팅

| 증상 | 원인 / 조치 |
|---|---|
| `G2B_SERVICE_KEY 환경변수가 비어 있습니다` | Secret 이름 오타 또는 미등록 |
| `API ResponseError ... 30 (NO_OPENAPI_SERVICE_ERROR)` | 활용신청 미승인. 공공데이터포털 마이페이지 확인 |
| `KOE320` 카카오 응답 | refresh_token 만료. 1-3 재실행 |
| 카톡 메시지 안 옴 | Actions 로그에서 `kakao.py` 단계 종료코드 확인 |
| 대시보드가 비어 보임 | Pages 배포 1~2분 대기 또는 브라우저 캐시 비우기 |
| 학술연구가 아닌데 매칭됨 | `RESEARCH_EXCLUDE` 패턴에 제외어 추가 |

---

## 3. 폴더 구조

```
nara-weekly/
├── .github/workflows/weekly.yml   # 수·금 10시 cron
├── scripts/
│   ├── collect.py                 # OpenAPI 호출 + 학술연구 필터 + 키워드 매칭 → tenders.json (attachments 1~5)
│   ├── kakao.py                   # access_token 갱신 + '나에게 보내기'
│   ├── render.py                  # 사이드바 + Chart.js + SheetJS + 키워드 설정 UI 생성
│   └── get_initial_token.py       # 최초 OAuth (로컬 1회)
├── keywords.yml                   # 관심 키워드 + 가중치
├── requirements.txt
├── data/tenders.json              # 누적 원본 (first_seen_at 보존용)
└── docs/                          # GitHub Pages 루트
    ├── index.html                 # 대시보드 (사이드바 + 차트 + 엑셀 + 키워드 설정)
    ├── tenders.json               # 클라이언트가 fetch 하는 데이터
    └── keywords.yml               # render.py 가 ../keywords.yml 에서 미러링 (UI 미리보기용)
```

`data/` 와 `docs/` 둘 다에 `tenders.json` 이 있는 이유:
`data/` 는 누적 원본(다음 회 first_seen_at 보존), `docs/` 는 Pages 서빙용.
`docs/keywords.yml` 은 키워드 설정 뷰의 "현재 yml 미리보기" 용 (render.py 가 자동 미러링).

## 4. 관련 프로젝트

[Nara Radar](../nara-radar) — 같은 데이터 소스를 쓰는 FastAPI + PostgreSQL + Next.js 풀스택 사내 운영용 버전.
사용자가 늘어나거나 Slack/팀 공유가 필요해지면 그쪽으로.
