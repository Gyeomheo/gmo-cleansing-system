# GMO Cleansing

법인별 GMO 원본(MX · CE)을 모아 매핑 엑셀 기준으로 클렌징하고, `TOTAL_MASTER_MX/CE.csv` 마스터에 누적한 뒤
공유 드라이브 **D2C 취합본**(RAW · 피벗 · Collection status Tracker)과 주간 메일까지 만든다.
**회차 실행에는 AI 가 필요 없다** — 버튼 한 번. Claude 는 코드를 고칠 때만 쓴다.
현재 상태 · 이력 · 확정 규칙 · 남은 일은 [docs/작업이력.md](docs/작업이력.md), 클렌징 코어(`src/pipeline.py`) 단계별 설명은 [docs/클렌징로직.md](docs/클렌징로직.md).

- 기존 클렌징 로직(`src/pipeline.py`)은 그대로 두고 그 위에 얹는다(+ 알파). 기본 경로 결과는 자동화 전 `run.py` 와 바이트 단위로 같다
- 2026 년만 추적한다(`src/config.py` `MIN_DATE = '2026-01-01'`)

### 한눈에 — 파일이 어디서 어디로 가나

```mermaid
flowchart LR
  subgraph SD["공유 드라이브 · 09. GMO Report"]
    W["2. 국가별 weekly<br/>nn주차 / MX·CE / 법인 MMDD"]
    D["3. Weekly Dashboard RAW File / D2C<br/>Wnn 취합본 (RAW · 피벗 · Tracker)"]
    O["old RAW files"]
  end
  subgraph PC["이 PC · C:/GMO_Cleansing"]
    I["1_Input"]
    C["클렌징<br/>기존 로직 + 알파"]
    M["2_Output<br/>TOTAL_MASTER · cleaned"]
    R["2_Output/Reports<br/>실행요약 · 취합검증 · 메일"]
  end
  G["Gmail<br/>지난주 메일에 전체 회신"]
  W -->|"① 수집"| I --> C --> M
  C --> R
  M -->|"② 검증 · ③ 취합"| D
  D -.->|"이전 취합본"| O
  D -->|"④ 메일 (보내기는 직접)"| G
```

## 0. 바로 시작 — 바탕화면 `Weekly Report`

(= `GMO클렌징_GUI.cmd`) 머리줄에서 **주차**(기본 = 이번 ISO 주차)와 **구분**(홀수 주 CE · 짝수 주 MX)을 확인 → **[▶ 전체 실행]**.
Excel 은 모두 닫고 돌린다.

![Weekly Report GUI](docs/img/gui.png)

```mermaid
flowchart LR
  A["① 수집 · 클렌징"] --> B{"② 취합 검증"}
  B -->|"오류 없음"| C["③ 취합본 만들기<br/>+ Tracker"]
  B -->|"오류"| X["멈춤 — Reports/취합검증_구분.md 확인"]
  C --> D["④ 메일<br/>본문 복사 · 지난 메일 열기"]
  D --> E(["보내기는 담당자가 직접"])
```

| 단계 | 하는 일 | 같은 명령 |
|---|---|---|
| ① 수집 · 클렌징 | 국가별 weekly `<nn>주차\<MX\|CE>\<법인 MMDD>` 에서 원본 수집 → 1_Input → 클렌징 · 마스터 갱신 | `run_all.py --root <주차 폴더>` |
| ② 취합 검증 | 최신 취합본 RAW + 이번 실행 결과를 합쳐 보고 `Reports/취합검증_<구분>.md` | `consolidate.py <CE\|MX>` |
| ③ 취합본 만들기 | 이전 취합본 → `old RAW files`, 새 `<YYMMDD>_Samsung D2C (2026) - W<nn>_v1_GMO_<구분>_RAW File.xlsx` + Tracker. ② 에 오류가 있으면 안 함. 같은 주차 파일이 있으면 새 버전을 만들지 먼저 묻는다 | `consolidate.py <CE\|MX> --apply` |
| ④ 메일 | 주간 메일 본문을 클립보드에 넣고 지난주 보낸 메일을 연다 · 리마인드 HTML | `weekly_mail.py --week <nn> --copy` |

- 왼쪽: **이번 주 원본**(법인별 ● 있음 ○ 없음 – 제출 안 함) · **작업 순서**(파일을 보고 자동 체크, '← 지금')
- 오른쪽: 원본 폴더 · 취합 폴더 · Gmail 열 곳(저장 · 열기), [제목 복사], 결과 버튼(OUTPUT · Reports · 실행요약 · 취합검증 · 취합 폴더), 로그
- [폴더 선택…] 으로 다른 주차 폴더도 돌릴 수 있다

GUI 없이: `이번주_실행.cmd`(① 만), 손으로 받은 CSV 는 `클렌징.cmd` 위에 끌어다 놓기 또는 `다운로드에서_넣고_실행.cmd`(다운로드 폴더 최근 7일).

## 1. 매주 흐름

**기존 로직 + 알파** — 클렌징 코어는 손대지 않고, 앞뒤에만 얹는다.

```mermaid
flowchart TB
  subgraph PRE["클렌징 전 · 알파"]
    P1["수집 보정<br/>머리줄 별칭 · 밀린 열 · S.com / GMO / DA"]
    P2["매핑 자동 추론<br/>확실한 것만 매핑 엑셀에 행 추가"]
  end
  subgraph CORE["기존 로직 · src/pipeline.py — 그대로"]
    K["날짜 · 숫자 · 미디어 · 제품 클렌징<br/>(매핑 엑셀 기준)"]
  end
  subgraph POST["클렌징 뒤 · 알파 — 빈 값 · Multi · Others 만 손댐"]
    Q1["캠페인 역추적"]
    Q2["미디어 표준표"]
  end
  PRE --> CORE --> POST --> U["마스터 갱신<br/>법인별 숫자가 처음 달라진 날부터 교체"]
  U --> S["실행 요약 · 매핑 제안"]
  S -.->|"제안에 O 표시 → 매핑등록.cmd"| P2
```

| 단계 | 하는 일 | 코드 |
|---|---|---|
| 수집 | 법인 폴더마다 최신 파일(zip 은 풀어서), GMO 시트는 날짜가 가장 최근인 시트. 머리줄 별칭(`Div`→BU · `Mon`→Month · `Local`→Funding), 밀려 들어간 미디어 열은 행 단위로 바로잡음, 값 표준화(Sales Channel · Partner 의 Epp · app · web → `S.com`, Funding `HQ` → `GMO`, BU `HA` → `DA`). DCT · DCT Billing 폴더는 대상 아님 | `tools/collect.py` |
| 넣기 | 열 이름으로 GMO 형식 확인, `BU` 열 있으면 CE, Subsidiary 최빈값 = 법인. 같은 자리(법인 + 구분)의 이전 파일은 `1_Input/_보관/<시각>/` | `tools/inputs.py` |
| 점검 | 매핑 엑셀 · 마스터 잠금, 형식 아닌 파일, 같은 자리 중복, 마스터보다 새 날짜가 있는지 | `tools/check.py` |
| 매핑 자동 추론 | 매핑 키에 없는 값 중 법칙으로 확실한 것만 매핑 엑셀에 등록(§3) | `tools/auto_mapping.py` |
| 클렌징 | 기존 그대로 (`run.py` → `src/`) + 캠페인 역추적(§4) + 미디어 표준표(§5) | `run.py` |
| 마스터 갱신 | 법인별로 숫자(Spend · Revenue · Imp · Click · Order)가 **처음 달라진 날**부터 교체. 분류만 바뀌면 `--remap` | `src/updater.py` |
| 매핑 제안 · 요약 | 원본에서 매핑 키에 없는 값(Spend 큰 순), 실행 직전 마스터와 비교한 법인별 변화 | `tools/mapping.py propose` · `tools/summary.py` |
| 매핑 점검(경고만, 취합 때) | 기준 = 취합본(진짜 마스터). ④ 출시 주보다 4주 넘게 앞선 신제품 Spend ⑤ 이전 취합본과 숫자까지 같은 행을 짝지어 분류가 바뀐 것 — 전 → 후 분류 · 원본 입력값(이번 cleaned 파일) · 캠페인, 행 단위는 `매핑점검_<구분>_분류바뀜.csv` ⑥ 법인 · 주차별 Multi · Other 비중이 앞 8주보다 15%p↑ · $20K↑(캠페인명 상위 3개 첨부) → `Reports/매핑점검_<구분>.md`, 취합검증에 건수. 따로: `map_check.py CE [--old 이전취합본] [--all-weeks]` | `tools/map_check.py` |

결과: `2_Output/Reports/` — `실행요약.md` · `매핑등록.xlsx`(있을 때만) · `캠페인규칙제안.xlsx` · `취합검증_<구분>.md` · `Tracker_<구분>_W<nn>.html` · `메일_W<nn>.html`.

## 2. 취합 (`tools/consolidate.py`)

**법인별 교체 규칙** — 겹치는 기간에서 숫자가 처음 달라진 날부터 새 것으로 바꾸고, 그 앞과 다른 법인은 손대지 않는다.

```text
              10/01               10/12               10/18
SEPR   RAW    ###########################
       NEW    ....................#####################   <- 10/12 부터 숫자가 다름
       OUT    ####################=====================   (= : 새 데이터로 교체)

SEAU   RAW    ###########################
       NEW    ...........................######           <- 겹치는 날은 숫자가 같음
       OUT    ###########################======           (뒤에 이어 붙이기만)

SEI    RAW    ##################                          <- 이번에 안 냄
       OUT    ##################                          (그대로)
```

```mermaid
flowchart LR
  A["최신 취합본 RAW"] --> M["법인별 합치기"]
  B["_이번실행_구분.csv"] --> M
  M --> C["완전 중복 제거 · 미디어 표준표 ·<br/>CPC 재계산 · 표기 통일"]
  C --> V{"검증"}
  V -->|"오류"| X["멈춤"]
  V -->|"통과"| W["Excel 로 쓰기<br/>RAW 교체 · 피벗 캐시 1개 · 새로고침"]
  W --> R{"다시 읽어<br/>모든 행 · 열 대조"}
  R -->|"같음"| T["Tracker 시트 · 이전 파일 → old RAW files"]
```


- 기준 = 공유 드라이브 `3. Weekly Dashboard RAW File\D2C` 의 최신 취합본 RAW. 법인별로 **값이 처음 달라진 날부터** 이번 실행 결과로 교체, 나머지 법인 · 날짜는 그대로
- 완전 중복 행 제거, 미디어 표준표 재검사, CPC = `IFERROR(Spend / Clicks, 0)` 다시 계산, 표기 통일
- 검증(오류면 ③ 안 함): 안 바뀐 법인 보존 · 법인별 교체 전후 · 변화 큰 법인(경고) · CPC · 날짜(미래 · 범위) · 법인 빈 값 · 미디어 표준표 밖 · 중복 행
- 쓰기는 Excel COM: 날짜는 Excel 일련번호 + `yyyy-mm-dd`(datetime 을 넘기면 9시간 당겨짐), 피벗은 **새 캐시 하나를 모든 피벗이 공유**하고 새로고침
- 저장 후 다시 읽어 **모든 열 · 모든 행** 대조. 이전 취합본은 `old RAW files` 로

## 3. 매핑 자동 추론 (`tools/auto_mapping.py`, run_all 이 클렌징 전에 실행)

매핑 키에 없는 원본 값 중 **법칙으로 확실한 것만** 매핑 엑셀에 자동 등록한다(메모 열 '자동: <법칙>'). 끄려면 `--no-auto`. 미디어는 대상 아님(§5).

```mermaid
flowchart TD
  OK["매핑 엑셀에 자동 등록<br/>메모: 자동 · 법칙 번호"]
  P["Reports/매핑등록.xlsx 제안<br/>사람이 확인 후 O"]
  S["매핑 키에 없는 원본 값"] --> G{"일반값 · 3글자 미만?<br/>Other · Multi · Tablet"}
  G -->|"예"| P
  G -->|"아니오"| R0{"0 · 별칭표에 있나<br/>Q8 · H8 · B8"}
  R0 -->|"예"| OK
  R0 -->|"아니오"| R1{"1 · 기존 키와 띄어쓰기 ·<br/>대소문자만 다름"}
  R1 -->|"예"| OK
  R1 -->|"아니오"| R2{"2 · 표준 제품명과 같음"}
  R2 -->|"예"| OK
  R2 -->|"아니오"| R3{"3 · 표준 시리즈명 안에<br/>단어로 · 후보 하나"}
  R3 -->|"예"| OK
  R3 -->|"아니오"| R4{"4 · 과거 이력 90% 이상<br/>한 표준값으로"}
  R4 -->|"예"| OK
  R4 -->|"아니오"| P
```


0. 별칭 풀기 — `3_Config/aliases.csv` (Q8 = Z Fold8 Ultra, H8 = Z Fold8, B8 = Z Flip8). 새 코드명은 여기에 한 줄
1. 띄어쓰기 · 대소문자 · 기호만 다른 기존 키 (Watch 9 = Watch9, Tab S10 Fe = Tab S10 FE)
2. 표준 제품명과 같음 (Z Fold 8 = Z Fold8)
3. 표준 시리즈명 안에 단어로 있고 후보 하나
4. 이전 cleaned 이력에서 90% 이상 한 결과로 갔고 그 결과가 표준값

일반값(Other · Multi · Tablet) · 후보 여럿 · 3글자 미만은 자동 등록 안 하고 `매핑등록.xlsx` 제안으로만.
제안을 고치려면 표준1~3 확인 · 고치고 **'등록' 칸에 O** → 저장 · 닫기 → **`매핑등록.cmd`**(등록 → 미리보기 → `--remap`).

## 4. 캠페인명 역추적 (`src/campaign.py` · `3_Config/campaign_rules.xlsx`)

캠페인명 택소노미의 제품 코드로 비어 있거나 Multi · Others 인 제품 분류를 채운다. 예: `CN~d2c-ce-tv_PR~nq4k_SB~tv_…` → TV / Neo QLED.

```mermaid
flowchart TD
  S["클렌징 끝난 행"] --> Q1{"맞는 규칙이 있나<br/>법인 · 구분 · PR · SB · 키워드 · 제외 · 시작일"}
  Q1 -->|"없음"| K["그대로"]
  Q1 -->|"있음"| Q2{"번들(+) 규칙 또는<br/>원본보다 우선 = O"}
  Q2 -->|"예"| REP["교체 (예: S26 + Buds → Multi)"]
  Q2 -->|"아니오"| Q3{"지금 분류가<br/>빈 값 · Multi · Others"}
  Q3 -->|"예"| FILL["채움 → Campaign_Rule 열에 기록"]
  Q3 -->|"아니오"| CF["그대로 · Summary_Campaign 에 '충돌'"]
```


- **원본 우선**: 클렌징 결과의 시리즈나 분류가 빈 값 · Multi · Other(s) 일 때만 채운다. 원본이 구체적인데 규칙과 다르면 바꾸지 않고 `cleaned_*_Summary_Campaign.csv` 에 '충돌(그대로)'로만 남긴다
- **번들 예외**: 키워드를 `+` 로 묶은 규칙(예: `s26|f1h26 + buds`)과 '원본보다 우선 = O' 규칙은 원본을 교체한다(→ Multi 등)
- 규칙 열: 사용(O) · 법인(비우면 전체) · 구분(MX/CE) · PR · SB · 키워드(`|` 하나, `+` 모두 — 캠페인명 · Product Series · Products 원본에서 찾음) · 제외 키워드 · Category · Series · Products · 시작일 · 원본보다 우선(O) · 근거
- 키워드가 `#` 로 시작하면 **코드 단어**: 앞뒤가 글자 · 숫자가 아닐 때만(코드끼리 붙은 `Q8H8B8` 은 됨). `#h8` 은 `f2h26-h8` 에 맞고 `Watch8` 엔 안 맞는다
- SB 만 있는 규칙도 된다(PR 규칙이 안 맞을 때 마지막으로). ROAS WIKI `ingestion._SB_TO_PRODUCT` 를 GMO 표준값으로 옮김: refrig → Refrigerator/Others, washmach → Washer/Others, aircon → Heating & Cooling/Multi, mon → Monitor/Multi, tv → TV/Others, wearsmart → Wearables/Multi, tab → Tablets/Multi …(규칙표 '근거' 열에 ROAS WIKI)
- Q8 · H8 · B8(담당자 2026-10-21): 단독이면 Q8 → Z Fold8 Ultra, H8 → Z Fold8, B8 → Z Flip8, B8 + Q8/H8 → Z Fold8 & Z Flip8. 모두 원본 우선
- 예: f1h26 · gs26se = S26 Series, 2026-07-01 이후 Watch Ultra → Watch9 Series / Watch Ultra2(Watch7 · 8 과 같이 나오면 제외), Watch9 Series 로 들어온 Watch Ultra → Watch Ultra2
- 택소노미(`CN~…_PR~…_SB~…`)는 SETK · SEAU CE · SEPR 만. SEC · SEDA · SIEL 은 자유 텍스트라 키워드 규칙만 먹는다
- 실행마다 규칙 없는 코드는 `Reports/캠페인규칙제안.xlsx` 로. O 표시 → **`캠페인규칙등록.cmd`**
- 바뀐 행은 cleaned 파일 `Campaign_Rule` 열에 남는다. 끄려면 `src/config.py` 의 `CAMPAIGN_BACKTRACE = False`

## 5. 미디어 표준표 (`3_Config/media_taxonomy.csv`)

미디어는 D2C 표준표에 있는 조합만 쓴다. **미디어 매핑 엑셀에 새 플랫폼을 넣지 않는다** — 표를 고치려면 이 CSV 를.

```mermaid
flowchart TD
  S["Media Type 1 · Type 2 · Platform"] --> A["약어 풀기<br/>PMAX → Performance Max · FB → Meta"]
  A --> B{"Type 1 이 표에 있나"}
  B -->|"아니오"| B2["Type 2 로 Type 1 찾기<br/>Other / PerformanceMax → Shopping Ads"]
  B -->|"예"| C
  B2 --> C{"Platform 이<br/>비었거나 Other 류"}
  C -->|"예"| D["표의 기본값 (YouTube → Google)<br/>또는 그 Type 2 의 유일한 Platform"]
  C -->|"아니오"| E
  D --> E{"표에 있는 조합"}
  E -->|"예"| OK["그대로"]
  E -->|"아니오"| OT["Others · Other DSPs<br/>원본은 Media Platform (Raw)"]
```


- 표에 없는 Platform → Others(DSP 는 Other DSPs), 표에 없는 Type 2 → Others / Others, Affiliate Platform 은 원본 그대로
- Type 1 이 표에 없으면 Type 2 로 찾는다(Other / PerformanceMax → Shopping Ads)
- Platform 이 비었거나 Other 류면 표의 `기본` 열 O 값(YouTube → Google), 없으면 그 Type 2 의 Platform 이 하나뿐일 때 그것(Performance Max · GDN · Demand Gen → Google). 띄어쓰기 · 대소문자 무시
- 약어는 `3_Config/media_aliases.csv`(칸 · 별칭 · 정식 이름)로 먼저 푼다 — PMAX → Performance Max, FB · IG → Meta, TTD → The Trade Desk 등
- 원본 값은 마스터 `Media Platform (Raw)` 에. 마스터 과거 행도 매 실행 같은 검사(`tools/master_media.py`)

## 6. Collection status Tracker · 메일

```mermaid
flowchart TD
  T["법인 · 3_Config/tracker_markets.csv"] --> N{"추적 = N<br/>(제출 안 함)"}
  N -->|"예"| GR["회색 줄 · 고정 사유만"]
  N -->|"아니오"| Y{"최근 2주 W-2 · W-1<br/>하나라도 Yes"}
  Y -->|"예"| OK["Collected · 리마인드 안 함"]
  Y -->|"아니오"| U["Uncollected True · Missing submission"]
  U --> RM["리마인드 대상 (그 주 구분만)<br/>안 준 날짜가 속한 주 ~ 이번 주-1"]
```

```mermaid
sequenceDiagram
  actor U as 담당자
  participant G as Weekly Report GUI
  participant M as Gmail
  U->>G: ④ 메일
  G->>G: Tracker HTML → 본문을 클립보드에
  G->>M: 보낸편지함 검색 열기 (최근 3주 GMO Weekly Report)
  U->>M: 지난주 메일 → 전체 회신
  U->>G: [제목 복사]
  U->>M: 제목 붙여넣기 · 본문 Ctrl+V · 보내기
  Note over M: 지난 메일이 아래로 인용 → 한 스레드에 업무 기록
```


- **Tracker 시트**(취합본 맨 앞, `tools/tracker.py`): 메일 양식(`D2C Weekly Report 취합_*.xlsx` › `3. 메일 작성`)의 표를 서식째 복사하고, 값은 **status 피벗** `GETPIVOTDATA`.
  - 최근 2주(보고 W 이면 W-2 · W-1) Yes / -, 두 주 모두 - 면 Uncollected True, Note = 고정 사유 또는 Missing submission. Collection status 칸만 테두리
  - 법인 · Strategic · Agency · 추적(Y/N) · 고정 사유 = `3_Config/tracker_markets.csv`. 추적 N(제출 안 함) = 회색 줄, 주차 비움, 사유만
  - 표 HTML 은 `Reports/Tracker_<구분>_W<nn>.html` 에도 저장 — 메일은 이걸 읽어서 Excel 없이 빠르게 만든다
- **주간 메일** ([④ 메일]): 본문 = 그 주 구분 표 하나(홀수 주 CE, 짝수 주 MX), W · 구분 굵게, 파일명은 굵은 글머리표
  1. 본문이 클립보드에 들어가고, 회사 Gmail **보낸편지함 최근 3주 'GMO Weekly Report'**(리마인드 제외, PCG 수신) — 계정 · PCG 주소는 `3_Config/mail_private.json`(git 제외, 이 PC 에만) 검색이 열린다
  2. 지난주 메일 → 전체 회신 → GUI **[제목 복사]**(`[CE] 2026 W43 GMO Weekly Report`) → 본문 Ctrl+V → **보내기는 직접**
  3. 지난 메일이 아래로 인용돼 업무 기록이 한 스레드에 쌓인다. 실제 파일은 CPCex 로 따로
- **리마인드**: `Reports/메일_W<nn>.html` 의 [본문 복사]. 그 주 구분만, 최근 2주 중 하나라도 Yes 면 제외, 범위 = 안 준 날짜가 속한 주 ~ 이번 주-1(주간 합계로 내는 법인은 마지막 행의 주까지 낸 것으로 봄). 법인별 회신 스레드는 담당자가 직접 관리(`3_Config/reminder_links.csv` 는 비워 둠 — 링크를 넣으면 [메일 열기] 버튼이 생김)
- 자동 초안(선택, 미설정): `tools/gas_mail/Code.gs` 를 Apps Script 에 넣고 `dryRun` → `installTrigger`. [④ 메일] 때마다 `내 드라이브/GMO_메일초안/` 경유로 5분 안에 회신 초안 · 리마인드 초안

## 7. 매핑 한 줄 등록 (Claude 에게 부탁하던 것)

```mermaid
flowchart LR
  A["매핑등록.xlsx 에 O<br/>또는 mapping.py add"] --> B["키 정규화 · 중복 확인<br/>매핑 엑셀 백업"]
  B --> C["매핑 엑셀에 행 추가"] --> D["미리보기"] --> E["run_all.py --remap<br/>마스터에 반영"]
```


```bash
python tools/mapping.py add MX "S26 FE" Smartphones "S26 Series" "S26 FE"
python tools/mapping.py add CE "Washing Machine" Washer Bespoke NONE
python tools/mapping.py test MX --cat Smartphones --series "S26 Series" --value "S26 FE"
```

- 원본 값: MX = `Products`, CE = `Product Category`
- 키(C_Key · A_Key)는 파이프라인과 같은 정규화로 자동 계산한다 — 예: `S26 FE` → `S26 Fe`. 손으로 넣으면 틀리기 쉽다
- 등록 전 매핑 엑셀은 `3_Config/_backup/` 에 백업. 마스터에 반영하려면 `python tools/run_all.py --remap`(법인만: `--remap SEPR,SETK`)
- 미디어는 매핑을 추가하지 않고 표준표(§5)를 고친다

## 8. 명령 모음

| 명령 | 내용 |
|---|---|
| `python tools/gui.py` | GUI (바탕화면 `Weekly Report`) |
| `python tools/run_all.py --week now\|<nn>` | 국가별 weekly 폴더에서 수집 → 점검 → 클렌징 → 제안 → 요약 |
| `python tools/run_all.py --root <주차 폴더>` | 폴더 직접 지정(주차 번호는 폴더 이름에서) |
| `python tools/run_all.py [파일 · 폴더 ...]` · `--downloads` | 손으로 받은 CSV · 다운로드 폴더 최근 7일 |
| `python tools/run_all.py --check-only` · `--no-auto` · `--open` | 점검까지만 · 매핑 자동 추론 끄기 · 끝나면 Reports 열기 |
| `python tools/run_all.py --remap [SEPR,SETK]` | 매핑 변경 반영: 숫자가 같아도 입력 기간 전체를 마스터에 다시 씀 |
| `python tools/collect.py <nn> <CE\|MX> [--put] [--root]` | 수집만(보기 · 1_Input 에 넣기) |
| `python tools/consolidate.py <CE\|MX> [--apply] [--week nn] [--dash-dir 시험 폴더]` | 취합 검증 · 취합본 만들기 |
| `python tools/tracker.py <CE\|MX> --week <nn> [--file]` | Tracker 만 다시(기본: 최신 취합본) |
| `python tools/weekly_mail.py --week <nn> [--copy] [--copy-subject] [--no-open]` | 메일 · 리마인드 |
| `python tools/mapping.py retire MX "S24 Series" "Other S Series"` | 세대가 밀린 시리즈를 Other 로(그 시리즈 행 전부, 백업 후). 과거 행까지는 다음 실행 `--remap` |
| `python tools/campaign_rules.py propose / apply` | 캠페인 규칙 제안표 / O 표시한 것 규칙표에 넣기 |
| `python tools/inputs.py` | 1_Input 현황 |
| `python run.py` | 예전 방식 그대로(클렌징만) |

Python: `%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe` (없으면 `py`).
Excel 에서 매핑 엑셀 · 마스터 CSV · 취합본을 열어 두면 점검에서 멈춘다(저장 실패 방지).

## 9. 설정 · 폴더

| 위치 | 내용 |
|---|---|
| `3_Config/` 매핑 엑셀 | MX · CE · MEDIA 매핑(기존). 백업 `3_Config/_backup/` |
| `3_Config/aliases.csv` | 제품 코드명 별칭(매핑 자동 추론) |
| `3_Config/campaign_rules.xlsx` | 캠페인 역추적 규칙 |
| `3_Config/media_taxonomy.csv` · `media_aliases.csv` | 미디어 표준표 · 약어 |
| `3_Config/tracker_markets.csv` | Tracker 법인 목록 · 추적 여부 · 고정 사유 |
| `3_Config/reminder_links.csv` | 리마인드 [메일 열기] 링크(지금은 비움) |
| `2_Output/` | `TOTAL_MASTER_*.csv` · cleaned · `_이번실행_<구분>.csv` · `Reports/` · 마스터 백업 `Backup/`(파일명 시각 기준) |
| `1_Input/` · `1_Input/_보관/` | 이번 입력 · 바뀐 입력 보관 |
| `assets/icon_weekly_v5.ico` · `tools/icons/make_icons.py` | 바로가기 · 작업표시줄 아이콘(W.) — 그림을 바꾸면 파일 이름 버전을 올린다 |

GUI · 아이콘 디자인 서식(Weekly ROAS 복기와 공통)은 [docs/작업이력.md](docs/작업이력.md) §3-2.
