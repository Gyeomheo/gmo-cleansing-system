# GMO Cleansing

법인별 GMO 원본 CSV(MX · CE)를 매핑 엑셀 기준으로 클렌징하고 `TOTAL_MASTER_MX/CE.csv` 마스터에 누적한다.
**회차 실행에는 AI 가 필요 없다** — 더블클릭 한 번. Claude 는 코드를 고칠 때만 쓴다.
현재 상태 · 이력 · 남은 일은 [docs/작업이력.md](docs/작업이력.md).

## 0. 바로 시작

**GUI: 바탕화면 `Weekly Report` (또는 `GMO클렌징_GUI.cmd`)** — 복기(ROAS) 데이터 준비 GUI 와 같은 디자인. 바탕화면 `복기 (ROAS WIKI)` 는 ROAS 쪽 GUI — 주차(기본 = 이번 ISO 주차) · 취합 구분(최신 취합본이 오래된 쪽이 기본) 확인 → **[▶ 전체 실행]**
1. ① 수집 · 클렌징 (= `run_all.py --root <주차 폴더>`)
2. ② 취합 검증 (= `consolidate.py <CE|MX>`) — 기준 = 최신 취합본 RAW, 법인별로 값이 처음 달라진 날부터 교체
3. ③ 취합본 만들기 (= `consolidate.py <CE|MX> --apply`) — 이전 취합본은 `old RAW files` 로, 새 `<YYMMDD>_… W<nn>_v1_GMO_<CE|MX>_RAW File.xlsx`,
   완전 중복 행 제거, RAW 교체 · 피벗 원본 범위 · 새로고침, 저장 후 다시 읽어 대조. ② 에 오류가 있으면 ③ 은 안 함
[폴더 선택…] 으로 다른 주차 폴더도. 결과 버튼: OUTPUT · Reports · 실행요약 · 취합검증 · 취합 폴더.


**매주: `이번주_실행.cmd` 더블클릭** — 이번 ISO 주차 폴더(국가별 weekly\<nn>주차\MX · CE)의 법인 파일을 모아
(zip 풀기 · GMO 시트 찾기 · 열 이름 기준으로 맞추기 · 같은 종류는 최신 폴더 우선) 1_Input 을 새로 채우고, 아래 2~4 까지 자동.
다른 주차: `python tools/run_all.py --week 43`

손으로 받은 CSV 가 있을 때:

1. 받은 CSV 들을 **`클렌징.cmd` 위에 끌어다 놓기** (또는 `다운로드에서_넣고_실행.cmd` — 다운로드 폴더 최근 7일)
   - 파일 이름은 상관없다. **내용으로** 법인 · MX/CE 를 판별해서 1_Input 의 제자리에 넣는다
   - 같은 자리(법인 + MX/CE)의 이전 파일은 `1_Input/_보관/<시각>/` 로 옮긴다
2. 자동으로 점검 → 클렌징 → 매핑 제안 → 실행 요약까지 돈다 (약 20초)
3. 끝나면 `2_Output/Reports` 가 열린다
   - `실행요약.md` — 법인별 행 · 마지막 날 · Spend/Revenue 변화 · Other·Multi 로 빠진 Spend, 확인할 것
   - `매핑등록.xlsx` — 매핑 엑셀에 없는 원본 값(있을 때만)
4. 매핑을 고칠 게 있으면 `매핑등록.xlsx` 에서 표준1~3 확인 · 고치고 **'등록' 칸에 O** → 저장 · 닫기 → **`매핑등록.cmd`**
   - 매핑 엑셀에 추가(백업 후) → 미리보기 → 입력 기간 전체를 다시 클렌징해 마스터에 반영(`--remap`)

Excel 에서 매핑 엑셀 · 마스터 CSV 를 열어 두면 점검에서 멈춘다(저장 실패 방지).

## 1. 흐름

```
받은 CSV ──▶ 넣기(내용 판별) ──▶ 점검 ──▶ run.py 클렌징 ──▶ 매핑 제안 ──▶ 실행 요약
                                              │                   │
                                     cleaned_*.csv           매핑등록.xlsx ── O 표시 ──▶ 매핑등록.cmd
                                     TOTAL_MASTER_*.csv                                (등록 → --remap 재실행)
```

| 단계 | 하는 일 | 코드 |
|---|---|---|
| 넣기 | 열 이름으로 GMO 형식 확인, `BU` 열 있으면 CE, Subsidiary 최빈값 = 법인. 같은 자리는 기존 이름 유지 | `tools/inputs.py` |
| 점검 | 매핑 엑셀 · 마스터 잠금, 형식 아닌 파일, 같은 자리 중복, 마스터보다 새 날짜가 있는지 | `tools/check.py` |
| 클렌징 | 기존 그대로 (`run.py` → `src/`) | `run.py` |
| 매핑 제안 | 원본에서 **매핑 키에 없는 값**을 찾아 Spend 큰 순. 지금 어떻게 클렌징되는지 · 표준값 제안 | `tools/mapping.py propose` |
| 실행 요약 | 이번 실행이 만든 백업(실행 직전 마스터)과 비교 | `tools/summary.py` |

## 2. 캠페인명 역추적 (`src/campaign.py` · `3_Config/campaign_rules.xlsx`)

캠페인명 택소노미의 제품 코드로 비어 있거나 Multi · Others 인 제품 분류를 채운다. 예: `CN~d2c-ce-tv_PR~nq4k_SB~tv_…` → TV / Neo QLED.

- **원본 우선**: 매핑 클렌징 결과의 시리즈나 분류가 빈 값 · Multi · Other(s) 일 때만 채운다. 원본이 구체적인데 규칙과 다르면 바꾸지 않고 `cleaned_*_Summary_Campaign.csv` 에 '충돌(그대로)'로만 남긴다
- **번들 예외**: 키워드를 `+` 로 묶은 규칙(예: `s26|f1h26 + buds`)은 한 캠페인에 여러 제품이 같이 있는 경우라 원본보다 우선해 교체한다(→ Multi)
- 규칙 열: 사용(O) · 법인(비우면 전체) · 구분(MX/CE) · PR · SB · 키워드(`|` 하나, `+` 모두 — 캠페인명 · Product Series · Products 원본에서 찾음) · 제외 키워드 · Category · Series · Products · 시작일 · 원본보다 우선(O) · 근거
- 예: 2026-07-01 이후 Watch Ultra → Watch9 Series / Watch Ultra2 (Watch7 · 8 과 같이 나오면 제외), Watch9 Series 로 들어온 Watch Ultra → Watch Ultra2
- 실행마다 규칙 없는 코드는 `Reports/캠페인규칙제안.xlsx` 로 (같은 코드가 다른 행에서 들어온 값으로 제안). O 표시 → **`캠페인규칙등록.cmd`**
- 바뀐 행은 cleaned 파일 `Campaign_Rule` 열에 어떤 규칙인지 남는다. 끄려면 `src/config.py` 의 `CAMPAIGN_BACKTRACE = False`

## 1-1. Collection status Tracker · 메일

- **Tracker 시트**(취합본 맨 앞): 메일 양식(`D2C Weekly Report 취합_*.xlsx` › `3. 메일 작성`)의 표를 서식째 복사하고, 값은 **status 피벗** 수식.
  최근 2주(보고 W 이면 W-2 · W-1) Yes / -, 두 주 모두 - 면 Uncollected True, Note = 고정 사유 또는 Missing submission.
  법인 · Strategic · 추적(Y/N) · 고정 사유 = `3_Config/tracker_markets.csv`. 추적 N(제출 안 함) = 회색 줄, 주차 비움, 사유만.
  파일의 `고정 사유` 칸에 적으면 다음 주 취합본으로 이어짐. 취합본을 만들 때 자동(`tools/tracker.py`)
- **[④ 메일]** (`tools/weekly_mail.py --week nn --copy`): 본문(그 주 구분 표 하나 — 홀수 주 CE, 짝수 주 MX)을 클립보드에 넣고
  회사 Gmail(`?authuser=<회사 Gmail>`)의 **보낸편지함 최근 3주 'GMO Weekly Report'**(리마인드 제외, PCG 수신)를 연다.
  지난주 메일 → 전체 회신 → GUI **[제목 복사]**(`[CE] 2026 W43 GMO Weekly Report`) → 본문 Ctrl+V → **보내기는 직접**.
  지난 메일이 아래로 인용돼 업무 기록이 쌓인다. 실제 파일은 CPCex 로 따로.
  리마인드(그 주 구분만, 최근 2주 모두 '-', 범위 = 안 준 날짜가 속한 주 ~ 이번 주-1)는 `Reports/메일_W<nn>.html` 의 [본문 복사].
  법인별 회신할 스레드는 담당자가 직접 관리(`3_Config/reminder_links.csv` 는 비워 둠 — 링크를 넣으면 [메일 열기] 버튼이 생김).
- 자동 초안(선택): `tools/gas_mail/Code.gs` 를 Apps Script 에 넣고 `dryRun` → `installTrigger`. 이후 [④ 메일] 때마다 5분 안에
  그 스레드에 회신 초안 · 리마인드 초안이 생김(`내 드라이브/GMO_메일초안/` 경유)

## 2-0. 미디어 표준표 (`3_Config/media_taxonomy.csv`)

미디어는 D2C 표준표에 있는 조합만 쓴다. 클렌징 뒤 표에 없는 Platform → Others(DSP 는 Other DSPs), 표에 없는 Type 2 → Others / Others,
Affiliate Platform 은 원본 그대로. Type 1 이 표에 없으면 Type 2 로 찾고(Other / PerformanceMax → Shopping Ads),
Platform 이 비었거나 Other 류면 표의 `기본` 열 O 값(YouTube → Google), 없으면 그 Type 2 의 Platform 이 하나뿐일 때 그것으로(Performance Max · GDN · Demand Gen → Google). 띄어쓰기 · 대소문자 무시.
약어는 `3_Config/media_aliases.csv`(칸 · 별칭 · 정식 이름)로 먼저 푼다 — PMAX → Performance Max, FB · IG → Meta, TTD → The Trade Desk 등. 새 약어는 여기에 한 줄. 원본 값은 마스터 `Media Platform (Raw)` 에. 마스터 과거 행도 매 실행 같은 검사(`tools/master_media.py`).
미디어 매핑 엑셀에 새 플랫폼을 넣지 않는다 — 표를 고치려면 이 CSV 를.

## 2-1. 매핑 자동 추론 (`tools/auto_mapping.py`, run_all 이 클렌징 전에 실행)

매핑 키에 없는 원본 값 중 **법칙으로 확실한 것만** 매핑 엑셀에 자동 등록한다(메모 열 '자동: <법칙>'). 끄려면 `--no-auto`.

0. 별칭 풀기 — `3_Config/aliases.csv` (Q8 = Z Fold8 Ultra, H8 = Z Fold8, B8 = Z Flip8). 새 코드명은 여기에 한 줄
1. 띄어쓰기 · 대소문자 · 기호만 다른 기존 키 (Watch 9 = Watch9, Tab S10 Fe = Tab S10 FE)
2. 표준 제품명과 같음 (Z Fold 8 = Z Fold8)
3. 표준 시리즈명 안에 단어로 있고 후보 하나
4. 이전 cleaned 이력에서 90% 이상 한 결과로 갔고 그 결과가 표준값

일반값(Other · Multi · Tablet) · 후보 여럿 · 3글자 미만은 자동 등록 안 하고 `매핑등록.xlsx` 제안으로만.

## 3. 매핑 한 줄 등록 (Claude 에게 부탁하던 것)

```bash
python tools/mapping.py add MX "S26 FE" Smartphones "S26 Series" "S26 FE"
python tools/mapping.py add CE "Washing Machine" Washer Bespoke NONE
python tools/mapping.py add MEDIA "Pinterest" "Paid Social" Others Others
python tools/mapping.py test MX --cat Smartphones --series "S26 Series" --value "S26 FE"
```

- 원본 값: MX = `Products`, CE = `Product Category`, MEDIA = `Media Platform`(Media Type 1 이 Affiliate 면 `Media Type 2`)
- 키(C_Key · A_Key · F_Key)는 파이프라인과 같은 정규화로 자동 계산한다 — 예: `S26 FE` → `S26 Fe`, `TikTok` → `Tik Tok`.
  손으로 넣으면 이 정규화를 틀리기 쉽다
- 등록 전 매핑 엑셀은 `3_Config/_backup/` 에 백업
- 등록 뒤 마스터에 반영하려면 `python tools/run_all.py --remap`

## 4. 명령 모음

| 명령 | 내용 |
|---|---|
| `python tools/run_all.py [파일 · 폴더 ...]` | 넣기 → 점검 → 클렌징 → 제안 → 요약 |
| `python tools/run_all.py --downloads` | 다운로드 폴더 최근 7일 CSV 중 GMO 형식만 넣고 실행 |
| `python tools/run_all.py --check-only` | 넣기 · 점검까지만 |
| `python tools/run_all.py --remap` | 매핑 변경 반영: 숫자가 같아도 입력 기간 전체를 마스터에 다시 씀 |
| `python tools/inputs.py` | 1_Input 현황 |
| `python tools/campaign_rules.py propose / apply` | 캠페인 규칙 제안표 만들기 / O 표시한 것 규칙표에 넣기 |
| `python run.py` | 예전 방식 그대로(클렌징만) |

Python: `%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe` (없으면 `py`).
