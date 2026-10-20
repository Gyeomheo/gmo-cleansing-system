# GMO Cleansing — Claude 작업 안내

- 먼저 [docs/작업이력.md](docs/작업이력.md)(현재 상태 · 확정 규칙 · 남은 일)와 [README.md](README.md)(운영)를 읽는다.
- 답변은 한국어로.
- **기존 클렌징 로직(`src/pipeline.py`) + 알파. 대체가 아니다.** 담당자 말: 옛 LLM 으로 뽑아낸 역작. 새 기능은 그 결과 위에 얹는다
  (매핑은 매핑 엑셀에 행 추가, 캠페인 보정은 클렌징 뒤 빈 값 · Multi · Others 만 채움). pipeline.py 로직을 바꿔야 하면 먼저 묻는다.
- 회차 실행은 코드만 돈다(`클렌징.cmd` → `tools/run_all.py`). 대화로만 처리한 단계를 남기지 않는다 — 새 수동 단계가 보이면 `tools/` 에 코드로 박제하고 README · 작업이력에 적는다.

## 매핑 등록 요청이 오면

"X 를 A / B / C 로 등록" 류는 엑셀을 직접 고치지 말고 명령으로:

```bash
python tools/mapping.py add MX "<원본 Products>" "<Category>" "<Series>" "<Products>"
python tools/mapping.py add CE "<원본 Product Category>" "<Category>" "<Series>" NONE
python tools/mapping.py add MEDIA "<원본 Media Platform>" "<Media Type 1>" "<Media Type 2>" "<Platform>"
```

- 키 정규화 · 백업 · 중복 확인 · 미리보기를 도구가 한다. 끝나면 미리보기 결과가 요청과 같은지 확인해 보고한다.
- 마스터에 반영하려면 `python tools/run_all.py --remap` (매핑만 바뀌면 일반 실행으로는 마스터가 안 바뀐다).
- Python: `%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe`

## 고칠 때

- `src/pipeline.py` 로직을 바꾸면 고치기 전 결과와 대조한다: 저장소를 스크래치에 복사 → 전/후 각각 실행 → `2_Output/*.csv` 를 `cmp` 로 비교.
- 실제 `2_Output` 마스터를 시험 삼아 돌리지 않는다(실행할 때마다 마스터가 갱신되고 백업이 쌓인다).
