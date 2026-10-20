"""점검 — 실행 전에 막힐 일을 먼저 찾는다.

    python tools/check.py

- 매핑 엑셀 · 마스터 CSV 가 Excel 에서 열려 있지 않은지(열려 있으면 저장 실패)
- 1_Input: GMO 형식 · 슬롯 중복 · 마지막 날이 지난 일요일보다 이른지
- 매핑 엑셀: 키 열 · 표준 열이 있는지, 같은 키에 같은 행이 두 번 있는지
"""
from __future__ import annotations

import pandas as pd

from common import config, is_locked
from inputs import status

MAPS = [
    (config.MEDIA_MAP_EXCEL, config.MEDIA_COLS_MAP),
    (config.PRODUCT_MX_EXCEL, config.PRODUCT_COLS_MAP_MX),
    (config.PRODUCT_CE_EXCEL, config.PRODUCT_COLS_MAP_CE),
]


def check_maps() -> tuple[list[str], list[str]]:
    errors, notes = [], []
    for path, cols in MAPS:
        if not path.exists():
            errors.append(f"매핑 엑셀 없음: {path.name}")
            continue
        if is_locked(path):
            errors.append(f"{path.name} 이(가) Excel 에서 열려 있음 — 닫고 다시")
            continue
        df = pd.read_excel(path).fillna("")
        need = [cols["key"]] + cols["std_cols"]
        miss = [c for c in need if c not in df.columns]
        if miss:
            errors.append(f"{path.name}: 열 없음 {miss}")
            continue
        dup = df.duplicated(subset=need).sum()
        if dup:
            notes.append(f"{path.name}: 완전히 같은 행 {dup}개(동작엔 영향 없음)")
    return errors, notes


def check_masters() -> list[str]:
    return [f"{p.name} 이(가) Excel 에서 열려 있음 — 닫고 다시"
            for p in config.MASTER_FILES.values() if is_locked(p)]


def run_check() -> bool:
    print("■ 점검")
    errors, notes = check_maps()
    errors += check_masters()
    input_problems = status()
    # 입력 문제 중 '형식 아님' · '중복'은 실행을 막는다. 기간 경고는 표에만.
    errors += input_problems
    for n in notes:
        print(f"  · {n}")
    if errors:
        print("\n  점검 실패:")
        for e in errors:
            print(f"  ✖ {e}")
        return False
    print("\n  ✔ 점검 통과")
    return True


if __name__ == "__main__":
    raise SystemExit(0 if run_check() else 1)
