"""마스터 미디어 표준표 맞추기 — TOTAL_MASTER_MX/CE 의 Media Type 1 · 2 · Platform 을 D2C 표준표로.

    python tools/master_media.py           미리보기
    python tools/master_media.py --apply   마스터에 반영 (run_all 이 클렌징 뒤에 부른다)

클렌징 단계(src/media_std.py)와 같은 판정을 마스터에 쌓인 과거 행(이번 입력에 없는 법인 · 기간)에도 적용한다.
원본 값은 마스터의 'Media Type 2 (Raw)' · 'Media Platform (Raw)' 열에 그대로 있다. 다른 열 · 다른 행은 손대지 않는다.
"""
from __future__ import annotations

import argparse
import shutil
from datetime import datetime

import pandas as pd

from common import config
from src.media_std import PF, T1, T2, load_taxonomy, standardize_media

COLS = {"Media Type 1": T1, "Media Type 2": T2, "Media Platform": PF}


def fix_masters(apply: bool) -> int:
    tax = load_taxonomy(config.MEDIA_TAXONOMY)
    if tax.empty:
        return 0
    total = 0
    print("■ 마스터 미디어 표준표 맞추기" + ("" if apply else " (미리보기)"))
    for kind, path in config.MASTER_FILES.items():
        if not path.exists():
            continue
        m = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        work = m[list(COLS)].rename(columns=COLS)
        work["Media Spend (USD)"] = m.get("Media Spend (USD)", "0")
        fixed, summary = standardize_media(work, tax)
        changed = (fixed[list(COLS.values())].values != work[list(COLS.values())].values).any(axis=1)
        n = int(changed.sum())
        total += n
        print(f"  {kind}: 바꿈 {n:,}행")
        for r in summary.head(12).itertuples(index=False):
            print(f"     {r.결과:<24} {r.전}  →  {r.후}   {r.Counts:,}행 ${r.Spend:,.0f}")
        if apply and n:
            shutil.copy2(path, config.BACKUP_DIR / f"Backup_{kind}_{datetime.now():%Y%m%d_%H%M%S}_media.csv")
            for col, wc in COLS.items():
                m[col] = fixed[wc].values
            m.to_csv(path, index=False, encoding="utf-8-sig")
    return total


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--apply", action="store_true")
    fix_masters(ap.parse_args().apply)
