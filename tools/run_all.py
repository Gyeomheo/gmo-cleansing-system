"""한 번에 — 넣기 → 점검 → 매핑 자동 추론 → 클렌징(run.py, 캠페인명 역추적 포함) → 매핑 · 캠페인 규칙 제안 → 실행 요약.

    python tools/run_all.py                 1_Input 에 있는 그대로 실행
    python tools/run_all.py 파일.csv ...     받은 CSV 를 판별해 넣고 실행 (클렌징.cmd 에 끌어다 놓기와 같음)
    python tools/run_all.py --downloads     다운로드 폴더 최근 7일 GMO CSV 를 넣고 실행
    python tools/run_all.py --check-only    넣기 · 점검까지만
    python tools/run_all.py --remap         매핑을 고친 뒤 — 입력 기간 전체를 마스터에 다시 씀(매핑등록.cmd 가 씀)
    python tools/run_all.py --remap SEPR    그 법인만(쉼표로 여러 개)
    python tools/run_all.py --week now      이번 ISO 주차 폴더(국가별 weekly/<nn>주차/MX · CE)에서 원본을 모아 실행
    python tools/run_all.py --week 43       그 주차로 (이번주_실행.cmd 가 --week now)
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time
from pathlib import Path

from common import REPORT_DIR, ROOT, config, current_slots
from check import run_check
from inputs import collect, downloads, put
from mapping import propose
from campaign_rules import propose as propose_campaign
from auto_mapping import run as auto_mapping
from master_media import fix_masters
from summary import inputs_markdown, write_summary
from prune_backups import run as prune_backups


def collect_week(week_arg: str, root: str | None = None) -> None:
    """그 주차 MX · CE 폴더 → 1_Input. 지난 입력은 _보관 으로 옮기고 이번 주 것만 둔다."""
    import shutil
    from datetime import date, datetime

    import collect as collector
    week = date.today().isocalendar()[1] if week_arg == "now" else int(week_arg)
    if root:
        m = re.search(r"(\d+)\s*주차", Path(root).name)
        week = int(m.group(1)) if m else week
    old = list(config.INPUT_DIR.glob("*.csv"))
    if old:
        dest = config.INPUT_DIR / "_보관" / f"{datetime.now():%Y%m%d_%H%M%S}_W{week}전"
        dest.mkdir(parents=True, exist_ok=True)
        for f in old:
            shutil.move(str(f), dest / f.name)
        print(f"■ 지난 입력 {len(old)}개 → {dest.relative_to(ROOT)}")
    for kind in ("MX", "CE"):
        collector.main([str(week), kind, "--put"] + (["--root", root] if root else []))
        print()


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("paths", nargs="*", type=Path)
    ap.add_argument("--downloads", action="store_true")
    ap.add_argument("--check-only", action="store_true")
    ap.add_argument("--open", action="store_true", help="끝나면 Reports 폴더 열기")
    ap.add_argument("--no-auto", action="store_true", help="매핑 자동 추론 등록 끄기")
    ap.add_argument("--week", help="now 또는 주차 번호 — 국가별 weekly 폴더에서 원본 수집 후 실행")
    ap.add_argument("--root", help="주차 폴더를 직접 지정(GUI 의 폴더 선택). 주차 번호는 폴더 이름에서")
    ap.add_argument("--remap", nargs="?", const="ALL", default="",
                    help="매핑 변경 반영 — 숫자가 같아도 입력 기간 전체를 다시 씀. --remap SEPR,SETK = 그 법인만")
    a = ap.parse_args(argv)

    t0 = time.time()
    if a.week or a.root:
        collect_week(a.week or "now", a.root)
    files = collect(a.paths) + (downloads() if a.downloads else [])
    if files:
        print("■ 넣기")
        put(files)
    if not run_check():
        return 1
    if a.check_only:
        return 0

    print()
    auto_mapping(apply=not a.no_auto)       # 확실한 매핑만 매핑 엑셀에 추가 — 이번 클렌징부터 반영
    infos = [i for v in current_slots().values() for i in v]
    print("\n■ 클렌징 (run.py)")
    since = time.time()
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONPATH=str(ROOT))
    remap = [] if not a.remap else ["--remap"] if a.remap == "ALL" else [f"--remap={a.remap}"]
    rc = subprocess.call([sys.executable, "-X", "utf8", str(ROOT / "run.py")] + remap, cwd=ROOT, env=env)
    if rc != 0:
        print(f"\n  ✖ run.py 실패(코드 {rc}) — {ROOT / 'cleansing_log.txt'} 확인")
        return rc

    print()
    fix_masters(apply=True)                 # 마스터에 쌓인 과거 행도 미디어 표준표로
    print()
    sheet = propose()
    camp_sheet = propose_campaign()
    write_summary(since, inputs_markdown(infos))
    print()
    prune_backups(apply=True)               # 백업은 종류마다 최신 1개만(요약이 직전 백업을 읽은 뒤)
    print(f"\n■ 끝 — {time.time() - t0:.0f}초.  결과: {config.OUTPUT_DIR}")
    print(f"  요약: {REPORT_DIR / '실행요약.md'}")
    if sheet:
        print(f"  다음: {sheet.name} 에서 '등록' 칸에 O → 매핑등록.cmd (등록 후 자동 재실행)")
    if camp_sheet:
        print(f"  캠페인: {camp_sheet.name} 에서 '사용' 칸에 O → 캠페인규칙등록.cmd")
    if a.open and os.name == "nt":
        os.startfile(REPORT_DIR)  # noqa: S606
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
