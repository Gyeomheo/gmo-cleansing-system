"""입력 넣기 · 현황 — 받은 CSV 를 내용으로 판별해 1_Input 의 제자리(슬롯)에 넣는다.

    python tools/inputs.py                    현황만 (1_Input 슬롯별 기간 · 경고)
    python tools/inputs.py 파일.csv 폴더 ...   해당 파일을 판별해서 넣기
    python tools/inputs.py --downloads        다운로드 폴더 최근 7일 CSV 중 GMO 형식만 넣기

- 슬롯 = (법인, MX/CE). 같은 슬롯 파일이 이미 있으면 그 이름을 그대로 쓰고
  기존 파일은 1_Input/_보관/<시각>/ 로 옮긴다 → 결과 파일명(cleaned_*~_<이름>)이 회차마다 같다.
- 새 슬롯은 <법인>.csv(MX) · <법인>_CE.csv(CE).
- 받은 파일은 복사만 한다(원본은 그대로).
"""
from __future__ import annotations

import argparse
import shutil
import time
from datetime import datetime
from pathlib import Path

from common import (ARCHIVE_DIR, InputInfo, config, current_slots, fmt_money,
                    inspect_csv, master_max_dates)

DOWNLOADS = Path.home() / "Downloads"


def print_table(infos: list[InputInfo], title: str) -> None:
    print(f"\n[{title}]")
    print(f"  {'파일':<28}{'법인':<7}{'구분':<5}{'행':>8}  {'기간':<23}{'Spend(USD)':>14}  경고")
    for i in infos:
        period = f"{i.date_min}~{i.date_max}" if i.date_min else "-"
        warn = " / ".join(i.warnings)
        print(f"  {i.path.name[:27]:<28}{i.sub:<7}{i.kind:<5}{i.rows:>8,}  {period:<23}"
              f"{fmt_money(i.spend):>14}  {warn}")


def status() -> list[str]:
    """1_Input 현황을 찍고 문제 목록을 돌려준다."""
    slots = current_slots()
    flat = [i for v in slots.values() for i in v]
    if not flat:
        print(f"\n[1_Input 현황] 비어 있음 — {config.INPUT_DIR}")
        return ["1_Input 이 비어 있음"]
    last = master_max_dates()
    for i in flat:
        if i.ok and i.slot in last and i.date_max <= last[i.slot]:
            i.warnings.append(f"새 날짜 없음(마스터 최신 {last[i.slot]}) — 수치 수정분만 반영")
    print_table(flat, "1_Input 현황")
    problems = []
    for slot, items in slots.items():
        if slot.startswith("?"):
            problems.append(f"{items[0].path.name}: GMO 형식 아님 ({'; '.join(items[0].warnings)})")
        elif len(items) > 1:
            names = ", ".join(i.path.name for i in items)
            problems.append(f"{slot} 슬롯 중복: {names} — 하나만 남길 것")
    return problems


def target_name(info: InputInfo, slots: dict[str, list[InputInfo]]) -> str:
    existing = slots.get(info.slot)
    if existing:
        return existing[0].path.name
    return f"{info.sub}.csv" if info.kind == "MX" else f"{info.sub}_CE.csv"


def collect(paths: list[Path]) -> list[Path]:
    files: list[Path] = []
    for p in paths:
        if p.is_dir():
            files += sorted(p.glob("*.csv"))
        elif p.suffix.lower() == ".csv":
            files.append(p)
    return [f for f in files if f.resolve().parent != config.INPUT_DIR.resolve()]


def put(files: list[Path]) -> int:
    infos = [inspect_csv(f) for f in files]
    gmo = [i for i in infos if i.ok]
    skipped = [i for i in infos if not i.ok]
    if skipped:
        print_table(skipped, "넣지 않음 — GMO 형식 아님")
    if not gmo:
        print("\n넣을 GMO CSV 가 없습니다.")
        return 0

    # 같은 슬롯이 여러 개 들어오면 마지막 날이 가장 늦은 것 하나만
    best: dict[str, InputInfo] = {}
    for i in gmo:
        if i.slot not in best or (i.date_max, i.path.stat().st_mtime) > (
                best[i.slot].date_max, best[i.slot].path.stat().st_mtime):
            best[i.slot] = i
    for i in gmo:
        if best[i.slot] is not i:
            i.warnings.append("같은 슬롯 중 더 최신 파일 있음 → 건너뜀")

    slots = current_slots()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    print_table(gmo, "넣을 파일")
    moved = 0
    for i in best.values():
        name = target_name(i, slots)
        dest = config.INPUT_DIR / name
        for old in slots.get(i.slot, []):
            arch = ARCHIVE_DIR / stamp
            arch.mkdir(parents=True, exist_ok=True)
            shutil.move(str(old.path), arch / old.path.name)
        shutil.copy2(i.path, dest)
        moved += 1
        print(f"  → {i.slot:<10} {i.path.name}  ⇒  1_Input/{name}")
    if any(slots.get(i.slot) for i in best.values()):
        print(f"  (기존 파일은 1_Input/_보관/{stamp}/ 로 옮김)")
    return moved


def downloads(days: int = 7) -> list[Path]:
    cutoff = time.time() - days * 86400
    return [p for p in DOWNLOADS.glob("*.csv") if p.stat().st_mtime >= cutoff]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("paths", nargs="*", type=Path)
    ap.add_argument("--downloads", action="store_true", help="다운로드 폴더 최근 7일 CSV")
    a = ap.parse_args(argv)
    files = collect(a.paths)
    if a.downloads:
        files += downloads()
    if files:
        put(files)
    problems = status()
    for p in problems:
        print(f"  ⚠ {p}")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
