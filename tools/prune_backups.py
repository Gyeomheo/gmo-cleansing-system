"""백업 정리 — 종류마다 최신 1개만 남긴다(담당자 2026-10-27).

    python tools/prune_backups.py            미리보기(지울 것만 표)
    python tools/prune_backups.py --apply    지우기 (run_all 이 끝날 때 자동)

- 2_Output/Backup: 이름에서 시각(YYYYMMDD_HHMMSS)을 뺀 모양이 같으면 같은 종류
  (Backup_MX_#.csv · Backup_MX_#_media.csv · Backup_CE_#.csv · 백과사전_#.xlsx · 덮어쓰기전_#_<취합본 이름> …)
  · Backup_<구분>_#.csv 는 mark_done 의 비교 기준(그 회차 실행 직전 마스터)이라 _media 와 섞지 않는다
  · 시각이 없는 파일(손으로 만든 백업)은 건드리지 않는다
- 1_Input/_보관: 날짜로 시작하는 폴더 중 최신 1개만
"""
from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path

from common import ARCHIVE_DIR, config

STAMP = re.compile(r"\d{8}_\d{6}")


def plan() -> list[Path]:
    groups: dict[str, list[Path]] = {}
    for p in config.BACKUP_DIR.iterdir() if config.BACKUP_DIR.exists() else []:
        m = STAMP.search(p.name)
        if p.is_file() and m:
            groups.setdefault(STAMP.sub("#", p.name, count=1), []).append(p)
    drop = [p for ps in groups.values() for p in sorted(ps, key=lambda x: STAMP.search(x.name).group())[:-1]]
    if ARCHIVE_DIR.exists():
        dirs = sorted(d for d in ARCHIVE_DIR.iterdir() if d.is_dir() and re.match(r"\d{8}", d.name))
        drop += dirs[:-1]
    return drop


def size(p: Path) -> int:
    return sum(f.stat().st_size for f in p.rglob("*") if f.is_file()) if p.is_dir() else p.stat().st_size


def run(apply: bool) -> int:
    drop = plan()
    total = sum(size(p) for p in drop)
    print(f"■ 백업 정리 — 종류마다 최신 1개만: {len(drop)}개 · {total / 1e6:,.0f}MB")
    if not apply:
        for p in drop:
            print(f"  - {p.relative_to(config.BASE_DIR)}")
        print("  미리보기만 — 지우려면 --apply")
        return 0
    for p in drop:
        shutil.rmtree(p) if p.is_dir() else p.unlink()
    print("  ✔ 지움")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--apply", action="store_true")
    return run(ap.parse_args(argv).apply)


if __name__ == "__main__":
    raise SystemExit(main())
