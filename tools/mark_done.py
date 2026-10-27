"""처리 표시 — 클렌징을 마친 주차의 국가별 weekly 폴더에 꼬리표를 붙이고, 클렌징본을 만든다(담당자 방식 그대로).

    python tools/mark_done.py 43 CE            미리보기(무엇을 어떻게 바꿀지 표만)
    python tools/mark_done.py 43 CE --apply    실제로 만들고 이름 바꾸기
    python tools/mark_done.py 43 CE --base 2_Output/Backup/Backup_CE_20261020_135138.csv
    python tools/mark_done.py 43 CE --root <시험 폴더>/43주차 --apply     (공유 드라이브 대신 복사본에서 시험)

- 원본 유지 / 클렌징본: 이번에 값을 쓴 파일은 **그대로 두고**, 옆에 '<이름> - 허겸수정(MMDD~)_a.xlsx' 복사본을 만들어
  GMO 시트에서 이번에 쓴 행의 값을 클렌징 결과(2_Output/cleaned_*)로 바꾼다. zip 안 파일이면 풀어 둔 폴더 안에.
  · 값을 쓴 zip 은 같은 폴더에 '<zip 이름>/' 으로 풀어 둔다(이미 있으면 그대로). 다음 수집은 풀어 둔 파일을 읽는다(같은 내용)
  · 바꾸는 칸: 머리줄 이름이 클렌징 결과 열과 같은 칸(날짜 열 · 수식이 든 열은 그대로). 클렌징 값이 빈 칸이면 원본 유지
  · 클렌징 결과와 행이 하나씩 맞는지(날짜 · Spend) 먼저 확인하고, 안 맞으면 그 법인은 만들지 않는다
  MMDD = 그 파일에서 쓴 행 중 마스터가 바뀐 첫 날(updater 함수로 클렌징 결과 vs 실행 직전 마스터 — 숫자가 처음 달라진 날 · 새 날짜)
- 값을 쓴 파일이 있는 법인 폴더 → '<법인 MMDD>_A'
- 안 쓴 파일(최신 파일에 같은 날짜가 있어서 등) → '<이름> XXX', 안 쓴 것만 있는 폴더 → '<법인 MMDD> XXX'
- '복사본' 파일 · 구분이 다른 파일은 그대로. 이미 표시된 것은 건너뛴다. collect 는 허겸수정본을 읽지 않는다(원본을 읽음)
- 비교 기준 마스터(--base): 수집 CSV(2_Output/_수집/W<nn>_<구분>) 를 만든 뒤 처음 생긴 백업 = 그 회차 실행 직전 마스터
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from common import config
from collect import (HEADER_ALIASES, MARK_DONE, MARK_SKIP, STAGE, list_sources, merge_sub, read_source,
                     sanitize_column_headers)

METRICS = ["Media Spend (USD)", "Revenue", "Impressions", "Clicks", "Orders"]
# 원본 머리줄 → 클렌징 결과 열 (collect · run.py 의 이름 바꾸기와 같은 것)
TO_CLEAN = {**HEADER_ALIASES, "Product Category2": "Product Category", "Products (optional)": "Products"}
# 바꾸는 칸 = 분류 칸만(숫자 · 날짜 · Week 는 원본 그대로). 최종값은 '<열>_cleaned' 가 있으면 그것(format_ce/mx_data 와 같음)
WRITE_COLS = ["Sales Channel", "Partner", "Media Type 1", "Media Type 2", "Media Platform", "Funding", "BU",
              "Product Category", "Product Series", "Products", "Campaign Name", "Mindset", "Quarter", "Month"]


def _base_backup(stage: Path, kind: str) -> Path | None:
    """수집 CSV 이후 처음 만들어진 백업(파일명 시각 기준 — copy2 라 mtime 은 원본 시각)."""
    made = min((p.stat().st_mtime for p in stage.glob("*.csv")), default=None)
    if made is None:
        return None
    for p in sorted(config.BACKUP_DIR.glob(f"Backup_{kind}_*.csv")):
        m = re.fullmatch(rf"Backup_{kind}_(\d{{8}}_\d{{6}})\.csv", p.name)
        if m and datetime.strptime(m.group(1), "%Y%m%d_%H%M%S").timestamp() >= made - 60:
            return p
    return None


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", "").str.strip(), errors="coerce").fillna(0)


def _daily(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={"Revenue (USD)": "Revenue"})
    for c in METRICS:
        df[c] = _num(df[c]) if c in df.columns else 0.0
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    return df.groupby("Date")[METRICS].sum().round(2)


def _first_change(new: pd.DataFrame, old: pd.DataFrame) -> str | None:
    """새 데이터에서 마스터와 숫자가 처음 다른 날(마스터에 없는 날 포함). 같으면 None."""
    n, o = _daily(new), _daily(old) if len(old) else None
    for d in n.index:
        if d < config.MIN_DATE:
            continue
        if o is None or d not in o.index or ((n.loc[d] - o.loc[d]).abs() >= 0.01).any():
            return d
    return None


def _updater_cutoff(cl: pd.DataFrame, old: pd.DataFrame) -> str | None:
    """updater.update_smart_db 와 똑같이 — 클렌징 결과 vs 실행 직전 마스터, 교체가 시작된 날. 교체 없으면 None."""
    import logging
    from src.updater import force_numeric_metrics, get_safe_divergence_date, safe_parse_dates
    new = cl.copy()
    new["Date"] = safe_parse_dates(new["Date"])
    if old.empty:                                       # 마스터에 없던 법인 · 비교 기준 없음 = 전부 새 값
        return new["Date"].min().strftime("%Y-%m-%d")
    old = old.copy()
    old["Date"] = safe_parse_dates(old["Date"])
    for df in (new, old):
        df["Week"] = df["Week"].astype(str)
    new, old = force_numeric_metrics(new), force_numeric_metrics(old)
    level = logging.getLogger().level
    logging.getLogger().setLevel(logging.WARNING)       # updater 의 상세 로그는 끔
    try:
        c = get_safe_divergence_date(old.sort_values("Date"), new.sort_values("Date"))
    finally:
        logging.getLogger().setLevel(level)
    return c.strftime("%Y-%m-%d") if c <= new["Date"].max() else None


def _mmdd(d: str) -> str:
    return d[5:7] + d[8:10]


def _cleaned(sub: str, kind: str, merged: pd.DataFrame) -> pd.DataFrame | None:
    """그 법인의 최신 클렌징 결과 — 합친 표(MIN_DATE 이후)와 행이 하나씩 맞을 때만."""
    tail = f"_{sub}_ce.csv" if kind == "CE" else f"_{sub}.csv"
    cands = [p for p in config.OUTPUT_DIR.glob("cleaned_*.csv") if p.name.lower().endswith(tail.lower())]
    if not cands:
        return None
    cl = pd.read_csv(max(cands, key=os.path.getmtime), low_memory=False)
    m = merged[merged["Date"] >= config.MIN_DATE]
    if len(m) != len(cl):
        return None
    same_date = (m["Date"].values == pd.to_datetime(cl["Date"]).dt.strftime("%Y-%m-%d").values).all()
    same_spend = ((_num(m["Media Spend (USD)"]).values - _num(cl["Media Spend (USD)"]).values) ** 2 < 1e-4).all()
    if not (same_date and same_spend):
        return None
    return cl


def _xl_date(v) -> str:
    try:
        if isinstance(v, (int, float)):
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
        d = pd.to_datetime(str(v), errors="coerce")
        return d.strftime("%Y-%m-%d") if pd.notna(d) else ""
    except (OverflowError, ValueError):
        return ""


def write_cleaned(xl, path: Path, sheet: str, rows: dict[int, pd.Series]) -> str:
    """path(복사본)의 sheet 에서 rows{머리줄 아래 k번째 줄: 클렌징 행} 의 분류 칸을 클렌징 값으로. 결과 한 줄."""
    wb = xl.Workbooks.Open(str(path))
    try:
        ws = wb.Worksheets(sheet)
        ur = ws.UsedRange
        r0, c0 = ur.Row, ur.Column
        vals = ur.Value2
        h = next(i for i in range(min(20, len(vals)))
                 if any("Subsidiary" in str(v) for v in vals[i]) and any("Media Platform" in str(v) for v in vals[i]))
        heads = [sanitize_column_headers(pd.DataFrame(columns=[str(v)])).columns[0] if v is not None else ""
                 for v in vals[h]]
        want = {k: pd.to_datetime(s["Date"]).strftime("%Y-%m-%d") for k, s in rows.items()}
        # 날짜 칸: 'Date' 머리줄, 없으면(SIEL 'Column2') 날짜가 가장 많이 맞는 칸
        hits = {j: sum(_xl_date(vals[h + 1 + k][j]) == d for k, d in want.items() if h + 1 + k < len(vals))
                for j in range(len(heads)) if heads[j] == "Date" or "Date" not in heads}
        di = max(hits, key=hits.get)
        if hits[di] != len(want):
            raise RuntimeError(f"행 위치가 안 맞음({len(want) - hits[di]:,}행) — 만들지 않음")
        changed, cols, seen = 0, [], set()
        last = h + 1 + max(rows)
        sample = next(iter(rows.values())).index
        for j, name in enumerate(heads):
            col = TO_CLEAN.get(name, name)
            if col not in WRITE_COLS or col in seen:
                continue
            src = f"{col}_cleaned" if f"{col}_cleaned" in sample else col
            if src not in sample:
                continue
            seen.add(col)
            rng = ws.Range(ws.Cells(r0 + h + 1, c0 + j), ws.Cells(r0 + last, c0 + j))
            if any(str(f[0]).startswith("=") for f in rng.Formula):
                continue                                    # 수식 칸은 그대로
            col_vals = [[vals[i][j]] for i in range(h + 1, last + 1)]
            n = 0
            for k, s in rows.items():
                v = s[src]
                if pd.isna(v) or str(v).strip() == "":
                    continue                                # 클렌징 값이 비면 원본 유지
                v = v.item() if hasattr(v, "item") else v
                old = col_vals[k][0]
                if str(old if old is not None else "").strip() != str(v).strip():
                    col_vals[k][0] = v
                    n += 1
            if n:
                rng.Value2 = col_vals
                changed += n
                cols.append(f"{name} {n:,}")
        wb.Save()
        return f"{len(rows):,}행 중 칸 {changed:,}개 바꿈" + (f" — {', '.join(cols)}" if cols else "")
    finally:
        wb.Close(False)


def run(week: int, kind: str, base: Path | None, apply: bool, root: Path | None = None) -> int:
    stage = STAGE / f"W{week}_{kind}"
    base = base or _base_backup(stage, kind)
    master = pd.read_csv(base, low_memory=False) if base else pd.DataFrame(columns=["Subsidiary"])
    print(f"■ 처리 표시 — {week}주차 {kind}\n  비교 기준 마스터: {base or '없음(모든 날짜를 새 값으로 봄)'}")
    with tempfile.TemporaryDirectory() as td:
        sources = list_sources(week, kind, Path(td), root)
        if not sources:
            return 0
        kind_dir = next(p for p in sources[0].origin.parents if p.name == kind)
        by_sub: dict[str, list] = {}
        unread: set[Path] = set()            # GMO 데이터를 못 읽은 파일 — 안 쓴 게 아니라 모르는 것이므로 XXX 안 함
        for src in sources:
            df, src.sheet = read_source(src)
            if not df.empty:
                by_sub.setdefault(src.sub, []).append((src, df))
            else:
                unread.add(src.origin)
                print(f"  ⚠ {src.folder}/{src.label}: GMO 데이터를 못 읽음 — XXX 안 함, 원본 확인")
        cutoff, cleaned = {}, {}
        for sub, items in by_sub.items():
            merged, _ = merge_sub(items, kind)
            if merged.empty:
                continue
            old = master[master["Subsidiary"].astype(str).str.strip() == sub]
            cleaned[sub] = _cleaned(sub, kind, merged)
            cutoff[sub] = (_updater_cutoff(cleaned[sub], old) if cleaned[sub] is not None
                           else _first_change(merged, old))
            if cleaned[sub] is None:
                print(f"  ⚠ {sub}: 클렌징 결과(2_Output/cleaned_*)와 행이 안 맞음 — 클렌징본은 만들지 않고 표시만")

        # 클렌징본 · XXX — 실제 파일(zip 은 zip) 단위
        makes, renames, used_top, unzips, held_top = [], [], {}, [], set()
        by_origin: dict[Path, list] = {}
        for src in sources:
            by_origin.setdefault(src.origin, []).append(src)
        for origin, srcs in sorted(by_origin.items()):
            top = next(p for p in [origin, *origin.parents] if p.parent == kind_dir)
            used = [s for s in srcs if s.taken]
            used_top[top] = used_top.get(top, False) or bool(used)
            if origin in unread and not used:
                held_top.add(top)                       # 못 읽은 파일이 든 폴더도 XXX 하지 않는다
                continue
            if not used:
                if not MARK_SKIP.search(origin.name):
                    renames.append((origin, origin.with_name(f"{origin.stem} XXX{origin.suffix}"), "안 씀"))
                continue
            is_zip = origin.suffix.lower() == ".zip"
            if is_zip and not origin.with_suffix("").exists():
                unzips.append(origin)                   # 쓴 zip 은 같은 폴더에 '<zip 이름>/' 으로 풀어 둔다
            for s in used:
                c = cutoff.get(s.sub)
                start = max(s.taken[0], c) if c and c <= s.taken[1] else s.taken[0]
                src_file = s.path                       # zip 안이면 꺼낸 임시 파일
                name = f"{src_file.stem} - 허겸수정({_mmdd(start)}~)_a{src_file.suffix}"
                # zip 이면 풀어 둔 폴더 안 원본 옆(자리는 풀고 나서 확정), 아니면 원본 옆
                out = (origin.with_suffix("") if is_zip else origin.parent) / name
                makes.append((s, src_file, out))
        for top, used in sorted(used_top.items()):
            if used and not re.search(r"_[aA]$", top.name):
                renames.append((top, top.with_name(top.name + "_A"), "폴더 — 씀"))
            elif not used and top not in held_top and not MARK_SKIP.search(top.name):
                renames.append((top, top.with_name(top.name + " XXX"), "폴더 — 안 씀"))
        unused_tops = {t for t, u in used_top.items() if not u}
        renames = [r for r in renames if not any(t in r[0].parents for t in unused_tops)]

        for z in unzips:
            print(f"  [압축 풀기] {z.relative_to(kind_dir)}\n        → {z.stem}/")
        for s, _, out in makes:
            print(f"  [클렌징본] {s.folder}/{s.label}\n        → {out.name}  (원본 유지)")
        for old, new, why in renames:
            print(f"  [{why}] {old.relative_to(kind_dir)}\n        → {new.name}")
        if not makes and not renames and not unzips:
            print("  바꿀 것 없음(이미 표시됨)")
            return 0
        if not apply:
            print("\n  미리보기만 — 만들고 바꾸려면 --apply")
            return 0

        # 풀기 → 만들기(폴더 이름 바꾸기 전에) → 이름 바꾸기(파일 먼저, 폴더 나중)
        for z in unzips:
            with zipfile.ZipFile(z) as zf:
                zf.extractall(z.with_suffix(""))
            print(f"  ✔ 압축 풂: {z.stem}/")
        placed = []
        for s, f, out in makes:
            if s.origin.suffix.lower() == ".zip":        # 풀어 둔 폴더에서 원본 자리 찾기
                hit = next(s.origin.with_suffix("").rglob(f.name), None)
                out = (hit.parent if hit else out.parent) / out.name
            placed.append((s, f, out))
        makes = placed
        todo = [(s, f, o) for s, f, o in makes if cleaned.get(s.sub) is not None and not o.exists()]
        made = 0
        if todo:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.DisplayAlerts = False
            xl.ScreenUpdating = False
            try:
                for s, f, out in todo:
                    items = by_sub[s.sub]
                    merged_parts = sorted((x for x, _ in items if x.taken), key=lambda x: x.part_no)
                    offset = 0                          # 합친 표에서 이 파일 조각이 시작하는 자리
                    for x in merged_parts:
                        if x is s:
                            break
                        offset += len(x.taken_idx)
                    cl = cleaned[s.sub]
                    pos_all = [i for i, d in enumerate(_merged_dates(items)) if d >= config.MIN_DATE]
                    cl_pos = {p: n for n, p in enumerate(pos_all)}   # 합친 표 자리 → 클렌징 행
                    rows = {}
                    for n, k in enumerate(s.taken_idx):
                        p = offset + n
                        if p in cl_pos:
                            rows[k] = cl.iloc[cl_pos[p]]
                    shutil.copy2(f, out)
                    try:
                        msg = write_cleaned(xl, out, s.sheet, rows) if rows else "2026 행 없음 — 원본 그대로 복사"
                        print(f"  ✔ {out.name}: {msg}")
                        made += 1
                    except Exception as e:
                        out.unlink(missing_ok=True)
                        print(f"  ⚠ {out.name}: {e!r}")
            finally:
                xl.Quit()
        for old, new, _ in renames:
            old.rename(new)
        print(f"\n  ✔ 압축 풂 {len(unzips)}개 · 클렌징본 {made}개 · 이름 바꿈 {len(renames)}개")
    return 0


def _merged_dates(items) -> list[str]:
    """merge_sub 가 합친 순서 그대로의 날짜 목록."""
    out = []
    for x, d in sorted((p for p in items if p[0].taken), key=lambda p: p[0].part_no):
        out += d.loc[x.taken_idx, "Date"].tolist()
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("week", type=int)
    ap.add_argument("kind", choices=["CE", "MX", "ce", "mx"])
    ap.add_argument("--apply", action="store_true", help="실제로 만들고 이름 바꾸기")
    ap.add_argument("--base", type=Path, help="비교 기준 마스터(백업 CSV)")
    ap.add_argument("--root", type=Path, help="주차 폴더를 직접 지정(시험용 — 그 안의 MX · CE 폴더)")
    a = ap.parse_args(argv)
    return run(a.week, a.kind.upper(), a.base, a.apply, a.root)


if __name__ == "__main__":
    raise SystemExit(main())
