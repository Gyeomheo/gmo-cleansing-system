"""실행 요약 — 이번 실행으로 마스터가 어떻게 바뀌었는지 법인별로 숫자로 남긴다.

    python tools/summary.py [--since <epoch>]

비교 기준 = 이번 실행이 만든 Backup_<MX|CE>_*.csv(실행 직전 마스터). 없으면 가장 최근 백업.
결과: 화면 + 2_Output/Reports/실행요약.md (매 실행 덮어씀, 지난 것은 Reports/_요약기록/)
"""
from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

from common import REPORT_DIR, config, fmt_money

METRICS = ["Media Spend (USD)", "Revenue"]
OUT = REPORT_DIR / "실행요약.md"


def _load(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, low_memory=False, dtype={"Subsidiary": str, "Date": str})
    for c in METRICS:
        df[c] = pd.to_numeric(df.get(c, 0), errors="coerce").fillna(0)
    df["Subsidiary"] = df["Subsidiary"].fillna("Unknown")
    return df


def _by_sub(df: pd.DataFrame) -> pd.DataFrame:
    # 제품 시리즈가 Other* · Multi 로 빠진 Spend — 매핑을 고치면 이 값이 줄어든다
    ser = df.get("Product Series", pd.Series("", index=df.index)).fillna("").astype(str).str.lower()
    df = df.assign(generic=df[METRICS[0]].where(ser.str.startswith(("other", "multi")), 0))
    return df.groupby("Subsidiary").agg(rows=("Date", "size"), last=("Date", "max"),
                                        spend=(METRICS[0], "sum"), revenue=(METRICS[1], "sum"),
                                        generic=("generic", "sum"))


def _stamp(p: Path) -> float:
    """Backup_MX_20261013_112013.csv → 만든 시각. (copy2 라 mtime 은 원본 시각이라 못 쓴다)"""
    try:
        return datetime.strptime(p.stem.split("_", 2)[2], "%Y%m%d_%H%M%S").timestamp()
    except (IndexError, ValueError):
        return 0.0


def backup_before(kind: str, since: float | None) -> Path | None:
    backups = sorted(config.BACKUP_DIR.glob(f"Backup_{kind}_*.csv"), key=_stamp)
    if since is not None:
        mine = [b for b in backups if _stamp(b) >= since - 1]
        return mine[0] if mine else None
    return backups[-1] if backups else None


def diff_table(kind: str, since: float | None) -> tuple[list[str], list[str]]:
    master = config.MASTER_FILES[kind]
    if not master.exists():
        return [f"(마스터 없음: {master.name})"], []
    after = _by_sub(_load(master))
    bk = backup_before(kind, since)
    before = _by_sub(_load(bk)) if bk else after.iloc[0:0]
    t = after.join(before, rsuffix="_b", how="outer").fillna({"rows": 0, "rows_b": 0, "spend": 0,
                                                              "spend_b": 0, "revenue": 0, "revenue_b": 0,
                                                              "generic": 0, "generic_b": 0,
                                                              "last": "", "last_b": ""})
    lines = ["| 법인 | 행 | 행 변화 | 마지막 날 | (전) | Spend 변화 | Revenue 변화 | Other·Multi Spend | (변화) |",
             "|---|---:|---:|---|---|---:|---:|---:|---:|"]
    flags = []
    for sub, r in t.iterrows():
        d_rows = int(r["rows"] - r["rows_b"])
        d_sp, d_rv = r["spend"] - r["spend_b"], r["revenue"] - r["revenue_b"]
        d_sp, d_rv = round(d_sp) or 0, round(d_rv) or 0      # -0 표시 방지
        d_gen = round(r["generic"] - r["generic_b"]) or 0
        changed = (d_rows or abs(d_sp) >= 0.01 or abs(d_rv) >= 0.01 or abs(d_gen) >= 0.01
                   or r["last"] != r["last_b"])
        mark = "**" if changed else ""
        lines.append(f"| {mark}{sub}{mark} | {int(r['rows']):,} | {d_rows:+,} | {r['last']} | {r['last_b']} |"
                     f" {d_sp:+,.0f} | {d_rv:+,.0f} | {r['generic']:,.0f} | {d_gen:+,.0f} |")
        if sub in ("Unknown", "", "nan"):
            flags.append(f"{kind} 마스터에 법인 없는 행 {int(r['rows']):,}개(Subsidiary 빈 값) — 원본 확인")
        if d_rows < 0:
            flags.append(f"{kind} {sub}: 행이 {d_rows:,} 줄었음 — 새 파일 기간이 짧은지 확인")
    head = f"기준 백업: {bk.name if bk else '없음'}"
    return [head, ""] + lines, flags


def unmapped_counts() -> tuple[int, int]:
    def n(p: Path) -> int:
        if not p.exists():
            return 0
        df = pd.read_csv(p, dtype=str).fillna("")
        col = df.columns[0]
        return int((~df[col].str.strip().isin(["", "-"])).sum())
    return n(REPORT_DIR / "Unmapped_Products.csv"), n(REPORT_DIR / "Unmapped_Media.csv")


def write_summary(since: float | None = None, inputs_md: list[str] | None = None) -> Path:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    md = [f"# GMO 클렌징 실행 요약 — {now:%Y-%m-%d %H:%M}", ""]
    if inputs_md:
        md += ["## 입력", ""] + inputs_md + [""]
    flags_all = []
    for kind in ("MX", "CE"):
        lines, flags = diff_table(kind, since)
        md += [f"## {kind} 마스터 ({config.MASTER_FILES[kind].name})", ""] + lines + [""]
        flags_all += flags
    up, um = unmapped_counts()
    md += ["## 확인할 것", ""]
    if up or um:
        flags_all.insert(0, f"미매핑 — 제품 {up}건 · 미디어 {um}건 → Reports/매핑등록.xlsx 에서 O 표시 후 매핑등록.cmd")
    md += [f"- {f}" for f in flags_all] or ["- 없음"]
    if OUT.exists():
        hist = REPORT_DIR / "_요약기록"
        hist.mkdir(exist_ok=True)
        shutil.copy2(OUT, hist / f"실행요약_{datetime.fromtimestamp(OUT.stat().st_mtime):%Y%m%d_%H%M%S}.md")
    OUT.write_text("\n".join(md) + "\n", encoding="utf-8")
    print("\n" + "\n".join(md))
    return OUT


def inputs_markdown(infos) -> list[str]:
    lines = ["| 파일 | 법인 | 구분 | 행 | 기간 | Spend(USD) | Revenue(USD) | 경고 |",
             "|---|---|---|---:|---|---:|---:|---|"]
    for i in infos:
        lines.append(f"| {i.path.name} | {i.sub} | {i.kind} | {i.rows:,} | {i.date_min}~{i.date_max} | "
                     f"{fmt_money(i.spend)} | {fmt_money(i.revenue)} | {' / '.join(i.warnings)} |")
    return lines


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", type=float, default=None)
    write_summary(ap.parse_args().since)
