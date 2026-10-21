"""매핑 점검 — 매핑이 '없는' 값이 아니라 '틀렸을' 수 있는 행을 경고로 남긴다(막지 않는다).

기준 = 취합본(진짜 마스터). 취합(consolidate) 때 '이전 취합본 RAW → 새 RAW' 로 자동 실행된다.
    python tools/map_check.py CE                       지금 최신 취합본만 (④ · ⑥)
    python tools/map_check.py CE --old <이전 취합본>   ⑤ 까지
    python tools/map_check.py CE --all-weeks           ⑥ 을 올해 모든 주차로(기준값 정할 때)

④ 출시 전 Spend   세대 번호가 있는 시리즈(S26 · Z Fold8 · Watch9 …)의 출시 주 = 전 법인 주간 Spend 가
                  처음으로 최고 주의 5% 를 넘은 주. 그보다 4주 넘게 앞선 Spend → 매핑이나 원본 날짜 확인
⑤ 분류 바뀜       이전 취합본과 법인 · 날짜별 숫자가 같은데 시리즈별 Spend 가 달라진 것 = 다시 분류된 것.
                  매핑을 고친 뒤 원치 않은 행까지 바뀌었는지 본다
⑥ Multi · Other   법인 · 주차별 Spend 중 Multi · Other* 비중이 앞 8주 중앙값보다 15%p 넘게 늘고 그 주 그 금액이
                  $20K 이상 → 아직 못 잡은 신제품일 수 있음. 그 행의 캠페인명 상위 3개를 같이 적는다
MX 는 Product Series, CE 는 Product Category 기준(CE 시리즈는 대부분 Multi · Others 라 못 쓴다).
결과: 2_Output/Reports/매핑점검_<MX|CE>.md
"""
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

from common import REPORT_DIR, config, fmt_money

SP = "Media Spend (USD)"
LEVEL = {"MX": "Product Series", "CE": "Product Category"}
LAUNCH_SHARE, LAUNCH_GRACE_DAYS = 0.05, 28
GENERIC_PP, GENERIC_MIN, BASE_WEEKS = 0.15, 20000, 8
MOVE_MIN = 100


def _prep(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ("Subsidiary", "Product Category", "Product Series", "Campaign Name"):
        df[c] = df[c].fillna("").astype(str).str.strip() if c in df.columns else ""
    df[SP] = pd.to_numeric(df[SP], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Week"] = pd.to_numeric(df["Week"], errors="coerce")
    return df


def is_generic(v: str) -> bool:
    return v.strip().lower().startswith(("other", "multi")) or v.strip() == ""


def _campaigns(rows: pd.DataFrame, n: int = 3) -> str:
    top = rows.groupby("Campaign Name")[SP].sum().sort_values(ascending=False).head(n)
    return " · ".join(f"{c[:40]} {fmt_money(v)}" for c, v in top.items() if c)


# ---------------------------------------------------------------- ④
def prelaunch(df: pd.DataFrame, kind: str) -> list[dict]:
    lv = LEVEL[kind]
    d = df[(df[SP] > 0) & df[lv].str.contains(r"\d") & ~df[lv].map(is_generic)]
    out = []
    for ser, g in d.groupby(lv):
        weekly = g.groupby(g["Date"].dt.to_period("W-SUN"))[SP].sum()
        launch = weekly[weekly >= weekly.max() * LAUNCH_SHARE].index.min().start_time
        early = g[g["Date"] < launch - pd.Timedelta(days=LAUNCH_GRACE_DAYS)]
        for sub, e in early.groupby("Subsidiary"):
            out.append({"구분": kind, "법인": sub, "분류": ser, "출시 주(추정)": f"{launch:%Y-%m-%d}",
                        "기간": f"{e['Date'].min():%m-%d}~{e['Date'].max():%m-%d}", "행": len(e),
                        "Spend": e[SP].sum(), "캠페인": _campaigns(e)})
    return sorted(out, key=lambda r: -r["Spend"])


# ---------------------------------------------------------------- ⑤
PRD = {"MX": "Products", "CE": "Products (Optional)"}
ROW_KEY = ["Subsidiary", "Date", "Campaign Name", SP, "Impressions", "Clicks", "Revenue"]


def _cls(df: pd.DataFrame, kind: str) -> pd.Series:
    """Category / Series / Products 한 줄."""
    p = df[PRD[kind]].fillna("").astype(str).str.strip() if PRD[kind] in df.columns else ""
    return df["Product Category"] + " / " + df["Product Series"] + " / " + p


def _paired(df: pd.DataFrame, old: pd.DataFrame, kind: str) -> pd.DataFrame:
    """이전 · 새 취합본에서 숫자까지 같은 행을 짝지어 분류가 바뀐 행만(대소문자만 다른 건 뺌)."""
    def keyed(x):
        k = x[ROW_KEY].copy()
        for c in ROW_KEY[3:]:
            k[c] = pd.to_numeric(k[c], errors="coerce").fillna(0).round(2)
        k["Date"] = k["Date"].dt.strftime("%Y-%m-%d")
        k["_n"] = k.groupby(ROW_KEY).cumcount()                 # 완전히 같은 행이 여럿이면 순서대로 짝
        return k.assign(cls=_cls(x, kind))
    m = keyed(df).merge(keyed(old), on=ROW_KEY + ["_n"], suffixes=("", "_전"))
    return m[m["cls"].str.lower() != m["cls_전"].str.lower()]


def _input_values(kind: str) -> dict[tuple, str]:
    """이번 실행 cleaned_*.csv 의 원본 입력값(클렌징 전 Category / Series / Products) — 키: 법인 · 날짜 · 캠페인 · Spend."""
    out = {}
    for f in sorted(config.OUTPUT_DIR.glob("cleaned_*.csv")):
        if "_Summary" in f.name:
            continue
        try:
            c = pd.read_csv(f, dtype=str, low_memory=False).fillna("")
        except Exception:                                       # noqa: BLE001 — 읽기 실패한 파일은 건너뜀
            continue
        if "Product Category_cleaned" not in c.columns:
            continue
        prd = next((x for x in ("Products", "Products (Optional)") if x in c.columns), None)
        raw = (c.get("Product Category", "") + " / " + c.get("Product Series", "")
               + " / " + (c[prd] if prd else ""))
        sp = pd.to_numeric(c[SP], errors="coerce").fillna(0).round(2)
        d = pd.to_datetime(c["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        for k, v in zip(zip(c["Subsidiary"].str.strip(), d, c["Campaign Name"].str.strip(), sp), raw):
            out.setdefault(k, v)
    return out


def reclassified(df: pd.DataFrame, old: pd.DataFrame, kind: str) -> tuple[list[dict], pd.DataFrame]:
    """(묶음 요약, 행 단위 상세). 묶음 = 법인 · 전 분류 · 후 분류."""
    m = _paired(df, old, kind)
    if m.empty:
        return [], m
    look = _input_values(kind)
    m["원본 입력값"] = [look.get(k, "") for k in zip(m["Subsidiary"], m["Date"], m["Campaign Name"], m[SP])]
    out = []
    for (sub, a, b), g in m.groupby(["Subsidiary", "cls_전", "cls"]):
        if g[SP].sum() < MOVE_MIN and len(g) < 10:
            continue
        inputs = g["원본 입력값"][g["원본 입력값"] != ""].value_counts().head(3)
        out.append({"구분": kind, "법인": sub, "전": a, "후": b, "행": len(g), "Spend": g[SP].sum(),
                    "기간": f"{g['Date'].min()[5:]}~{g['Date'].max()[5:]}",
                    "원본 입력값": " · ".join(f"{v} ({n})" for v, n in inputs.items()),
                    "캠페인": " · ".join(f"{c[:40]}" for c in g.groupby("Campaign Name")[SP].sum()
                                       .sort_values(ascending=False).head(3).index if c)})
    detail = m.rename(columns={"cls_전": "전 분류", "cls": "후 분류"})[
        ["Subsidiary", "Date", "Campaign Name", SP, "전 분류", "후 분류", "원본 입력값"]]
    return sorted(out, key=lambda r: -r["Spend"]), detail


# ---------------------------------------------------------------- ⑥
def generic_jump(df: pd.DataFrame, kind: str, all_weeks: bool) -> list[dict]:
    lv = LEVEL[kind]
    d = df[df["Week"].notna()].assign(gen=lambda x: x[SP].where(x[lv].map(is_generic), 0))
    out = []
    for sub, g in d.groupby("Subsidiary"):
        w = g.groupby("Week")[[SP, "gen"]].sum()
        w = w[w[SP] > 0]
        share = w["gen"] / w[SP]
        weeks = w.index[BASE_WEEKS:] if all_weeks else w.index[-2:]
        for wk in weeks:
            base = share[share.index < wk].tail(BASE_WEEKS)
            if len(base) < 3:
                continue
            up = share[wk] - base.median()
            if up >= GENERIC_PP and w.loc[wk, "gen"] >= GENERIC_MIN:
                rows = g[(g["Week"] == wk) & g[lv].map(is_generic)]
                out.append({"구분": kind, "법인": sub, "주차": int(wk), "비중": share[wk], "평소": base.median(),
                            "Multi·Other Spend": w.loc[wk, "gen"], "캠페인": _campaigns(rows)})
    return out


def run(kind: str, new: pd.DataFrame, old: pd.DataFrame | None = None, all_weeks: bool = False,
        label: str = "", out_dir: Path = REPORT_DIR) -> tuple[Path, int]:
    """new = 새 취합본 RAW, old = 이전 취합본 RAW(없으면 ⑤ 생략). (보고서, 경고 수)."""
    df = _prep(new)
    pre = prelaunch(df, kind)
    moves, detail = reclassified(df, _prep(old), kind) if old is not None else ([], None)
    jumps = generic_jump(df, kind, all_weeks)
    md = [f"# 매핑 점검 — {kind} · {datetime.now():%Y-%m-%d %H:%M}", "",
          f"- 대상: {label or '취합본'}", "- 경고만 남긴다(막지 않음). 맞으면 넘기고, 틀리면 매핑등록 · 원본 확인.", "",
          f"## ④ 출시 전 Spend — {len(pre)}건", "",
          f"출시 주 = 전 법인 주간 Spend 가 처음 최고 주의 {LAUNCH_SHARE:.0%} 를 넘은 주, "
          f"그보다 {LAUNCH_GRACE_DAYS // 7}주 넘게 앞선 행", ""]
    if pre:
        md += ["| 법인 | 분류 | 출시 주(추정) | 기간 | 행 | Spend | 캠페인 상위 |", "|---|---|---|---|---:|---:|---|"]
        md += [f"| {r['법인']} | {r['분류']} | {r['출시 주(추정)']} | {r['기간']} | {r['행']:,} |"
               f" {fmt_money(r['Spend'])} | {r['캠페인']} |" for r in pre]
    md += ["", f"## ⑤ 분류 바뀜 — {len(moves)}건", "",
           ("이전 취합본과 숫자까지 같은 행을 짝지어 Category / Series / Products 가 바뀐 것(대소문자만 다른 건 제외). "
            f"행 단위는 매핑점검_{kind}_분류바뀜.csv, 원본 입력값 = 이번 실행 cleaned 파일의 클렌징 전 값")
           if old is not None else "이전 취합본 없음 — 생략", ""]
    if moves:
        md += ["| 법인 | 전 | 후 | 원본 입력값 | 행 | Spend | 기간 | 캠페인 상위 |",
               "|---|---|---|---|---:|---:|---|---|"]
        md += [f"| {r['법인']} | {r['전']} | {r['후']} | {r['원본 입력값']} | {r['행']:,} | {fmt_money(r['Spend'])} |"
               f" {r['기간']} | {r['캠페인']} |" for r in moves]
    md += ["", f"## ⑥ Multi · Other 비중 급증 — {len(jumps)}건", "",
           f"앞 {BASE_WEEKS}주 중앙값보다 {GENERIC_PP:.0%}p 넘게 늘고 그 주 ${GENERIC_MIN:,} 이상"
           + (" (올해 모든 주차)" if all_weeks else " (법인별 최근 2주)"), ""]
    if jumps:
        md += ["| 법인 | 주차 | 비중 | 평소 | Multi·Other Spend | 캠페인 상위 |", "|---|---:|---:|---:|---:|---|"]
        md += [f"| {r['법인']} | W{r['주차']} | {r['비중']:.0%} | {r['평소']:.0%} |"
               f" {fmt_money(r['Multi·Other Spend'])} | {r['캠페인']} |" for r in jumps]
    out_dir.mkdir(parents=True, exist_ok=True)
    if detail is not None and len(detail):
        detail.to_csv(out_dir / f"매핑점검_{kind}_분류바뀜.csv", index=False, encoding="utf-8-sig")
    out = out_dir / f"매핑점검_{kind}.md"
    out.write_text("\n".join(md) + "\n", encoding="utf-8")
    n = len(pre) + len(moves) + len(jumps)
    print(f"  매핑 점검 — ④ 출시 전 {len(pre)} · ⑤ 분류 바뀜 {len(moves)} · ⑥ Multi·Other 급증 {len(jumps)}  → {out.name}")
    return out, n


if __name__ == "__main__":
    from consolidate import latest_file, read_raw

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["MX", "CE", "mx", "ce"])
    ap.add_argument("--file", type=Path, help="점검할 취합본(기본: 최신)")
    ap.add_argument("--old", type=Path, help="비교할 이전 취합본(⑤)")
    ap.add_argument("--all-weeks", action="store_true")
    ap.add_argument("--out", type=Path, default=REPORT_DIR)
    a = ap.parse_args()
    k = a.kind.upper()
    f = a.file or latest_file(k)
    run(k, read_raw(f)[0], read_raw(a.old)[0] if a.old else None, a.all_weeks, f.name, a.out)
