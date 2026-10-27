"""미디어 플랜 점검 — 플랜에 월별 계획이 있는 매체(Media Type 1)에 Weekly Raw 실적이 있는지.

    python tools/media_plan_check.py            → 2_Output/Reports/미디어플랜점검.xlsx · .md
    python tools/media_plan_check.py --plan-dir <폴더> --mx <MX 취합본> --ce <CE 취합본>

- 플랜: 'Media plan Check list/<1Q|2Q|3Q>/<법인>/' 의 '*Media_Plan*.xlsx'(old 폴더 제외), 시트 '1. Media Plan'.
  머리줄(Sub · Month 가 있는 줄)에서 열을 이름으로 찾는다. 계획 = 'Net Media Cost (Total)' > 0 인 행
  · MX/CE: Product Category 가 MX 분류(smartphones · tablets · pc · wearables · TWS · mx_other …)면 MX
  · Media Channel → Media Type 1 (search → Paid Search, pmax · shopping → Shopping Ads, social · dgen → Paid Social …)
  · 법인 오타는 고침(SEIL → SIEL)
- 실적: 최신 MX · CE 취합본 RAW(로컬 마스터는 빠진 달이 있어 쓰지 않는다). 그 법인 · 월 · Media Type 1 에 행이 있으면 실적 있음
- 판정은 Media Type 1 · 분기 단위(담당자 2026-10-27): 그 분기에 계획이 있는데 분기 실적 Spend 가 0 이면 '실적 없음'.
  월별은 '참고 월별' 시트. '그 기간 실적 매체' 칸 — 다른 매체로 분류해 낸 경우(SEDA RTB · jampp → App 등) 확인용
"""
from __future__ import annotations

import argparse
import os
import re
from pathlib import Path

import fastexcel
import openpyxl
import pandas as pd

from common import REPORT_DIR
from consolidate import DASH_DIR, NAME_RE

PLAN_DIR = Path(r"G:\공유 드라이브\Samsung Perform_1\문서-1\2026\02. GMO Consulting\05. Media plan Check list")
CHANNEL = {"search": "Paid Search", "pmax": "Shopping Ads", "shopping": "Shopping Ads", "social": "Paid Social",
           "dgen": "Paid Social", "display": "Display", "affiliate": "Affiliate", "app": "App", "olv": "Video",
           "epp": "Others", "other": "Others"}
MX_CATS = {"smartphones", "tablets", "pc", "wearables", "tws", "hearables", "mx_other", "mobile", "xr"}
SUB_FIX = {"SEIL": "SIEL"}


def latest_raw(kind: str) -> Path | None:
    best = None
    for f in DASH_DIR.glob("*.xlsx"):
        m = NAME_RE.match(f.name)
        if m and m.group(4) == kind and (best is None or (int(m.group(2)), int(m.group(3))) > best[0]):
            best = ((int(m.group(2)), int(m.group(3))), f)
    return best[1] if best else None


def read_plans(plan_dir: Path) -> pd.DataFrame:
    rows = []
    for f in sorted(plan_dir.glob("[1-4]Q/*/*.xlsx")):
        rel = os.path.relpath(f, plan_dir)
        if "Media_Plan" not in f.name or "old" in [p.lower() for p in Path(rel).parts]:
            continue
        ws = openpyxl.load_workbook(f, read_only=True, data_only=True)["1. Media Plan"]
        col = None
        for r in ws.iter_rows(values_only=True):
            if col is None:
                h = [re.sub(r"\s+", " ", str(x or "")).strip() for x in r]
                if "Sub" in h and "Month" in h:
                    find = lambda k: next(j for j, x in enumerate(h) if x.startswith(k))  # noqa: E731
                    col = {k: find(v) for k, v in (("sub", "Sub"), ("month", "Month"), ("ch", "Media Channel"),
                                                   ("pf", "Platform"), ("cat", "Product Category"),
                                                   ("cost", "Net Media Cost (Total)"))}
                continue
            month = str(r[col["month"]] or "").strip()
            cost = pd.to_numeric(r[col["cost"]], errors="coerce")
            if not r[col["sub"]] or not re.fullmatch(r"M\d{2}", month) or pd.isna(cost) or cost <= 0:
                continue
            sub = str(r[col["sub"]]).strip().upper()
            ch = str(r[col["ch"]] or "").strip()
            cat = str(r[col["cat"]] or "").strip().lower()
            rows.append({"Q": Path(rel).parts[0], "법인": SUB_FIX.get(sub, sub), "구분": "MX" if cat in MX_CATS else "CE",
                         "월": int(month[1:]), "Media Type 1": CHANNEL.get(ch.lower(), ch),
                         "플랜 Platform": str(r[col["pf"]] or "").strip(), "계획(USD)": float(cost), "플랜 파일": rel})
    return pd.DataFrame(rows)


def read_actuals(raws: dict[str, Path]) -> pd.DataFrame:
    parts = []
    for kind, path in raws.items():
        df = fastexcel.read_excel(str(path)).load_sheet("RAW").to_pandas()
        df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
        sp = next(c for c in df.columns if c.startswith("Media Spend (USD)"))
        d = pd.DataFrame({"법인": df["Subsidiary"].astype(str).str.strip().str.upper(),
                          "날짜": pd.to_datetime(df["Date"], errors="coerce"),
                          "Media Type 1": df["Media Type 1"].astype(str).str.strip(),
                          "Spend": pd.to_numeric(df[sp], errors="coerce").fillna(0)})
        d = d[d["날짜"].dt.year == 2026]
        d["월"], d["구분"] = d["날짜"].dt.month, kind
        parts.append(d)
    return pd.concat(parts, ignore_index=True)


def _compare(plan: pd.DataFrame, act: pd.DataFrame, unit: str) -> pd.DataFrame:
    """unit = '월' 또는 'Q' — 그 단위로 계획 vs 실적. 판정: 실적 Spend 합이 0 이하이면 '실적 없음'."""
    keys = ["법인", "구분", unit, "Media Type 1"]
    pg = plan.groupby(keys + ([] if unit == "Q" else ["Q"]), as_index=False).agg(
        **{"계획 월": ("월", lambda s: ", ".join(f"M{x:02d}" for x in sorted(set(s)))),
           "계획(USD)": ("계획(USD)", "sum"),
           "플랜 Platform": ("플랜 Platform", lambda s: ", ".join(sorted(set(s)))),
           "플랜 파일": ("플랜 파일", "first")})
    ag = act.groupby(keys, as_index=False).agg(**{"실적(USD)": ("Spend", "sum")})
    m = pg.merge(ag, on=keys, how="left").fillna({"실적(USD)": 0})
    media = (act[act["Spend"] > 0].groupby(["법인", "구분", unit])["Media Type 1"]
             .agg(lambda s: ", ".join(sorted(set(s)))).rename("그 기간 실적 매체"))
    last = act[act["Spend"] > 0].groupby(["법인", "구분"])["날짜"].max().dt.strftime("%Y-%m-%d").rename("마지막 실적일")
    m = m.merge(media, on=["법인", "구분", unit], how="left").merge(last, on=["법인", "구분"], how="left")
    m["판정"] = m["실적(USD)"].map(lambda v: "OK" if v > 0 else "실적 없음")
    return m.sort_values(["구분", "법인", unit, "Media Type 1"])


def run(plan_dir: Path, raws: dict[str, Path]) -> Path:
    plan = read_plans(plan_dir)
    act = read_actuals(raws)
    act["Q"] = act["월"].map(lambda x: f"{(x - 1) // 3 + 1}Q")
    m = _compare(plan, act, "Q")                 # 판정은 분기 단위(담당자 2026-10-27)
    monthly = _compare(plan, act, "월")           # 참고 — 월별
    gap = m[m["판정"] != "OK"]

    plans_by = plan.groupby(["Q", "법인", "구분"])["계획(USD)"].sum().round(0).reset_index()
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = REPORT_DIR / "미디어플랜점검.xlsx"
    with pd.ExcelWriter(out) as xw:
        gap.to_excel(xw, sheet_name="실적 없음", index=False)
        m.to_excel(xw, sheet_name="전체 비교(분기)", index=False)
        monthly.to_excel(xw, sheet_name="참고 월별", index=False)
        plans_by.to_excel(xw, sheet_name="플랜 파일 현황", index=False)
        plan.to_excel(xw, sheet_name="플랜 원본 행", index=False)

    lines = [f"# 미디어 플랜 점검 — 계획 있는 매체(Media Type 1)의 Weekly Raw 실적", "",
             f"- 플랜: `{plan_dir}` ({plan['플랜 파일'].nunique()}개 파일)",
             "- 실적: " + " · ".join(f"{k} `{p.name}`" for k, p in raws.items()), "",
             "## 실적 없음 — 법인별", ""]
    for (sub, kind), g in gap.groupby(["법인", "구분"]):
        lines.append(f"- **{sub} {kind}** (마지막 실적일 {g['마지막 실적일'].iloc[0]})")
        for t1, gg in g.groupby("Media Type 1"):
            qs = ", ".join(f"{q}({mm})" for q, mm in zip(gg["Q"], gg["계획 월"]))
            lines.append(f"  - {t1}: {qs} — 계획 ${gg['계획(USD)'].sum():,.0f} "
                         f"(플랜 Platform {', '.join(sorted(set(', '.join(gg['플랜 Platform']).split(', '))))})")
    lines += ["", "## 플랜 파일", "", "| Q | 법인 | 구분 | 계획(USD) |", "|---|---|---|---:|"]
    lines += [f"| {q} | {s} | {k} | {v:,.0f} |" for q, s, k, v in plans_by.itertuples(index=False, name=None)]
    (REPORT_DIR / "미디어플랜점검.md").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\n  ✔ {out}")
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--plan-dir", type=Path, default=PLAN_DIR)
    ap.add_argument("--mx", type=Path)
    ap.add_argument("--ce", type=Path)
    a = ap.parse_args(argv)
    raws = {k: p for k, p in (("MX", a.mx or latest_raw("MX")), ("CE", a.ce or latest_raw("CE"))) if p}
    run(a.plan_dir, raws)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
