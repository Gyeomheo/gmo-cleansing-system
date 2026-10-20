"""캠페인 규칙 제안 · 등록 — 규칙이 없는 캠페인 제품 코드(PR~)를 표로 만들고, O 표시한 것을 규칙표에 넣는다.

    python tools/campaign_rules.py propose
        2_Output/cleaned_*.csv(이번 입력분) → Reports/캠페인규칙제안.xlsx
        대상: 클렌징 후에도 시리즈 · 분류가 빈 값 · Other · Multi 로 남은 행의 PR 코드 (Spend 큰 순)
        제안값: 같은 PR 코드가 다른 행 · 법인에서 구체적인 값으로 들어온 경우 그중 Spend 80% 이상인 값
    python tools/campaign_rules.py apply
        제안표에서 '사용' 칸에 O 인 행을 3_Config/campaign_rules.xlsx 에 추가(백업 후)

규칙의 뜻 · 열 설명은 src/campaign.py 맨 위.
"""
from __future__ import annotations

import argparse
import shutil
from datetime import datetime

import pandas as pd

from common import REPORT_DIR, _read, config, is_locked
from src.pipeline import sanitize_column_headers
from src.campaign import CAT, COLS, ON, PRD, SER, extract_codes, is_generic, load_rules

SHEET = REPORT_DIR / "캠페인규칙제안.xlsx"
RULES = config.CAMPAIGN_RULES_EXCEL
BACKUP = config.CONFIG_DIR / "_backup"


def cleaned_frames() -> list[tuple[str, pd.DataFrame]]:
    """이번 1_Input 파일에 해당하는 cleaned_*.csv 만 (구분, df)."""
    out = []
    for inp in sorted(config.INPUT_DIR.glob("*.csv")):
        hits = sorted(config.OUTPUT_DIR.glob(f"cleaned_*~_{inp.stem}.csv"))
        if not hits:
            continue
        df = _read(hits[-1])
        if df is None or "Campaign Name" not in df.columns or CAT not in df.columns:
            continue
        # cleaned 파일은 MX 도 BU 열(MX)이 있으므로 구분은 입력 파일 머리줄로 본다(run.py 와 같은 기준)
        head = _read(inp)
        kind = "CE" if head is not None and config.COL_BU in sanitize_column_headers(head).columns else "MX"
        out.append((kind, df.fillna("")))
    return out


def _split_and_learn(allr: pd.DataFrame, generic: pd.Series):
    """(시리즈를 더 나눌 수 있는 분류인가, {(구분, PR): 제안값})."""
    # 시리즈를 더 나눌 수 있는 분류만 — 매핑에 Other · Multi 말고 다른 시리즈가 있는 분류
    # (CE 의 Heating & Cooling · Monitor · Dishwasher 등은 표준 시리즈가 Multi 하나뿐)
    splittable = {}
    for kind, path in (("MX", config.PRODUCT_MX_EXCEL), ("CE", config.PRODUCT_CE_EXCEL)):
        m = pd.read_excel(path).fillna("").astype(str)
        m = m[~is_generic(m["B_Standard"])]
        splittable[kind] = set(m["A_Standard"].str.lower())
    div_cats = {"MX": {c.lower() for c in config.DIV_RULES["MX"]},
                "CE": {c.lower() for d in ("VD", "DA") for c in config.DIV_RULES[d]}}
    cat_l = allr["cat"].str.strip().str.lower()
    can_split = pd.Series([c in splittable[k] or c in ("", "multi") for k, c in zip(allr["구분"], cat_l)],
                          index=allr.index)
    # 같은 PR 코드가 구체적인 값으로 들어온 행 → 제안값 (Spend 80% 이상 한 값, 그 구분의 분류일 때만)
    in_div = pd.Series([c in div_cats[k] for k, c in zip(allr["구분"], cat_l)], index=allr.index)
    spec = allr[~generic & in_div].groupby(["구분", "PR", "cat", "ser", "prd"])["spend"].sum().reset_index()
    learned = {}
    for (kind, pr), part in spec.groupby(["구분", "PR"]):
        top = part.sort_values("spend", ascending=False).iloc[0]
        tot = part["spend"].sum()
        if tot > 0 and top["spend"] / tot >= 0.8:
            learned[(kind, pr)] = (top["cat"], top["ser"], top["prd"], f"같은 코드 다른 행 {top['spend'] / tot:.0%}")

    return can_split, learned


def propose() -> object:
    frames = cleaned_frames()
    if not frames:
        print("■ 캠페인 규칙 제안 — cleaned 파일 없음")
        return None
    rows = []
    for kind, df in frames:
        c = extract_codes(df["Campaign Name"])
        sp = pd.to_numeric(df.get("Media Spend (USD)", 0), errors="coerce").fillna(0)
        rows.append(pd.DataFrame({
            "구분": kind, "법인": df["Subsidiary"].str.strip().str.upper(), "PR": c["PR"], "SB": c["SB"],
            "cat": df[CAT], "ser": df[SER], "prd": df[PRD] if PRD in df.columns else "",
            "rule": df.get("Campaign_Rule", ""), "spend": sp}))
    allr = pd.concat(rows, ignore_index=True)
    # 복수 제품 코드(vdmul · mxmulti · damul · multi …)는 원래 Multi 라 대상이 아니다
    allr = allr[(allr["PR"] != "") & ~allr["PR"].str.contains("mul")]
    allr["rule"] = allr["rule"].str.replace(" 충돌", "")
    generic = is_generic(allr["ser"]) | is_generic(allr["cat"])

    can_split, learned = _split_and_learn(allr, generic)

    # MX 파일에서 MX 제품으로 들어온 코드 — CE 파일에 나오면 MX 캠페인이 섞인 것(중복 집계 의심)
    mx_codes = set(allr.loc[(allr["구분"] == "MX") & ~generic, "PR"])
    rules = load_rules(RULES)
    have = set(zip(rules["구분"], rules["법인"], rules["PR"])) if not rules.empty else set()
    todo = allr[generic & can_split & (allr["rule"] == "")]
    g = (todo.groupby(["구분", "법인", "PR", "SB", "cat", "ser"])
         .agg(행=("spend", "size"), spend=("spend", "sum")).reset_index()
         .sort_values("spend", ascending=False))
    out = []
    for r in g.itertuples(index=False):
        if (r.구분, r.법인, r.PR) in have or (r.구분, "", r.PR) in have:
            continue        # 규칙은 있는데 조건(분류 불일치 등)이 안 맞은 것 — 규칙표에서 직접 본다
        cat, ser, prd, why = learned.get((r.구분, r.PR), ("", "", "", ""))
        if r.구분 == "CE" and r.PR in mx_codes:
            why = "⚠ MX 제품 코드 — CE 파일에 MX 캠페인이 섞임(MX 파일과 중복인지 확인)"
        out.append(["", r.법인, r.구분, r.PR, "", "", "", cat, ser, prd, "", "", why,
                    f"{r.cat} / {r.ser}", r.행, round(r.spend), r.SB])
    if not out:
        print("■ 캠페인 규칙 제안 — 남은 코드 없음")
        if SHEET.exists() and not is_locked(SHEET):
            SHEET.unlink()
        return None
    df = pd.DataFrame(out, columns=COLS + ["지금 결과", "행", "Spend(USD)", "SB 예"])
    if is_locked(SHEET):
        print(f"  ✖ {SHEET.name} 이(가) 열려 있어 새로 못 씀 — 닫고 다시")
        return None
    with pd.ExcelWriter(SHEET, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="제안")
        ws = w.sheets["제안"]
        for i, width in enumerate((6, 8, 6, 12, 10, 10, 10, 16, 18, 16, 10, 8, 40, 30, 7, 11, 12)):
            ws.column_dimensions[chr(65 + i)].width = width
        ws.freeze_panes = "E2"
    print(f"■ 캠페인 규칙 제안 — 코드 {len(df)}건(Spend 큰 순) → {SHEET}")
    print("  Category · Series 채우고 '사용' 칸에 O → python tools/campaign_rules.py apply → 다시 실행(--remap)")
    return SHEET


def apply_sheet() -> int:
    if not SHEET.exists() or is_locked(SHEET) or is_locked(RULES):
        print("  ✖ 제안표가 없거나, 제안표 · 규칙표가 Excel 에서 열려 있음")
        return 0
    df = pd.read_excel(SHEET, dtype=str).fillna("")
    pick = df[df["사용"].str.strip().str.upper().isin(ON)]
    pick = pick[(pick["Category"] != "") | (pick["Series"] != "")]
    if pick.empty:
        print("  '사용' 칸에 O 이고 Category · Series 를 채운 행이 없습니다.")
        return 0
    old = pd.read_excel(RULES, dtype=str).fillna("") if RULES.exists() else pd.DataFrame(columns=COLS)
    if RULES.exists():
        BACKUP.mkdir(exist_ok=True)
        shutil.copy2(RULES, BACKUP / f"{RULES.stem}_{datetime.now():%Y%m%d_%H%M%S}{RULES.suffix}")
    new = pd.concat([old, pick[COLS]], ignore_index=True)
    with pd.ExcelWriter(RULES, engine="openpyxl") as w:
        new.to_excel(w, index=False, sheet_name="규칙")
        ws = w.sheets["규칙"]
        for col, wd in zip("ABCDEFGHIJKLM", (6, 8, 6, 12, 12, 18, 14, 16, 18, 16, 12, 10, 80)):
            ws.column_dimensions[col].width = wd
        ws.freeze_panes = "A2"
    for r in pick.itertuples(index=False):
        print(f"  ✔ {r.법인 or '*'} {r.구분} PR~{r.PR} → {r.Category} / {r.Series}")
    return len(pick)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["propose", "apply"])
    a = ap.parse_args()
    raise SystemExit(0 if (propose() if a.cmd == "propose" else apply_sheet()) else 1)
