"""매핑 자동 추론 — 매핑 키에 없는 원본 값 중 법칙으로 확실한 것만 매핑 엑셀에 자동 등록한다.

    python tools/auto_mapping.py            미리보기(등록 안 함)
    python tools/auto_mapping.py --apply    매핑 엑셀에 등록(백업 후) — run_all 이 클렌징 전에 부른다

기존 클렌징 로직은 그대로다. 매핑 엑셀에 행을 더할 뿐이고, 더한 행은 '메모' 열에 '자동: <법칙>' 으로 표시된다.

법칙 (위에서부터 처음 맞는 것 하나)
  0. 별칭 풀기   3_Config/aliases.csv 의 코드명을 정식 이름으로 (Q8 → Z Fold8 Ultra, H8 → Z Fold8, B8 → Z Flip8)
  1. 같은 키      띄어쓰기 · 대소문자 · 기호만 다른 기존 키 → 그 키의 규칙 그대로
                 (Watch 9 = Watch9, Tab S10 Fe = Tab S10 FE)
  2. 표준 제품명  표준 Products(C) 와 같으면 → 그 표준                              (Z Fold 8 = Z Fold8)
  3. 표준 시리즈  표준 Series(B) 안에 단어로 들어 있고 후보가 하나뿐이면 → 그 시리즈  (Watch 7 → Watch7 Series)
  4. 이전 이력    지난 cleaned 파일에서 같은 원본 값(띄어쓰기 무시)이 90% 이상 한 결과로 갔고,
                 그 결과가 매핑표의 표준값이면 → 그 결과 (원본이 그대로 지나간 오타 이력은 제외)
자동 등록 안 함: 일반값(Other · Multi · Tablet …), 3글자 미만, 후보가 둘 이상 → 매핑등록.xlsx 제안으로만
"""
from __future__ import annotations

import argparse
import re

import pandas as pd

from common import _read, config
from mapping import KINDS, add_rules, is_generic, load_rules, scan_new_values
from src.campaign import is_generic as generic_s

ALIASES = config.CONFIG_DIR / "aliases.csv"
GENERIC_WORDS = {"other", "others", "multi", "none", "tablet", "tablets", "smartphone", "smartphones",
                 "pc", "wearable", "wearables", "hearables", "all", "brand", "generic"}


def compact(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v).lower())


def load_aliases() -> dict[str, str]:
    if not ALIASES.exists():
        return {}
    df = pd.read_csv(ALIASES, dtype=str, encoding="utf-8-sig").fillna("")
    return {compact(a): b.strip() for a, b in zip(df.iloc[:, 0], df.iloc[:, 1]) if compact(a) and b.strip()}


def expand_aliases(value: str, aliases: dict[str, str]) -> str:
    """'Q8 & H8' → 'Z Fold8 Ultra & Z Fold8'. 별칭이 없으면 그대로."""
    tokens = re.split(r"[\s_&+,/|]+", str(value).strip())
    if not aliases or not any(compact(t) in aliases for t in tokens if t):
        return value
    return " & ".join(aliases.get(compact(t), t) for t in tokens if t)


def history_pairs(kind: str) -> dict[str, tuple[list[str], float]]:
    """지난 cleaned 파일들: compact(원본 값) → (가장 많은 구체적 결과, 비율)."""
    _, cols = KINDS[kind]
    raw_col = {"MX": "Products", "CE": "Product Category", "MEDIA": "Media Platform"}[kind]
    out_cols = [f"{c}_cleaned" for c in cols["raw_cols"]]
    folders = [config.OUTPUT_DIR] + sorted(config.OUTPUT_DIR.glob("원본보존_*"))
    parts = []
    for folder in folders:
        for f in folder.glob("cleaned_*.csv"):
            if "_Summary_" in f.name:
                continue
            df = _read(f)
            if df is None or raw_col not in df.columns or any(c not in df.columns for c in out_cols):
                continue
            is_ce = config.COL_BU in df.columns and set(df[config.COL_BU].dropna().unique()) - {"MX"}
            if kind in ("MX", "CE") and bool(is_ce) != (kind == "CE"):
                continue
            sp = pd.to_numeric(df.get("Media Spend (USD)", 0), errors="coerce").fillna(0)
            parts.append(pd.DataFrame({"raw": df[raw_col].fillna("").map(compact),
                                       "res": df[out_cols].fillna("").astype(str).agg(" | ".join, axis=1),
                                       "ser": df[out_cols[1]].fillna(""), "spend": sp.abs() + 1e-9}))
    if not parts:
        return {}
    h = pd.concat(parts, ignore_index=True)
    h = h[(h["raw"].str.len() >= 3) & ~generic_s(h["ser"])]
    out = {}
    for raw, g in h.groupby("raw"):
        s = g.groupby("res")["spend"].sum().sort_values(ascending=False)
        out[raw] = (s.index[0].split(" | "), float(s.iloc[0] / s.sum()))
    return out


def infer_one(kind: str, value: str, rules: pd.DataFrame, aliases: dict, hist: dict) -> tuple[list[str], str]:
    kcol, std = KINDS[kind][1]["key"], KINDS[kind][1]["std_cols"]
    v = expand_aliases(value, aliases)
    alias_note = f"별칭 {value} → {v}, " if v != value else ""
    c = compact(v)
    if len(c) < 3 or c in GENERIC_WORDS or is_generic(v):
        return [], ""
    ck = rules[kcol].map(compact)
    hit = rules[ck == c]
    if not hit.empty:
        uniq = hit[std].drop_duplicates()
        if len(uniq) == 1:
            return uniq.iloc[0].tolist(), alias_note + f"1 같은 키 '{hit.iloc[0][kcol]}'"
        return [], ""                                   # 같은 키에 규칙이 여럿 → 사람이
    if kind != "CE":                                     # CE 는 키 = 분류라 2 · 3 이 맞지 않는다
        for i, name in ((2, "2 표준 제품명"), (1, "3 표준 시리즈")):
            col = std[i]
            if i == 2:
                m = rules[rules[col].map(compact) == c]
            else:
                pat = rf"(?:^|[^a-z0-9]){re.escape(c)}(?:$|[^a-z0-9])"
                squashed = rules[col].str.lower().str.replace(r"\s+", "", regex=True)
                m = rules[squashed.str.contains(pat, regex=True)
                          & ~rules[col].map(is_generic)]
            m = m[~m[std[1]].map(is_generic)]
            if m.empty:
                continue
            uniq = m[std[:2]].drop_duplicates()
            if len(uniq) != 1:
                return [], ""
            a, b = uniq.iloc[0].tolist()
            prod = m.iloc[0][std[2]] if i == 2 else b
            return [a, b, prod], alias_note + f"{name} '{m.iloc[0][col]}'"
    if c in hist:
        res, share = hist[c]
        # 결과가 매핑표의 표준값일 때만 — 못 찾아서 원본이 그대로 지나간 이력(오타 등)은 배우지 않는다
        sets = [set(rules[s_]) for s_ in std]
        valid = (len(res) == 3 and res[0] in sets[0] and res[1] in sets[1]
                 and (res[2] in sets[2] or (kind == "MEDIA" and res[0] == "Affiliate")))
        if share >= 0.9 and valid:
            return res, alias_note + f"4 이전 이력 {share:.0%}"
    return [], ""


def infer() -> pd.DataFrame:
    g = scan_new_values()
    if g.empty:
        return pd.DataFrame()
    g = g[g["구분"] != "MEDIA"]          # 미디어는 매핑을 늘리지 않는다 — 표준표 검사(src/media_std.py)가 Others 로
    aliases = load_aliases()
    rules = {k: load_rules(k) for k in KINDS}
    hist = {k: history_pairs(k) for k in KINDS}
    rows = []
    for (kind, key), part in g.groupby(["구분", "등록 키"], sort=False):
        top = part.iloc[0]
        std, why = infer_one(kind, top["원본 값"], rules[kind], aliases, hist[kind])
        if std:
            if kind == "CE":
                std = [std[0], std[1], "NONE"]
            rows.append([kind, top["원본 값"], key, *std, why, round(part["spend"].sum()),
                         ", ".join(sorted({s for x in part["법인"] for s in x.split(", ")}))])
    cols = ["구분", "원본 값", "등록 키", "표준1", "표준2", "표준3", "법칙", "Spend", "법인"]
    return pd.DataFrame(rows, columns=cols)


def run(apply: bool) -> int:
    df = infer()
    if df.empty:
        print("■ 매핑 자동 추론 — 확실한 것 없음")
        return 0
    print(f"■ 매핑 자동 추론 — {len(df)}건" + (" 등록" if apply else " (미리보기, --apply 로 등록)"))
    recs = df.sort_values("Spend", ascending=False).to_dict("records")
    for r in recs:
        print(f"  {r['구분']:<5} {r['원본 값']:<26} → {r['표준1']} / {r['표준2']} / {r['표준3']}"
              f"   [{r['법칙']}]  ${r['Spend']:,} {r['법인']}")
    if apply:
        add_rules([(r["구분"], r["원본 값"], [r["표준1"], r["표준2"], r["표준3"]]) for r in recs],
                  notes=[f"자동: {r['법칙']}" for r in recs])
    return len(df)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--apply", action="store_true")
    run(ap.parse_args().apply)
