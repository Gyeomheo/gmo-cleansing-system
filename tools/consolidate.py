"""취합 — 이번 클렌징 결과를 Weekly Dashboard RAW 취합본에 법인별로 반영한 새 버전을 만든다.

    python tools/consolidate.py CE              미리보기 + 검증만 (파일 안 만듦)
    python tools/consolidate.py CE --apply      검증 통과 시 새 취합본 생성
    python tools/consolidate.py MX --apply --week 44

기준 데이터 = 가장 최근 취합본의 RAW 시트 (로컬 TOTAL_MASTER 가 아니다 — 마스터에는 빠진 달이 있다).
법인별로 기존 updater 와 같은 방식: 숫자(Spend · Revenue · Imp · Click · Order)가 처음 달라진 날부터
이번 결과로 교체하고, 그 앞 행과 이번에 없는 법인은 RAW 그대로 둔다. 그 뒤 전체 RAW 에 미디어 표준표.

만들기 (--apply, 담당자 작업 방식 그대로)
  1. 지금 최신 취합본을 'old RAW files' 에 복사
  2. 새 이름 <YYMMDD>_Samsung D2C (2026) - W<nn>_v<k>_GMO_<MX|CE>_RAW File.xlsx 로 복사 → Excel 로 RAW 시트 교체
     · 피벗(KPI · pacing · status …) 원본 범위를 새 행 수로 바꾸고 새로고침 → 저장
  3. 이전 최신 파일은 D2C 폴더에서 'old RAW files' 로 (1 의 복사본이 있으므로 덮어쓰지 않고 건너뜀)
검증 결과는 2_Output/Reports/취합검증_<MX|CE>.md — 오류가 하나라도 있으면 파일을 만들지 않는다.
"""
from __future__ import annotations

import argparse
import re
import shutil
from datetime import date, datetime
from pathlib import Path

import fastexcel
import numpy as np
import pandas as pd

from common import REPORT_DIR, config, is_locked
from src.media_std import PF, T1, T2, load_taxonomy, standardize_media
from src.pipeline import sanitize_column_headers
from src.updater import force_numeric_metrics, get_safe_divergence_date

DASH_DIR = Path(r"G:\공유 드라이브\Samsung Perform_1\문서-1\2026\09. GMO Report\3. Weekly Dashboard RAW File\D2C")
OLD_DIR = DASH_DIR / "old RAW files"
NAME_RE = re.compile(r"^(\d{6})_Samsung D2C \(2026\) - W(\d+)_v(\d+)_GMO_(MX|CE)_RAW File\.xlsx$")
METRICS = ["Media Spend (USD)", "Revenue", "Impressions", "Clicks", "Orders"]


# ---------------------------------------------------------------- 파일 찾기
def latest_file(kind: str) -> Path | None:
    best = None
    for f in DASH_DIR.glob("*.xlsx"):
        m = NAME_RE.match(f.name)
        if m and m.group(4) == kind:
            key = (int(m.group(2)), int(m.group(3)))
            if best is None or key > best[0]:
                best = (key, f)
    return best[1] if best else None


def new_name(kind: str, week: int) -> str:
    v = 1
    for f in list(DASH_DIR.glob("*.xlsx")) + list(OLD_DIR.glob("*.xlsx")):
        m = NAME_RE.match(f.name)
        if m and m.group(4) == kind and int(m.group(2)) == week:
            v = max(v, int(m.group(3)) + 1)
    return f"{date.today():%y%m%d}_Samsung D2C (2026) - W{week}_v{v}_GMO_{kind}_RAW File.xlsx"


# ---------------------------------------------------------------- 데이터
def read_raw(path: Path) -> tuple[pd.DataFrame, list[str]]:
    df = fastexcel.read_excel(str(path)).load_sheet("RAW").to_pandas()
    header = list(df.columns)
    df.columns = sanitize_column_headers(df.iloc[0:0].copy()).columns  # 이름만 표준화(줄바꿈 등)
    return df, header


def _norm(df: pd.DataFrame) -> pd.DataFrame:
    """비교용 사본 — updater 와 같은 정리(날짜 · 주차 · 법인 · 숫자 반올림)."""
    d = df.copy()
    d["Date"] = pd.to_datetime(d["Date"], errors="coerce")
    # 두 자리로 — updater 는 주차를 문자열로 정렬하므로 '10' 이 '4' 보다 앞에 온다(→ 이른 변경을 놓침)
    d["Week"] = pd.to_numeric(d["Week"], errors="coerce").fillna(0).astype(int).map("{:02d}".format)
    d["Subsidiary"] = d["Subsidiary"].fillna("Unknown").astype(str).str.strip()
    return force_numeric_metrics(d)


def merge(raw: pd.DataFrame, new: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """법인별 교체. 원본 행은 값 그대로 쓴다(비교만 정리된 사본으로)."""
    raw_n, new_n = _norm(raw), _norm(new)
    parts, plan = [], []
    subs = list(dict.fromkeys(new_n["Subsidiary"]))
    keep = ~raw_n["Subsidiary"].isin(subs)
    parts.append(raw[keep])
    for sub in subs:
        s_new = new_n[new_n["Subsidiary"] == sub]
        s_old = raw_n[raw_n["Subsidiary"] == sub]
        if s_old.empty:
            cutoff = s_new["Date"].min()
        else:
            cutoff = get_safe_divergence_date(s_old.sort_values("Date"), s_new.sort_values("Date"))
        old_keep = (raw_n["Subsidiary"] == sub) & (raw_n["Date"] < cutoff)
        new_take = (new_n["Subsidiary"] == sub) & (new_n["Date"] >= cutoff)
        parts += [raw[old_keep], new[new_take]]
        replaced = (raw_n["Subsidiary"] == sub) & (raw_n["Date"] >= cutoff)
        plan.append({"법인": sub, "교체 시작": cutoff.date() if pd.notna(cutoff) else None,
                     "RAW 행(교체)": int(replaced.sum()), "새 행": int(new_take.sum()),
                     "Spend 전": float(raw_n.loc[replaced, METRICS[0]].sum()),
                     "Spend 후": float(new_n.loc[new_take, METRICS[0]].sum()),
                     "새 데이터 마지막 날": s_new["Date"].max().date()})
    out = pd.concat(parts, ignore_index=True, sort=False)
    out = out.reindex(columns=raw.columns)
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = out.sort_values(["Date", "Subsidiary"], kind="stable").reset_index(drop=True)
    return out, plan


def drop_duplicates(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """완전히 같은 행은 하나만 남긴다(담당자 2026-10-20). 뺀 행 수 · 숫자 합계를 돌려준다."""
    dup = df.duplicated()
    gone = df[dup]
    info = {"행": int(dup.sum()),
            "숫자 있는 행": int((gone[[m for m in ("Media Spend (USD)", "Revenue", "Clicks") if m in gone.columns]]
                              .apply(pd.to_numeric, errors="coerce").abs().sum(axis=1) > 0).sum()),
            "Spend": float(pd.to_numeric(gone.get("Media Spend (USD)", 0), errors="coerce").sum()),
            "Revenue": float(pd.to_numeric(gone.get("Revenue", 0), errors="coerce").sum()),
            "법인별": gone.groupby("Subsidiary").size().to_dict() if len(gone) else {}}
    return df[~dup].reset_index(drop=True), info


def canonical_case(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    """분류 · 시리즈가 매핑표 표준값과 대소문자만 다르면 표준 표기로 (SEI 'multi' → 'Multi')."""
    m = pd.read_excel(config.PRODUCT_CE_EXCEL if kind == "CE" else config.PRODUCT_MX_EXCEL).fillna("")
    for col, std in (("Product Category", "A_Standard"), ("Product Series", "B_Standard")):
        canon = {str(v).strip().lower(): str(v).strip() for v in m[std] if str(v).strip()}
        low = df[col].fillna("").astype(str).str.strip().str.lower()
        hit = low.isin(canon) & (df[col].fillna("").astype(str).str.strip() != low.map(canon))
        df.loc[hit, col] = low[hit].map(canon)
    return df


def recompute_cpc(df: pd.DataFrame) -> pd.DataFrame:
    """CPC = IFERROR(Spend / Clicks, 0) — 전체 행 (담당자 2026-10-20)."""
    sp = pd.to_numeric(df["Media Spend (USD)"], errors="coerce").fillna(0)
    ck = pd.to_numeric(df["Clicks"], errors="coerce").fillna(0)
    df["CPC"] = np.where(ck != 0, sp / ck.replace(0, np.nan), 0.0)
    df["CPC"] = pd.Series(df["CPC"]).replace([np.inf, -np.inf], 0).fillna(0).values
    return df


def media_pass(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    tax = load_taxonomy(config.MEDIA_TAXONOMY)
    cols = {"Media Type 1": T1, "Media Type 2": T2, "Media Platform": PF}
    work = df[list(cols)].rename(columns=cols)
    work["Media Spend (USD)"] = df["Media Spend (USD)"]
    fixed, _ = standardize_media(work, tax)
    changed = (fixed[list(cols.values())].values != work[list(cols.values())].fillna("").astype(str).values)
    n = int(changed.any(axis=1).sum())
    for c, w in cols.items():
        df[c] = fixed[w].values
    return df, n


# ---------------------------------------------------------------- 검증
def validate(raw: pd.DataFrame, out: pd.DataFrame, new: pd.DataFrame, plan: list[dict]) -> list[tuple[str, str, str]]:
    """[(결과 OK/경고/오류, 항목, 내용)]"""
    res = []
    raw_n, out_n, new_n = _norm(raw), _norm(out), _norm(new)
    changed_subs = {p["법인"] for p in plan}

    same = True
    for sub in sorted(set(raw_n["Subsidiary"]) - changed_subs):
        a, b = raw_n[raw_n["Subsidiary"] == sub], out_n[out_n["Subsidiary"] == sub]
        if len(a) != len(b) or abs(a[METRICS[0]].sum() - b[METRICS[0]].sum()) > 0.01:
            same = False
            res.append(("오류", "안 바뀐 법인 보존", f"{sub}: 행 {len(a)}→{len(b)}"))
    if same:
        kept = len(set(raw_n["Subsidiary"]) - changed_subs)
        res.append(("OK", "안 바뀐 법인 보존", f"{kept}개 법인 행 · Spend 그대로"))

    for p in plan:
        sub, cut = p["법인"], pd.Timestamp(p["교체 시작"])
        a = raw_n[(raw_n["Subsidiary"] == sub) & (raw_n["Date"] < cut)]
        b = out_n[(out_n["Subsidiary"] == sub) & (out_n["Date"] < cut)]
        c = new_n[(new_n["Subsidiary"] == sub) & (new_n["Date"] >= cut)]
        d = out_n[(out_n["Subsidiary"] == sub) & (out_n["Date"] >= cut)]
        ok_before = len(a) == len(b) and abs(a[METRICS[0]].sum() - b[METRICS[0]].sum()) < 0.01
        ok_after = all(abs(c[m].sum() - d[m].sum()) < 0.5 for m in METRICS)
        res.append(("OK" if ok_before and ok_after else "오류", f"{sub} 교체",
                    f"{p['교체 시작']}~ 교체, 그 앞 {len(b):,}행 그대로, 이후 합계 = 이번 결과"
                    if ok_before and ok_after else f"{p['교체 시작']} 기준 합계 불일치"))
        total_before = raw_n.loc[raw_n["Subsidiary"] == sub, METRICS[0]].sum()
        delta = p["Spend 후"] - p["Spend 전"]
        if total_before and abs(delta) / total_before > 0.3:
            res.append(("경고", f"{sub} 변화 큼", f"Spend {delta:+,.0f} (법인 전체의 {delta / total_before:+.0%})"))

    sp = pd.to_numeric(out["Media Spend (USD)"], errors="coerce").fillna(0)
    ck = pd.to_numeric(out["Clicks"], errors="coerce").fillna(0)
    expect = np.where(ck != 0, sp / ck.replace(0, np.nan), 0.0)
    got = np.round(pd.to_numeric(out["CPC"], errors="coerce").fillna(0), 2)
    bad_cpc = int((got != np.round(np.nan_to_num(expect), 2)).sum())
    res.append(("오류" if bad_cpc else "OK", "CPC", f"IFERROR(Spend/Clicks, 0) 와 소수 둘째 자리 불일치 {bad_cpc}행"))

    dmin, dmax = out_n["Date"].min(), out_n["Date"].max()
    bad_date = out_n["Date"].isna().sum()
    future = (out_n["Date"] > pd.Timestamp(date.today())).sum()
    early = (out_n["Date"] < pd.Timestamp(config.MIN_DATE)).sum()
    res.append(("오류" if bad_date or future else ("경고" if early else "OK"), "날짜",
                f"{dmin.date()} ~ {dmax.date()}, 빈 날짜 {bad_date} · 미래 {future} · {config.MIN_DATE} 이전 {early}"))
    blank_sub = (out["Subsidiary"].fillna("").astype(str).str.strip() == "").sum()
    res.append(("경고" if blank_sub else "OK", "법인 빈 값", f"{blank_sub}행"))

    tax = load_taxonomy(config.MEDIA_TAXONOMY)
    ok = set(zip(tax["t1"], tax["t2"], tax["pf"]))
    aff = {(a, b) for a, b, c in ok if c == "*"}
    combos = list(zip(out["Media Type 1"].fillna(""), out["Media Type 2"].fillna(""),
                      out["Media Platform"].fillna("")))
    outside = sum(1 for a, b, c in combos if (a, b, c) not in ok and (a, b) not in aff and a)
    res.append(("경고" if outside else "OK", "미디어 표준표", f"표 밖 조합 {outside}행"))

    def dups(df):
        d = df[df.duplicated()]
        live = d[[m for m in ("Media Spend (USD)", "Revenue", "Clicks") if m in d.columns]].abs().sum(axis=1) > 0
        return int(live.sum()), int((~live).sum())
    (lb, zb), (la, za) = dups(raw), dups(out)
    res.append(("경고" if la > lb else "OK", "중복 행(숫자 있음)",
                f"{lb} → {la}  — 늘면 같은 데이터가 두 번 들어간 것"))
    res.append(("OK", "중복 행(숫자 0)", f"{zb} → {za}  — 원본에 있는 빈 행, 합계에 영향 없음"))
    return res


def write_report(kind: str, src: Path, name: str, plan: list[dict], checks: list, media_n: int, n_rows: int,
                 dup_raw: dict, dup_new: dict) -> Path:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    md = [f"# 취합 검증 — {kind} · {datetime.now():%Y-%m-%d %H:%M}", "",
          f"- 기준 취합본: `{src.name}`", f"- 새 취합본: `{name}` ({n_rows:,}행)",
          f"- 미디어 표준표로 바꾼 행: {media_n:,}",
          f"- 완전 중복 제거 — 기준 RAW: {dup_raw['행']:,}행(숫자 있는 행 {dup_raw['숫자 있는 행']:,}, "
          f"Spend {dup_raw['Spend']:,.0f}, Revenue {dup_raw['Revenue']:,.0f}) {dup_raw['법인별']}",
          f"- 완전 중복 제거 — 이번 결과: {dup_new['행']:,}행(숫자 있는 행 {dup_new['숫자 있는 행']:,}) "
          f"{dup_new['법인별']}",
          "",
          "## 법인별 반영", "",
          "| 법인 | 교체 시작 | 새 데이터 마지막 날 | RAW 행(교체) | 새 행 | Spend 전 | Spend 후 |",
          "|---|---|---|---:|---:|---:|---:|"]
    for p in plan:
        md.append(f"| {p['법인']} | {p['교체 시작']} | {p['새 데이터 마지막 날']} | {p['RAW 행(교체)']:,} "
                  f"| {p['새 행']:,} | "
                  f"{p['Spend 전']:,.0f} | {p['Spend 후']:,.0f} |")
    md += ["", "## 검증", "", "| 결과 | 항목 | 내용 |", "|---|---|---|"]
    md += [f"| {'✔' if r == 'OK' else ('⚠' if r == '경고' else '✖')} {r} | {k} | {v} |" for r, k, v in checks]
    out = REPORT_DIR / f"취합검증_{kind}.md"
    out.write_text("\n".join(md) + "\n", encoding="utf-8")
    return out


# ---------------------------------------------------------------- Excel 쓰기
def write_workbook(src: Path, dest: Path, header: list[str], df: pd.DataFrame) -> None:
    """src 를 dest 로 복사 → Excel 로 RAW 시트 교체 · 피벗 원본 범위 갱신 · 새로고침 · 저장."""
    import pythoncom
    import win32com.client

    excel_epoch = pd.Timestamp("1899-12-30")

    def cell(v):
        """Excel(COM) 이 받는 기본형으로 — 날짜는 Excel 일련번호(COM 에 datetime 을 주면 시간대 변환으로
        9시간 당겨진다), numpy 수 → int/float, NaN/NaT → 빈 칸."""
        if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT:
            return None
        if isinstance(v, pd.Timestamp):
            return float((v.normalize() - excel_epoch).days)
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            return None if np.isnan(v) else float(v)
        return v

    shutil.copy2(src, dest)
    rows = [tuple(header)] + [tuple(cell(v) for v in r) for r in df.itertuples(index=False, name=None)]
    nrow, ncol = len(rows), len(header)

    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(str(dest))
        ws = wb.Worksheets("RAW")
        ws.Cells.ClearContents()
        date_col = [i for i, c in enumerate(df.columns) if c == "Date"]
        step = 20000
        for i in range(0, nrow, step):
            chunk = rows[i:i + step]
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + len(chunk), ncol)).Value = chunk
        for j in date_col:                        # 원본 RAW 와 같은 날짜 서식
            ws.Range(ws.Cells(2, j + 1), ws.Cells(nrow, j + 1)).NumberFormat = "yyyy-mm-dd"
        src_ref = f"RAW!R1C1:R{nrow}C{ncol}"
        # 피벗: 새 캐시 하나를 만들어 첫 피벗에 붙이고 나머지는 그 캐시를 같이 쓰게 한다
        # (캐시 원본 범위를 직접 바꾸는 건 Excel 이 막고, 피벗마다 지정하면 캐시가 갈라져 파일이 커진다)
        cache = wb.PivotCaches().Create(SourceType=1, SourceData=src_ref)     # 1 = xlDatabase
        first = None
        for sh in wb.Worksheets:
            for j in range(1, sh.PivotTables().Count + 1):
                pt = sh.PivotTables().Item(j)
                if first is None:
                    pt.ChangePivotCache(cache)
                    first = pt
                else:
                    pt.CacheIndex = first.CacheIndex
        if first is not None:
            first.PivotCache().Refresh()
        wb.RefreshAll()                           # 피벗 전부 최신화(같은 캐시라도 한 번 더)
        xl.CalculateUntilAsyncQueriesDone()
        bad = [f"{sh.Name}/{sh.PivotTables().Item(j).Name}"
               for sh in wb.Worksheets for j in range(1, sh.PivotTables().Count + 1)
               if sh.PivotTables().Item(j).SourceData.replace("'", "") != src_ref]
        if bad:
            raise RuntimeError("피벗 원본 범위가 새 RAW 와 다름: " + ", ".join(bad))
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()


def verify_written(dest: Path, out: pd.DataFrame) -> list[str]:
    """저장한 파일을 다시 읽어 대조 — RAW 행 수 · Spend 합계 · 피벗 캐시 행 수 · 캐시 개수."""
    import zipfile
    errors = []
    raw = fastexcel.read_excel(str(dest)).load_sheet("RAW").to_pandas()
    raw.columns = sanitize_column_headers(raw.iloc[0:0].copy()).columns
    if len(raw) != len(out):
        errors.append(f"RAW 행 {len(raw):,} ≠ {len(out):,}")
    else:
        a_d = pd.to_datetime(raw["Date"], errors="coerce").reset_index(drop=True)
        b_d = pd.to_datetime(out["Date"], errors="coerce").reset_index(drop=True)
        n_bad = int((a_d != b_d).sum())
        if n_bad:
            errors.append(f"날짜가 다른 행 {n_bad:,} (예: {b_d[a_d != b_d].iloc[0]} → {a_d[a_d != b_d].iloc[0]})")
        for col in out.columns:                   # 모든 열 · 모든 행
            if col == "Date" or col not in raw.columns:
                continue
            x, y = raw[col].reset_index(drop=True), out[col].reset_index(drop=True)
            xn, yn = pd.to_numeric(x, errors="coerce"), pd.to_numeric(y, errors="coerce")
            num = yn.notna() | xn.notna()
            bad_num = num & ~np.isclose(xn.fillna(0), yn.fillna(0), atol=1e-6)
            xs = x.where(x.notna(), "").astype(str).str.strip()
            ys = y.where(y.notna(), "").astype(str).str.strip()
            bad_str = ~num & (xs != ys)
            n = int((bad_num | bad_str).sum())
            if n:
                i = (bad_num | bad_str).idxmax()
                errors.append(f"{col} 다른 행 {n:,} (예: {y[i]!r} → {x[i]!r})")
    a = pd.to_numeric(raw["Media Spend (USD)"], errors="coerce").sum()
    b = pd.to_numeric(out["Media Spend (USD)"], errors="coerce").sum()
    if abs(a - b) > 0.5:
        errors.append(f"RAW Spend {a:,.2f} ≠ {b:,.2f}")
    z = zipfile.ZipFile(dest)
    defs = [n for n in z.namelist() if "pivotCacheDefinition" in n and n.endswith(".xml")]
    for n in defs:
        m = re.search(r'recordCount="(\d+)"', z.read(n).decode("utf8", "ignore"))
        if m and int(m.group(1)) != len(out):
            errors.append(f"피벗 캐시 행 {m.group(1)} ≠ {len(out):,}")
    if len(defs) > 1:
        errors.append(f"피벗 캐시가 {len(defs)}개로 갈라짐")
    return errors


# ---------------------------------------------------------------- 실행
def create(src: Path, name: str, header: list[str], out: pd.DataFrame) -> int:
    """새 취합본 만들기 — 이전 것은 old RAW files 로, 저장 후 대조 실패 시 새 파일 삭제."""
    if is_locked(src):
        print(f"  ⚠ {src.name} 이 열려 있다는 표시(~$)가 있음 — 원본은 복사만 하므로 계속")
    OLD_DIR.mkdir(exist_ok=True)
    backup = OLD_DIR / src.name
    if not backup.exists():
        shutil.copy2(src, backup)
    dest = DASH_DIR / name
    print(f"  만드는 중: {name}  (Excel — 1~2분)")
    try:
        write_workbook(src, dest, header, out)
        errors = verify_written(dest, out)
    except Exception:
        if dest.exists():
            dest.unlink()        # 만들다 만 새 파일만 지운다 — 이전 취합본은 그대로
        raise
    if errors:
        dest.unlink()
        print("  ✖ 저장 후 대조 실패 — 새 파일을 지우고 이전 취합본을 그대로 둠: " + " / ".join(errors))
        return 1
    print("  ✔ 저장 후 대조: RAW 행 · Spend 합계 · 피벗 캐시(1개, 행 수 일치)")
    try:                                          # 메일용 Collection status (status 피벗 기준)
        import tracker
        m = NAME_RE.match(name)
        tracker.run(m.group(4), int(m.group(2)), path=dest, prev=backup)
    except Exception as e:                        # Tracker 실패는 취합본을 막지 않는다
        print(f"  ⚠ Tracker 못 만듦: {e}")
    if backup.exists() and src.exists():
        src.unlink()             # 같은 파일이 old RAW files 에 있으므로 D2C 폴더에서는 뺀다(=옮김)
    print(f"  ✔ 새 취합본: {dest}\n  ✔ 이전 취합본: {backup}")
    return 0


def run(kind: str, apply: bool, week: int | None = None, dash_dir: Path | None = None) -> int:
    global DASH_DIR, OLD_DIR
    if dash_dir:                                  # 시험용 — 실제 공유 드라이브 대신 사본 폴더
        DASH_DIR, OLD_DIR = Path(dash_dir), Path(dash_dir) / "old RAW files"
    kind = kind.upper()
    print(f"■ 취합 — {kind}")
    src = latest_file(kind)
    new_path = config.OUTPUT_DIR / f"_이번실행_{kind}.csv"
    if src is None:
        print(f"  ✖ {DASH_DIR} 에 {kind} 취합본 없음")
        return 1
    if not new_path.exists():
        print(f"  ✖ 이번 실행 결과 없음({new_path.name}) — 먼저 클렌징을 돌릴 것")
        return 1
    raw, header = read_raw(src)
    new = pd.read_csv(new_path, low_memory=False)
    new = new.reindex(columns=raw.columns)
    raw, dup_raw = drop_duplicates(raw)          # 기준 RAW 에 원래 있던 완전 중복
    new, dup_new = drop_duplicates(new)          # 이번 원본 파일에 있던 완전 중복
    out, plan = merge(raw, new)
    out, media_n = media_pass(out)
    out = recompute_cpc(out)
    out = canonical_case(out, kind)
    checks = validate(raw, out, new, plan)
    week = week or date.today().isocalendar()[1]
    name = new_name(kind, week)
    report = write_report(kind, src, name, plan, checks, media_n, len(out), dup_raw, dup_new)
    print(f"  기준: {src.name}  ({len(raw) + dup_raw['행']:,}행) → 새 {len(out):,}행")
    for label, d in (("기준 RAW", dup_raw), ("이번 결과", dup_new)):
        if d["행"]:
            print(f"  · 중복 제거({label}): {d['행']:,}행 (숫자 있는 행 {d['숫자 있는 행']:,}, "
                  f"Spend {d['Spend']:,.0f}, Revenue {d['Revenue']:,.0f})  {d['법인별']}")
    for p in plan:
        print(f"  · {p['법인']:<6} {p['교체 시작']}~ 교체  RAW {p['RAW 행(교체)']:,}행 → 새 {p['새 행']:,}행"
              f"  Spend {p['Spend 전']:,.0f} → {p['Spend 후']:,.0f}")
    for r, k, v in checks:
        print(f"  {'✔' if r == 'OK' else ('⚠' if r == '경고' else '✖')} {k}: {v}")
    print(f"  검증 보고서: {report}")
    if any(r == "오류" for r, _, _ in checks):
        print("  ✖ 오류가 있어 취합본을 만들지 않음")
        return 1
    if not apply:
        print(f"  (미리보기 — 만들 파일: {name}.  --apply 로 생성)")
        return 0
    return create(src, name, header, out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["MX", "CE", "mx", "ce"])
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--week", type=int)
    ap.add_argument("--dash-dir", type=Path, help="시험용 취합 폴더(기본: 공유 드라이브 D2C)")
    a = ap.parse_args()
    raise SystemExit(run(a.kind, a.apply, a.week, a.dash_dir))
