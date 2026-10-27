"""원본 수집 — 국가별 weekly 폴더의 법인 엑셀에서 GMO 표준 시트를 뽑아 1_Input 형식 CSV 로 만든다.

    python tools/collect.py 43 CE            43주차 CE → 미리보기(표만)
    python tools/collect.py 43 CE --put      → 2_Output/_수집/W43_CE/ 에 CSV + 1_Input 에 넣기(inputs.put)
    python tools/collect.py 43 MX --put

폴더: <WEEKLY_ROOT>/<nn>주차/<MX|CE>/<법인 MMDD[꼬리표]>/ 안의 xlsx · xlsm · zip(안의 엑셀까지)
규칙
- 시트: 'Subsidiary' 와 'Media Platform' 이 있는 머리줄을 찾고(처음 20줄 안), 그런 시트 중 가장 최근 날짜까지 있는 것
  (SETK 는 'Global raw data_S.com' 이 비어 있고, 옛 'SETK 2026' 시트(6월까지 · MX 섞임)와
   이번 주 'CE Daily From w40' 시트가 같이 있다 → 최신 날짜 기준으로 고른다)
- 부분 기간 파일이어도 된다 — 마스터는 겹치는 날짜부터만 교체한다(updater)
- 같은 종류 파일(파일 이름에서 날짜 · 숫자를 뺀 것, 예: 'Monks - HQ - CE - AWIN')은 **최신 폴더 것**을 쓰고,
  최신 파일에 없는 날짜만 이전 폴더 파일에서 채운다. 종류가 다른 파일(AWIN + google_social_display)은 합친다
- 건너뜀: '복사본' · '~$' 파일, 구분이 다른 파일(CE 폴더 안의 '- MX -' 파일 등).
  'XXX'(안 쓴 것 — tools/mark_done.py) 는 그대로 읽는다 — 최신 파일에 덮여 어차피 안 쓰이고, 합치는 순서가 표시 전과 같아진다
- ' - 허겸수정(MMDD~)_a' 파일 = 클렌징 값을 넣은 복사본(mark_done) → 건너뛰고 옆의 원본을 읽는다(표시 전과 같은 결과)
- 열 정리: Div → BU, 'Products (optional)' 등은 파이프라인이 처리.
  날짜 열 이름이 없으면(SIEL 'Column2') 날짜가 든 열을 Date 로
- CE 인데 BU 열이 없으면(SETK) 캠페인명 BS~da / BS~vd 로 채운다
"""
from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

import fastexcel
import pandas as pd

from common import config, fmt_money
from src.campaign import extract_codes
from src.pipeline import sanitize_column_headers

WEEKLY_ROOT = Path(r"G:\공유 드라이브\Samsung Perform_1\문서-1\2026\09. GMO Report\2. 국가별 weekly")
STAGE = config.OUTPUT_DIR / "_수집"
SKIP_NAME = ("복사본", "~$")
MARK_DONE = re.compile(r"\s*-?\s*\(?허겸수정\s*\(?\s*\d{4}~?\s*\)?\s*_[aA]$")   # 처리 표시(mark_done)
MARK_SKIP = re.compile(r"(^|[\s_\-])XXX($|[\s_\-.])")                          # 안 쓴 것 표시


@dataclass
class Source:
    sub: str
    folder: str          # 'SEDA 1020'
    order: str           # 폴더 날짜 MMDD — 클수록 최신
    path: Path
    label: str           # 보여 줄 이름(zip 안이면 zip/파일)
    stream: str          # 같은 종류 판별 키
    sheet: str = ""      # 고른 시트
    origin: Path | None = None   # 폴더 안의 실제 파일(zip 안이면 그 zip)
    taken: tuple[str, str] | None = None   # merge_sub 가 이 파일에서 쓴 기간(첫 날, 마지막 날)
    taken_idx: list | None = None          # 쓴 행(read_source 표의 행 번호 = 머리줄 아래 몇 번째 줄)
    part_no: int = -1                      # 합친 표에서 몇 번째 조각인지


def _skip(name: str) -> bool:
    return any(k in name for k in SKIP_NAME) or bool(MARK_DONE.search(Path(name).stem))


def _stream(name: str) -> str:
    s = Path(name).stem.lower()
    s = re.sub(r"\d+([_.\-/]\d+)*", " ", s)        # 날짜 · 숫자 덩어리
    s = re.sub(r"\b(v\d+|week|w)\b", " ", s)
    return re.sub(r"[\s_\-]+", " ", s).strip()


def _kind_ok(name: str, kind: str) -> bool:
    other = "MX" if kind == "CE" else "CE"
    return not re.search(rf"(^|[^A-Za-z]){other}([^A-Za-z]|$)", name)


def list_sources(week: int, kind: str, tmp: Path, root: Path | None = None) -> list[Source]:
    base = (root or WEEKLY_ROOT / f"{week}주차") / kind
    if not base.exists():
        print(f"  · 폴더 없음: {base}")
        return []
    out: list[Source] = []
    for folder in sorted(p for p in base.iterdir() if p.is_dir()):
        m = re.match(r"\s*([A-Za-z]+)\s*(\d{3,4})?", folder.name)
        if not m:
            continue
        sub, order = m.group(1).upper(), (m.group(2) or "0000").zfill(4)
        seen = set()
        for f in sorted(folder.rglob("*")):
            if f.suffix.lower() in (".xlsx", ".xlsm") and not _skip(f.name):
                if _kind_ok(f.name, kind) and f.name not in seen:
                    seen.add(f.name)
                    out.append(Source(sub, folder.name, order, f, str(f.relative_to(folder)), _stream(f.name),
                                      origin=f))
            elif f.suffix.lower() == ".zip":
                dest = tmp / folder.name / f.stem
                with zipfile.ZipFile(f) as z:
                    for info in z.infolist():
                        n = Path(info.filename).name
                        if (n.lower().endswith((".xlsx", ".xlsm")) and not _skip(n)
                                and _kind_ok(n, kind) and n not in seen):
                            seen.add(n)
                            target = dest / n
                            target.parent.mkdir(parents=True, exist_ok=True)
                            target.write_bytes(z.read(info))
                            out.append(Source(sub, folder.name, order, target, f"{f.name}/{n}", _stream(n),
                                              origin=f))
    return out


def _header_row(df: pd.DataFrame) -> int | None:
    for i in range(min(20, len(df))):
        row = [str(v) for v in df.iloc[i].tolist() if v is not None]
        if any("Subsidiary" in v for v in row) and any("Media Platform" in v for v in row):
            return i
    return None


def read_source(src: Source) -> tuple[pd.DataFrame, str]:
    """가장 데이터가 많은 GMO 시트 → 정리된 DataFrame, 시트 이름."""
    book = fastexcel.read_excel(str(src.path))
    best, best_name = None, ""
    for name in book.sheet_names:
        try:
            head = book.load_sheet(name, header_row=None, n_rows=20).to_pandas()
        except Exception:
            try:                                  # 수식 오류 칸 — 아래 본문 읽기와 같은 이유
                head = book.load_sheet(name, header_row=None, n_rows=20, dtypes="string").to_pandas()
            except Exception:
                continue
        h = _header_row(head)
        if h is None:
            continue
        try:
            df = book.load_sheet(name, header_row=h).to_pandas()
        except Exception:
            # 수식 오류 칸(#NAME? 등)이 있으면 형 추론이 실패한다 — 글자로 다시 읽는다(SEPR W41 · W42, 2026-10-27)
            try:
                df = book.load_sheet(name, header_row=h, dtypes="string").to_pandas()
            except Exception:
                continue
        df = _clean(df)
        if df.empty:
            continue
        # 가장 최근 날짜까지 있는 시트(같으면 행이 많은 것) — 같은 파일에 옛 시트가 남아 있는 경우가 있다
        if best is None or (df["Date"].max(), len(df)) > (best["Date"].max(), len(best)):
            best, best_name = df, name
    return (best if best is not None else pd.DataFrame()), best_name


# 표준 템플릿('Global raw data_S.com')의 앞 9칸 순서
STD_ORDER = ["Subsidiary", "Sales Channel", "Partner", "Media Type 1", "Media Type 2", "Media Type 2_RAW (optional)",
             "Media Platform", "Media Platform_RAW (optional)", "Funding"]
# 법인 파일마다 다른 열 이름 → 표준 (SETK CE 일별 시트: 'Mon', Funding 칸 이름이 'Local')
HEADER_ALIASES = {"Div": "BU", "Mon": "Month", "Local": "Funding"}
_TAX = None


def _media_score(t1: pd.Series, t2: pd.Series, pf: pd.Series, fund: pd.Series) -> pd.Series:
    """행마다 미디어 · Funding 이 표준표에 얼마나 맞는지 0~3."""
    global _TAX
    from src.media_std import ALIASES, _key, load_taxonomy
    if _TAX is None:
        tax = load_taxonomy(config.MEDIA_TAXONOMY)
        k1, k2, kp = tax["t1"].map(_key), tax["t2"].map(_key), tax["pf"].map(_key)
        _TAX = (set(zip(k1, k2)), set(zip(k1, k2, kp)), {(a, b) for a, b, c in zip(k1, k2, kp) if c == "*"})
    ok2, ok3, aff = _TAX
    a = t1.map(_key)
    b = t2.map(_key).map(lambda v: ALIASES["Type 2"].get(v, v))
    c = pf.map(_key).map(lambda v: ALIASES["Platform"].get(v, v))
    s2 = pd.Series([(x, y) in ok2 for x, y in zip(a, b)], index=t1.index)
    s3 = pd.Series([(x, y, z) in ok3 or (x, y) in aff for x, y, z in zip(a, b, c)], index=t1.index)
    s4 = fund.map(_key).isin(["gmo", "local", "hq"])
    return s2.astype(int) + s3.astype(int) + s4.astype(int)


def _fix_shuffled_media(df: pd.DataFrame) -> pd.DataFrame:
    """헤더 순서가 표준과 다른 시트에서, 표준 순서로 붙여 넣은 행을 행 단위로 바로잡는다.
    (SEDA 월간 리포트: 헤더는 Type1 · Platform · Funding · Type2 … 인데 Affiliate 행만 표준 순서)"""
    cols = list(df.columns)
    need = STD_ORDER[3:]
    if len(cols) < len(STD_ORDER) or cols[:len(STD_ORDER)] == STD_ORDER or any(c not in cols for c in need):
        return df
    by_name = {c: df[c].fillna("").astype(str) for c in need}
    by_pos = {c: df.iloc[:, i].fillna("").astype(str) for i, c in enumerate(STD_ORDER) if i >= 3}
    s_name = _media_score(by_name["Media Type 1"], by_name["Media Type 2"], by_name["Media Platform"],
                          by_name["Funding"])
    s_pos = _media_score(by_pos["Media Type 1"], by_pos["Media Type 2"], by_pos["Media Platform"], by_pos["Funding"])
    flip = s_pos > s_name
    if flip.any():
        df = df.copy()
        for c in need[1:]:
            df.loc[flip, c] = by_pos[c][flip].replace("", None)
        print(f"        - 헤더 순서와 다른 행 {int(flip.sum()):,}개를 표준 순서로 읽음")
    return df


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df[[c for c in df.columns if c and not str(c).startswith(("__UNNAMED", "Unnamed"))]]
    df = sanitize_column_headers(df)
    df = df.rename(columns={k: v for k, v in HEADER_ALIASES.items() if k in df.columns and v not in df.columns})
    df = _fix_shuffled_media(df)
    if "Date" not in df.columns:
        for c in df.columns:
            s = df[c].dropna()
            if len(s) and pd.api.types.is_datetime64_any_dtype(df[c]):
                df = df.rename(columns={c: "Date"})
                break
    if "Subsidiary" not in df.columns or "Date" not in df.columns:
        return df.iloc[0:0]
    sub = df["Subsidiary"].astype(str).str.strip()
    dates = pd.to_datetime(df["Date"], errors="coerce")
    df = df[(sub != "") & (sub.str.lower() != "none") & (sub.str.lower() != "nan") & dates.notna()].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    return df


def _standardize_values(out: pd.DataFrame, kind: str, notes: list[str]) -> None:
    """표준 값으로 — Sales Channel · Partner, Funding, BU (담당자 확정 2026-10-20)."""
    for col in ("Sales Channel", "Partner"):          # S.com 하위 채널(Epp · app · web) = S.com (담당자 2026-10-20)
        if col in out.columns:
            v = out[col].astype(str).str.strip().str.lower()
            sub_ch = v.str.startswith(("epp", "app", "web")) | (v == "s.com")
            n_fix = int((sub_ch & (out[col] != "S.com")).sum())
            out.loc[sub_ch, col] = "S.com"
            if n_fix:
                notes.append(f"{col} Epp · app · web → S.com {n_fix:,}행")
    if "Funding" in out.columns:
        # Funding 은 Local · GMO 둘뿐(담당자 2026-10-21). HQ · D2C 예산 = GMO, SCIC(법인 이름) = Local
        v = out["Funding"].astype(str).str.strip().str.upper()
        for src, dst in (("HQ", "GMO"), ("D2C", "GMO"), ("SCIC", "Local"), ("GMO", "GMO"), ("LOCAL", "Local")):
            hit = (v == src) & (out["Funding"] != dst)
            out.loc[hit, "Funding"] = dst
            if hit.any() and src not in ("GMO", "LOCAL"):
                notes.append(f"Funding {src} → {dst} {int(hit.sum()):,}행")
        odd = out.loc[~out["Funding"].isin(["Local", "GMO"]), "Funding"].astype(str).str.strip()
        if len(odd):
            notes.append("⚠ Funding Local · GMO 밖: " + " · ".join(f"{k or '빈 값'} {n:,}행"
                                                              for k, n in odd.value_counts().items()))
    if kind == "CE" and config.COL_BU in out.columns:
        out[config.COL_BU] = out[config.COL_BU].replace({"HA": "DA", "ha": "DA"})   # HA(Home Appliance) = DA


def merge_sub(items: list[tuple[Source, pd.DataFrame]], kind: str) -> tuple[pd.DataFrame, list[str]]:
    """한 법인의 파일들 → 하나. 같은 종류는 최신 폴더 우선, 빠진 날짜만 이전 것."""
    notes, parts = [], []
    by_stream: dict[str, list] = {}
    for src, df in items:
        by_stream.setdefault(src.stream, []).append((src, df))
    for stream, lst in by_stream.items():
        lst.sort(key=lambda x: x[0].order, reverse=True)
        covered: set[str] = set()
        for src, df in lst:
            take = df[~df["Date"].isin(covered)]
            if len(take):
                parts.append(take)
                src.taken = (take["Date"].min(), take["Date"].max())
                src.taken_idx, src.part_no = take.index.tolist(), len(parts) - 1
                notes.append(f"{src.folder}/{src.label} [{src.sheet}]: {len(take):,}행 "
                             f"{take['Date'].min()}~{take['Date'].max()}")
            elif len(df):
                notes.append(f"{src.folder}/{src.label}: 최신 파일에 같은 날짜가 있어 안 씀")
            covered |= set(df["Date"])
    if not parts:
        return pd.DataFrame(), notes
    out = pd.concat(parts, ignore_index=True, sort=False)
    _standardize_values(out, kind, notes)
    if kind == "CE" and config.COL_BU not in out.columns and "Campaign Name" in out.columns:
        bs = extract_codes(out["Campaign Name"])["BS"].str.upper()
        out[config.COL_BU] = bs.where(bs.isin(["DA", "VD"]), "")
        notes.append("BU 열 없음 → 캠페인명 BS~ 코드로 채움")
    return out, notes


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("week", type=int)
    ap.add_argument("kind", choices=["CE", "MX", "ce", "mx"])
    ap.add_argument("--put", action="store_true", help="1_Input 에 넣기")
    ap.add_argument("--root", type=Path, help="주차 폴더를 직접 지정(그 안의 MX · CE 폴더를 읽음)")
    a = ap.parse_args(argv)
    kind = a.kind.upper()
    stage = STAGE / f"W{a.week}_{kind}"
    with tempfile.TemporaryDirectory() as td:
        sources = list_sources(a.week, kind, Path(td), a.root)
        if not sources:
            return 0
        by_sub: dict[str, list] = {}
        print(f"■ 수집 — {a.week}주차 {kind} ({len(sources)}개 파일)")
        for src in sources:
            df, src.sheet = read_source(src)
            if df.empty:
                print(f"  · {src.folder}/{src.label}: GMO 시트 없음 — 건너뜀")
                continue
            by_sub.setdefault(src.sub, []).append((src, df))
        stage.mkdir(parents=True, exist_ok=True)
        written = []
        print(f"\n  {'법인':<6}{'행':>8}  {'기간':<23}{'Spend(USD)':>14}")
        for sub, items in sorted(by_sub.items()):
            df, notes = merge_sub(items, kind)
            if df.empty:
                continue
            sp = next((c for c in df.columns if c.startswith("Media Spend (USD)")), None)
            spend = pd.to_numeric(df[sp], errors="coerce").fillna(0).sum() if sp else 0
            print(f"  {sub:<6}{len(df):>8,}  {df['Date'].min()}~{df['Date'].max():<12}{fmt_money(spend):>14}")
            for n in notes:
                print(f"        - {n}")
            name = f"{sub}.csv" if kind == "MX" else f"{sub}_CE.csv"
            path = stage / name
            df.to_csv(path, index=False, encoding="utf-8-sig")
            written.append(path)
    print(f"\n  CSV: {stage}")
    if a.put and written:
        from inputs import put, status
        print("\n■ 1_Input 에 넣기")
        put(written)
        status()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
