"""tools 공통 — 입력 CSV 판별 · 슬롯 이름 · 잠금 확인.

입력 CSV 하나 = 슬롯 하나 = (법인, MX/CE). BU 열이 있으면 CE, 없으면 MX.
run.py 의 분기(config.COL_BU 존재 여부)와 같은 기준이다.
단, BU 값이 MX(또는 빈 값)뿐이면 MX — 1_Input 에 넣을 때 BU 열을 빼서 run.py 도 MX 로 돈다
(W44 SEDA MX 원본에 BU=MX 열이 있어 CE 로 클렌징된 일, 2026-10-27).
"""
from __future__ import annotations

import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src import config  # noqa: E402
from src.pipeline import sanitize_column_headers  # noqa: E402

REQUIRED = ["Subsidiary", "Media Type 1", "Media Type 2", "Media Platform",
            "Product Category", "Product Series", "Date"]
ARCHIVE_DIR = config.INPUT_DIR / "_보관"
REPORT_DIR = config.OUTPUT_DIR / "Reports"


@dataclass
class InputInfo:
    path: Path
    ok: bool = False
    sub: str = ""
    kind: str = ""              # MX / CE
    rows: int = 0
    date_min: str = ""
    date_max: str = ""
    spend: float = 0.0
    revenue: float = 0.0
    warnings: list[str] = field(default_factory=list)

    @property
    def slot(self) -> str:
        return f"{self.sub}_{self.kind}" if self.ok else ""


def _read(path: Path) -> pd.DataFrame | None:
    for enc in ("utf-8-sig", "cp949"):
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False, dtype=str)
        except UnicodeDecodeError:
            continue
        except Exception:
            return None
    return None


def _num(s: pd.Series) -> float:
    return float(pd.to_numeric(s.astype(str).str.replace(",", "").str.strip(),
                               errors="coerce").fillna(0).sum())


def _first_col(df: pd.DataFrame, prefix: str) -> str | None:
    return next((c for c in df.columns if c.startswith(prefix)), None)


def mx_only_bu(df: pd.DataFrame) -> bool:
    """BU 열이 있는데 값이 MX · 빈 값뿐 = MX 파일(헤더는 정리된 상태로 받는다)."""
    if config.COL_BU not in df.columns:
        return False
    v = df[config.COL_BU].fillna("").astype(str).str.strip().str.upper()
    return bool(v.isin(["MX", ""]).all())


def inspect_csv(path: Path) -> InputInfo:
    """CSV 내용으로 GMO 입력인지, 어느 법인 · MX/CE 인지, 기간 · 합계를 본다."""
    info = InputInfo(path=path)
    df = _read(path)
    if df is None:
        info.warnings.append("읽기 실패(Excel 에서 열려 있거나 CSV 가 아님)")
        return info
    df = sanitize_column_headers(df)
    df = df.rename(columns={"Product Category2": "Product Category"})
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        info.warnings.append("GMO 형식 아님 — 없는 열: " + ", ".join(missing))
        return info

    # 날짜가 유효한 행만 본다(파이프라인도 그 행만 남긴다 — 끝의 '요약' 행 등은 빠짐)
    from src.utils import process_and_filter_dates
    level = logging.getLogger().level
    logging.getLogger().setLevel(logging.WARNING)
    try:
        dated = process_and_filter_dates(df[["Subsidiary", "Date"]].copy(), "Date")
    finally:
        logging.getLogger().setLevel(level)
    info.rows = len(dated)
    info.kind = "CE" if config.COL_BU in df.columns and not mx_only_bu(df) else "MX"
    if info.kind == "CE":
        n_mx = int((df[config.COL_BU].fillna("").astype(str).str.strip().str.upper() == "MX").sum())
        if n_mx:
            info.warnings.append(f"CE 파일에 BU=MX {n_mx:,}행 섞임")
    elif config.COL_BU in df.columns:
        info.warnings.append("BU 열이 MX 뿐 → MX 로 보고, 넣을 때 BU 열을 뺌")
    if dated.empty:
        info.warnings.append(f"{config.MIN_DATE} 이후 날짜 없음")
        return info

    subs = dated["Subsidiary"].fillna("").str.strip().replace("", "(빈 값)")
    counts = subs.value_counts()
    info.sub = str(counts.index[0]).upper()
    if len(counts) > 1:
        info.warnings.append("법인 여러 개: " + ", ".join(
            f"{k} {v:,}행" for k, v in counts.items()))
    info.ok = True
    info.date_min, info.date_max = dated["Date"].min(), dated["Date"].max()

    spend_col = _first_col(df, "Media Spend (USD)") or _first_col(df, "Media Spend")
    rev_col = _first_col(df, "Revenue (USD)") or _first_col(df, "Revenue")
    if spend_col:
        info.spend = _num(df[spend_col])
    if rev_col:
        info.revenue = _num(df[rev_col])
    return info


def master_max_dates() -> dict[str, str]:
    """마스터 CSV 에 들어 있는 슬롯별 마지막 날짜 — {'SEAU_MX': '2026-09-30', ...}."""
    out: dict[str, str] = {}
    for kind, path in config.MASTER_FILES.items():
        if not path.exists():
            continue
        m = pd.read_csv(path, usecols=["Subsidiary", "Date"], dtype=str)
        for sub, d in m.groupby(m["Subsidiary"].fillna("").str.upper())["Date"].max().items():
            out[f"{sub}_{kind}"] = d
    return out


def current_slots() -> dict[str, list[InputInfo]]:
    """1_Input 의 CSV 를 슬롯별로 묶는다(같은 슬롯에 두 개 이상이면 중복)."""
    slots: dict[str, list[InputInfo]] = {}
    for p in sorted(config.INPUT_DIR.glob("*.csv")):
        info = inspect_csv(p)
        slots.setdefault(info.slot or f"?{p.name}", []).append(info)
    return slots


def is_locked(path: Path) -> bool:
    """Excel 이 열고 있으면 True — ~$ 잠금 파일 또는 쓰기 거부."""
    if (path.parent / f"~${path.name}").exists():
        return True
    if not path.exists():
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True


def fmt_money(v: float) -> str:
    return f"{v:,.0f}"
