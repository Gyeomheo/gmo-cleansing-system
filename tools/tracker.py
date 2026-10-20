"""Collection status Tracker — 취합본에 메일 양식과 같은 'Tracker' 시트를 넣는다. 판정은 status 피벗 기준.

    python tools/tracker.py CE --week 43          최신 CE 취합본에 Tracker 넣기 · 갱신
    (consolidate.py --apply 가 새 취합본을 만들 때 자동으로 부른다)

    Week 칸      =IFERROR(IF(GETPIVOTDATA(<값 필드>, status!<피벗>, "Subsidiary", 법인, "Week", 주차) > 0,
                 "Yes", "-"), "-")
                 보고 주차 W 이면 W-2 · W-1 두 주 (메일 양식과 같다)
    Uncollected  두 주 모두 '-' 이면 True
    Note         '고정 사유' 칸이 있으면 그것, 없고 '-' 가 있으면 'Missing submission'
    추적 = N     (tracker_markets.csv) 제출 안 하는 법인 — 회색 줄, 주차 · Uncollected 비움, Note 는 고정 사유만
법인 · Strategic · Agency · 고정 사유의 기본값 = 3_Config/tracker_markets.csv.
이전 취합본 Tracker 의 '고정 사유' 칸에 적은 값은 다음 취합본으로 이어진다(파일에서 고쳐도 됨).
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from common import config

MARKETS = config.CONFIG_DIR / "tracker_markets.csv"
SHEET = "Tracker"


def load_markets(kind: str) -> pd.DataFrame:
    df = pd.read_csv(MARKETS, dtype=str, encoding="utf-8-sig").fillna("")
    return df[df["구분"].str.upper() == kind].reset_index(drop=True)


def carried_reasons(prev: Path | None) -> dict[str, str]:
    """이전 취합본 Tracker 의 고정 사유 {법인: 사유}."""
    if not prev or not prev.exists():
        return {}
    import fastexcel
    try:
        book = fastexcel.read_excel(str(prev))
        if SHEET not in book.sheet_names:
            return {}
        t = book.load_sheet(SHEET, header_row=3).to_pandas()
    except Exception:
        return {}
    t.columns = [str(c).replace(" ", "").strip() for c in t.columns]
    if "Market" not in t.columns or "고정 사유" not in t.columns:
        return {}
    return {str(m).strip(): str(r).strip() for m, r in zip(t["Market"], t["고정 사유"])
            if str(m).strip() and str(r).strip() not in ("", "None", "nan")}


TEMPLATE_DIR = Path(r"G:\공유 드라이브\Samsung Perform_1\문서-1\2026\09. GMO Report\2. 국가별 weekly")
TEMPLATE_SHEET = "3. 메일 작성"
TEMPLATE_RANGE = {"MX": "C38:I51", "CE": "C54:I67"}     # 'MX'/'CE' 줄 · Collection status · 머리줄 · 법인 11줄
NB = " "                                           # 메일 양식 칸 앞의 공백(Outlook 붙여넣기 그대로)
GREEN, BLUE = 0x40C06F, 0xFF0000                     # 양식의 Yes 초록(#6FC040) · Note 파랑(#0000FF), BGR
GRAY = 0x808080                                      # 양식의 회색 줄(제출 안 하는 법인)


def template_file() -> Path:
    files = sorted(TEMPLATE_DIR.glob("D2C Weekly Report 취합_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not files:
        raise FileNotFoundError(f"메일 양식 파일 없음: {TEMPLATE_DIR}\D2C Weekly Report 취합_*.xlsx")
    return files[-1]


def _copy_template(wb, ws, kind: str) -> None:
    """메일 양식 표를 서식째 Tracker!B2 로 복사."""
    xl = wb.Application
    tpl = xl.Workbooks.Open(str(template_file()), ReadOnly=True, UpdateLinks=0)
    try:
        tpl.Worksheets(TEMPLATE_SHEET).Range(TEMPLATE_RANGE[kind]).Copy(Destination=ws.Range("B2"))
        src = tpl.Worksheets(TEMPLATE_SHEET).Range(TEMPLATE_RANGE[kind])
        for j in range(1, src.Columns.Count + 1):                # 열 너비도 같게
            ws.Columns(1 + j).ColumnWidth = src.Columns(j).ColumnWidth
    finally:
        xl.CutCopyMode = False
        tpl.Close(SaveChanges=False)


def _tidy_head(ws) -> None:
    """양식의 빨간 'MX'/'CE' 줄은 비우고, 'Collection status' 줄은 그 칸(E3:F3)에만 테두리(담당자 2026-10-20)."""
    ws.Range("B2:H2").Clear()
    for rng in ("B2:H2", "B3:D3", "G3:H3"):
        for edge in (7, 8, 10, 11, 12):                 # 왼 · 위 · 오른 · 안쪽 세로 · 안쪽 가로
            ws.Range(rng).Borders(edge).LineStyle = -4142
    for edge in (7, 8, 10):
        ws.Range("E3:F3").Borders(edge).LineStyle = 1
        ws.Range("E3:F3").Borders(edge).Weight = 1


def write_tracker(wb, kind: str, week: int, reasons: dict[str, str]) -> None:
    """열린 통합문서(COM)에 Tracker 시트를 메일 양식 그대로 만들고, 판정은 status 피벗 수식으로."""
    st = wb.Worksheets("status")
    pt = st.PivotTables().Item(1)
    value_field = pt.DataFields.Item(1).SourceName
    anchor = f"status!{pt.TableRange1.Cells(1, 1).Address}"
    weeks = [week - 2, week - 1]                         # 최근 2주

    if SHEET in [s.Name for s in wb.Worksheets]:
        ws = wb.Worksheets(SHEET)
        ws.Cells.Clear()
    else:
        ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
        ws.Name = SHEET
    _copy_template(wb, ws, kind)              # B2:H15 = 양식 (B2 'MX'/'CE', E3 Collection status, 4행 머리줄)
    _tidy_head(ws)
    ws.Range("E4").Value = f"{NB}Week {weeks[0]}"
    ws.Range("F4").Value = f"{NB}Week {weeks[1]}"
    ws.Range("E4:F4").Font.Bold = ws.Range("B4").Font.Bold

    m = load_markets(kind)
    m["고정 사유"] = [reasons.get(k, v) for k, v in zip(m["Market"], m["고정 사유"])]
    first, last = 5, 4 + len(m)
    if len(m) > 11:                                      # 양식보다 법인이 많으면 마지막 줄 서식을 아래로
        ws.Range("B15:H15").Copy(Destination=ws.Range(f"B16:H{last}"))
    for i in range(len(m)):
        row = first + i
        ws.Cells(row, 2).Value = m.iloc[i]["Market"]
        ws.Cells(row, 3).Value = m.iloc[i]["Strategic"]
        ws.Cells(row, 4).Value = m.iloc[i]["Agency"]
        ws.Cells(row, 10).Value = m.iloc[i]["고정 사유"]
        if str(m.iloc[i].get("추적", "Y")).strip().upper() == "N":
            # 제출 안 하는 법인(회색 · Stop submission 등): 주차 칸은 비우고 사유만, 줄은 양식의 회색
            ws.Range(f"E{row}:G{row}").ClearContents()
            ws.Cells(row, 8).Formula = f'=IF($J{row}<>"",$J{row},"")'
            ws.Range(f"B{row}:H{row}").Interior.Color = GRAY
            continue
        for k, w in enumerate(weeks):
            ws.Cells(row, 5 + k).Formula = (f'=IFERROR(IF(GETPIVOTDATA("{value_field}",{anchor},"Subsidiary",'
                                            f'TRIM(SUBSTITUTE($B{row},CHAR(160),"")),"Week",{w})>0,"Yes","-"),"-")')
        ws.Cells(row, 7).Formula = f'=IF(AND(E{row}="-",F{row}="-"),"True","")'
        ws.Cells(row, 8).Formula = (f'=IF($J{row}<>"",$J{row},'
                                    f'IF(OR(E{row}="-",F{row}="-"),"{NB}Missing submission",""))')
    # 값에 따라 바뀌는 색은 양식 색 그대로 칸에 직접 칠한다(조건부 서식은 Gmail 붙여넣기에서 빠질 수 있다)
    wb.Application.Calculate()
    for row in range(first, last + 1):
        for col in (5, 6):
            c = ws.Cells(row, col)
            c.Font.Color = GREEN if str(c.Value).replace(NB, "").strip() == "Yes" else 0x000000
    ws.Range(f"G{first}:G{last}").Font.Color = 0x000000
    ws.Range(f"H{first}:H{last}").Font.Color = BLUE
    ws.Range("J4").Value = "고정 사유"
    ws.Range("K4").Value = ("< 미취합 사유 · 특이사항(메일 Note 에 그대로, 다음 주 파일로 이어짐). "
                            "비우면 '-' 일 때 Missing submission")
    ws.Range("J4:K4").Font.Color = 0x808080
    ws.Range(f"J{first}:J{last}").Interior.Color = 0xF2F2F2
    ws.Columns("J").ColumnWidth = 30
    ws.Range("B1").Value = f"메일용 — B3:H{last} 복사 (W{week} 보고, 판정 = status 피벗, 최근 2주)"
    ws.Range("B1").Font.Color = 0x808080
    wb.Application.Calculate()


def read_back(path: Path, kind: str) -> pd.DataFrame:
    """저장된 Tracker 값(수식 결과)을 읽는다 — 확인용."""
    import fastexcel
    t = fastexcel.read_excel(str(path)).load_sheet(SHEET, header_row=3).to_pandas()
    t.columns = [str(c).replace(" ", "").strip() for c in t.columns]
    return t


def run(kind: str, week: int, path: Path | None = None, prev: Path | None = None) -> int:
    import pythoncom
    import win32com.client

    from consolidate import latest_file
    path = path or latest_file(kind)
    reasons = carried_reasons(prev) if prev else carried_reasons(path)
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        write_tracker(wb, kind, week, reasons)
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()
    t = read_back(path, kind)
    cols = [c for c in t.columns if c in ("Market", "Strategic") or str(c).startswith("Week") or c in
            ("Uncollected 2+ Weeks", "Note")]
    print(f"■ Tracker — {path.name}")
    print(t[cols].fillna("").to_string(index=False))
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["MX", "CE", "mx", "ce"])
    ap.add_argument("--week", type=int, required=True, help="보고 주차(W). Tracker 는 W-2 · W-1 을 보여 준다")
    ap.add_argument("--file", type=Path, help="취합본 경로(기본: 최신)")
    a = ap.parse_args()
    raise SystemExit(run(a.kind.upper(), a.week, a.file))
