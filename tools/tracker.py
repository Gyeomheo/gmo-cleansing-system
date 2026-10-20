"""Collection status Tracker — 취합본에 메일 양식과 같은 'Tracker' 시트를 넣는다. 판정은 status 피벗 기준.

    python tools/tracker.py CE --week 43          최신 CE 취합본에 Tracker 넣기 · 갱신
    (consolidate.py --apply 가 새 취합본을 만들 때 자동으로 부른다)

    Week 칸      =IFERROR(IF(GETPIVOTDATA(<값 필드>, status!<피벗>, "Subsidiary", 법인, "Week", 주차) > 0,
                 "Yes", "-"), "-")
                 보고 주차 W 이면 W-2 · W-1 두 주 (메일 양식과 같다)
    Uncollected  두 주 모두 '-' 이면 True
    Note         고정 사유(tracker_markets.csv)가 있으면 그것, 없고 '-' 가 있으면 'Missing submission'
    추적 = N     (tracker_markets.csv) 제출 안 하는 법인 — 회색 줄, 주차 · Uncollected 비움, Note 는 고정 사유만
법인 · Strategic · Agency · 고정 사유의 기본값 = 3_Config/tracker_markets.csv.
시트에는 메일에 붙일 표만 둔다 — 사유를 바꿀 땐 tracker_markets.csv 를 고친다.
"""
from __future__ import annotations

import argparse
import html
from pathlib import Path

import pandas as pd

from common import config

MARKETS = config.CONFIG_DIR / "tracker_markets.csv"
SHEET = "Tracker"


def load_markets(kind: str) -> pd.DataFrame:
    df = pd.read_csv(MARKETS, dtype=str, encoding="utf-8-sig").fillna("")
    return df[df["구분"].str.upper() == kind].reset_index(drop=True)


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
    first, last = 5, 4 + len(m)
    if len(m) > 11:                                      # 양식보다 법인이 많으면 마지막 줄 서식을 아래로
        ws.Range("B15:H15").Copy(Destination=ws.Range(f"B16:H{last}"))
    for i in range(len(m)):
        row = first + i
        ws.Cells(row, 2).Value = m.iloc[i]["Market"]
        ws.Cells(row, 3).Value = m.iloc[i]["Strategic"]
        ws.Cells(row, 4).Value = m.iloc[i]["Agency"]
        reason = str(m.iloc[i]["고정 사유"]).strip()
        if str(m.iloc[i].get("추적", "Y")).strip().upper() == "N":
            # 제출 안 하는 법인(회색 · Stop submission 등): 주차 칸은 비우고 사유만, 줄은 양식의 회색
            ws.Range(f"E{row}:G{row}").ClearContents()
            ws.Cells(row, 8).Value = reason
            ws.Range(f"B{row}:H{row}").Interior.Color = GRAY
            continue
        for k, w in enumerate(weeks):
            ws.Cells(row, 5 + k).Formula = (f'=IFERROR(IF(GETPIVOTDATA("{value_field}",{anchor},"Subsidiary",'
                                            f'TRIM(SUBSTITUTE($B{row},CHAR(160),"")),"Week",{w})>0,"Yes","-"),"-")')
        ws.Cells(row, 7).Formula = f'=IF(AND(E{row}="-",F{row}="-"),"True","")'
        if reason:                                       # 고정 사유(tracker_markets.csv)가 있으면 그것
            ws.Cells(row, 8).Value = reason
        else:
            ws.Cells(row, 8).Formula = f'=IF(OR(E{row}="-",F{row}="-"),"{NB}Missing submission","")'
    # 값에 따라 바뀌는 색은 양식 색 그대로 칸에 직접 칠한다(조건부 서식은 Gmail 붙여넣기에서 빠질 수 있다)
    wb.Application.Calculate()
    for row in range(first, last + 1):
        for col in (5, 6):
            c = ws.Cells(row, col)
            c.Font.Color = GREEN if str(c.Value).replace(NB, "").strip() == "Yes" else 0x000000
    ws.Range(f"G{first}:G{last}").Font.Color = 0x000000
    ws.Range(f"H{first}:H{last}").Font.Color = BLUE
    wb.Application.Calculate()                       # 시트에는 메일 표만(안내 · 고정 사유 칸 없음 — 담당자 2026-10-20)


def read_back(path: Path, kind: str) -> pd.DataFrame:
    """저장된 Tracker 값(수식 결과)을 읽는다 — 확인용."""
    import fastexcel
    t = fastexcel.read_excel(str(path)).load_sheet(SHEET, header_row=3).to_pandas()
    t.columns = [str(c).replace(" ", "").strip() for c in t.columns]
    return t


def bgr_to_hex(c) -> str:
    c = int(c)
    return f"#{c & 0xFF:02X}{(c >> 8) & 0xFF:02X}{(c >> 16) & 0xFF:02X}"


def table_html(ws, rng: str) -> str:
    """Excel 범위 → 인라인 서식 HTML 표 (병합 · 바탕 · 글자색 · 굵기 · 정렬 · 테두리 · 열 너비)."""
    r = ws.Range(rng)
    nrow, ncol = r.Rows.Count, r.Columns.Count
    widths = [round(r.Columns(j).Width * 1.33) for j in range(1, ncol + 1)]
    skip = set()
    out = ['<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-family:Aptos,Calibri,Arial;'
           'font-size:11pt">']
    for i in range(1, nrow + 1):
        out.append("<tr>")
        for j in range(1, ncol + 1):
            if (i, j) in skip:
                continue
            c = r.Cells(i, j)
            span = ""
            if c.MergeCells:
                ma = c.MergeArea
                if ma.Cells(1, 1).Address != c.Address:
                    continue
                rs, cs = ma.Rows.Count, ma.Columns.Count
                for a in range(rs):
                    for b in range(cs):
                        skip.add((i + a, j + b))
                span = f' rowspan="{rs}" colspan="{cs}"'
            text = c.Text or ""
            style = [f"width:{widths[j - 1]}px", "padding:2px 6px", "white-space:nowrap"]
            if c.Interior.Pattern != -4142:
                style.append(f"background:{bgr_to_hex(c.Interior.Color)}")
            style.append(f"color:{bgr_to_hex(c.Font.Color)}")
            if c.Font.Bold:
                style.append("font-weight:bold")
            style.append({-4108: "text-align:center", -4152: "text-align:right"}.get(c.HorizontalAlignment,
                                                                                      "text-align:left"))
            for k, side in ((7, "left"), (8, "top"), (9, "bottom"), (10, "right")):   # 변마다 그대로
                b = c.Borders(k)
                if b.LineStyle not in (-4142, None):
                    style.append(f"border-{side}:1px solid {bgr_to_hex(b.Color)}")
            out.append(f'<td{span} style="{";".join(style)}">{html.escape(text) or "&nbsp;"}</td>')
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)


def fragment_path(kind: str, week: int) -> Path:
    return config.OUTPUT_DIR / "Reports" / f"Tracker_{kind}_W{week}.html"


def run(kind: str, week: int, path: Path | None = None, prev: Path | None = None) -> int:
    import pythoncom
    import win32com.client

    from consolidate import latest_file
    path = path or latest_file(kind)
    reasons: dict[str, str] = {}
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        write_tracker(wb, kind, week, reasons)
        ws = wb.Worksheets(SHEET)
        last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
        # 메일용 표 HTML 을 여기서 같이 저장 — 메일 단계는 Excel 을 띄우지 않고 이 파일만 읽는다(속도)
        frag = fragment_path(kind, week)
        frag.parent.mkdir(parents=True, exist_ok=True)
        frag.write_text(table_html(ws, f"B3:H{last}"), encoding="utf-8")
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
