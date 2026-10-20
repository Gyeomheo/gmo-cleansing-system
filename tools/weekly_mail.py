"""주간 메일 · 리마인드 메일 — Gmail 에 붙여 넣을 HTML 을 만든다(보내지는 않는다).

    python tools/weekly_mail.py --week 43            → 2_Output/Reports/메일_W43.html (브라우저로 열림)

주간 메일(메일 양식 시트 '3. 메일 작성' 문구 그대로)
  제목  [<이번 주 구분>] 2026 W<nn> GMO Weekly Report   (담당자 2026-10-20)
  본문  이번 주 구분(MX/CE, 최신 취합본이 W<nn> 인 쪽) · 첨부 파일명 · Collection status 표 MX · CE
        표는 각 최신 취합본의 Tracker 시트를 색 · 굵기 · 테두리까지 인라인 서식으로 옮긴다(Gmail 붙여넣기용)
리마인드 메일
  이번 주 구분 Tracker 에서 Note 가 Missing submission 인 법인마다 한 통.
  다른 구분도 미제출이면 구분 없이(담당자 양식 그대로), 한쪽만이면 'MX data' / 'CE data'.
  주차 범위 = RAW 에서 그 법인의 마지막 제출 주 다음 ~ 보고 주차-1.
"""
from __future__ import annotations

import argparse
import html
import os
from pathlib import Path

import fastexcel
import pandas as pd

from common import REPORT_DIR
from consolidate import NAME_RE, latest_file
from tracker import TEMPLATE_SHEET, template_file

DRAFT_DROP = Path(r"G:\내 드라이브\GMO_메일초안")                  # GAS(tools/gas_mail) 가 읽는 폴더

NB = "\xa0"


def contacts() -> dict[str, str]:
    """메일 양식 시트의 PIC(C8) · CC(C9:C12) — Excel 없이 읽는다."""
    try:
        d = fastexcel.read_excel(str(template_file())).load_sheet(TEMPLATE_SHEET, header_row=None,
                                                                  n_rows=12).to_pandas()
        cell = lambda r: str(d.iat[r - 1, 2]) if d.shape[0] >= r and d.iat[r - 1, 2] is not None else ""  # noqa: E731
        return {"to": cell(8), "cc": " ".join(cell(r) for r in (9, 10, 11, 12))}
    except Exception:
        return {"to": "", "cc": ""}


def tracker_tables(week: int) -> dict:
    """그 주 구분(홀수 CE · 짝수 MX) 하나만: {kind: (주차, 파일, 표 html)}, '_contacts'.
    표 HTML 은 Tracker 를 만들 때 저장해 둔 파일을 읽는다 — 없을 때만 Excel 로 Tracker 를 다시 만든다."""
    import tracker
    kind = "CE" if week % 2 else "MX"
    res = {"_contacts": contacts()}
    f = latest_file(kind)
    if f and int(NAME_RE.match(f.name).group(2)) == week:
        frag = tracker.fragment_path(kind, week)
        if not frag.exists():
            tracker.run(kind, week, path=f)
        res[kind] = (week, f, frag.read_text(encoding="utf-8"))
    return res


def tracker_df(path: Path) -> pd.DataFrame:
    t = fastexcel.read_excel(str(path)).load_sheet("Tracker", header_row=3).to_pandas()
    t.columns = [str(c).replace(NB, "").strip() for c in t.columns]
    return t


def last_week(path: Path) -> dict[str, int]:
    """법인별 '다 낸' 마지막 주차 — 안 준 날짜부터 미제출로 본다(담당자 2026-10-20).

    일별 데이터: 마지막 날의 다음 날이 속한 주부터 미제출 → 그 앞 주까지가 '다 낸 주'
                 (SEI CE 가 9/30(수)까지면 W40 은 10/1~ 이 비어 있으므로 W39 까지 낸 것)
    주간 합계(최근 8주 동안 날짜가 주에 하루뿐): 마지막 행의 주를 다 낸 것으로 본다(월요일 날짜 등)
    """
    raw = fastexcel.read_excel(str(path)).load_sheet("RAW", use_columns=["Subsidiary", "Date"]).to_pandas()
    raw["Date"] = pd.to_datetime(raw["Date"], errors="coerce")
    out = {}
    for sub, g in raw.dropna(subset=["Date"]).groupby("Subsidiary"):
        last = g["Date"].max()
        recent = g[g["Date"] > last - pd.Timedelta(weeks=8)]["Date"]
        days_per_week = recent.groupby(recent.dt.isocalendar().week).nunique().mean()
        if days_per_week <= 1.5:                           # 주간 합계 행
            out[sub] = int(last.isocalendar().week)
        else:
            out[sub] = int((last + pd.Timedelta(days=1)).isocalendar().week) - 1
    return out


def missing_ranges(kind: str, path: Path, report_week: int) -> dict[str, tuple[int, int]]:
    """Tracker 에서 최근 2주가 모두 '-' 이고 Missing submission 인 법인 → (시작 주, 끝 주).
    2주 중 하나라도 Yes 면 리마인드하지 않는다(담당자 2026-10-20)."""
    t = tracker_df(path)
    lw = last_week(path)
    week_cols = [c for c in t.columns if str(c).startswith("Week")]
    out = {}
    for _, r in t.iterrows():
        note = str(r.get("Note", "")).replace(NB, "").strip()
        any_yes = any(str(r[c]).replace(NB, "").strip() == "Yes" for c in week_cols)
        if note == "Missing submission" and not any_yes:
            sub = str(r["Market"]).strip()
            out[sub] = (lw.get(sub, 0) + 1, report_week - 1)
    return out


def weeks_text(a: int, b: int) -> str:
    return f"Week {a}" if a == b else f"Week {a}~{b}"


def reminder_body(sub: str, ranges: dict[str, tuple[int, int]]) -> str:
    """담당자 양식(2026-10-20): 구분별 Status 줄. ranges = {'MX': (a, b), 'CE': (a, b)} 중 미제출인 것."""
    kinds = ", ".join(ranges)                                   # 'MX, CE' / 'CE'
    lo, hi = min(v[0] for v in ranges.values()), max(v[1] for v in ranges.values())
    status = chr(10).join(f"{k} : [{weeks_text(*v)}] currently pending" for k, v in ranges.items())
    return f"""Dear {sub} Team,
I would like to kindly remind you that the {kinds} data for {weeks_text(lo, hi)} has not been received.
Please ensure that the reports are submitted according to the following requirements:

Submission Requirement:
Frequency: Weekly
Data Range: Up to the preceding Sunday
Status:
{status}

Please share the missing files at your earliest convenience to avoid any further gaps.
Best regards,
Gyeom"""


def build(week: int) -> Path:
    tabs = tracker_tables(week)
    kinds = [k for k in ("MX", "CE") if k in tabs]
    this = next((k for k in kinds if tabs[k][0] == week), None)
    parts = ['<html><head><meta charset="utf-8"><title>메일 W%d</title>'
             '<style>body{font-family:Aptos,Calibri,Arial;font-size:11pt;max-width:1100px;margin:24px auto;'
             'padding:0 16px}.box{margin:12px 0 28px}.meta{color:#666;font-size:10pt}'
             'h2{font-size:13pt;margin-top:32px}</style></head><body>' % week]
    # ---- 주간 메일
    label = this or "MX/CE"
    subject = f"[{label}] 2026 W{week} GMO Weekly Report"
    c = tabs.get("_contacts", {})
    parts.append("<h2>① 주간 메일 — 이 스레드에 회신</h2>")
    parts.append(f'<p class="meta">제목: <b>{html.escape(subject)}</b><br>받는 사람: {html.escape(c.get("to", ""))}'
                 f'<br>참조: {html.escape(c.get("cc", ""))}<br>아래 점선 안을 그대로 복사 → Gmail 에 붙여넣기</p>')
    # 담당자 실제 양식(2026-10-20)
    body = ["Dear PCG Team,<br>This is Gyeom Heo from Artience.<br><br>",
            f"I am sharing the <b>W{week}</b> Weekly Raw Ad Platform Report.<br>",
            (f"for this week’s update, only the <b>{this}</b> data has been included.<br>" if this else ""),
            f"{NB}<br>{NB}1) Attached File (CPCex Attached)<br>",
            (f'<ul style="margin:0;padding-left:24px"><li><b>{html.escape(tabs[this][1].name)}</b></li></ul>'
             if this else ""),
            f"{NB}2) Collection status<br>"]
    if this:                                       # 격주로 그 주 구분 표 하나 (홀수 주 CE, 짝수 주 MX)
        body.append(tabs[this][2])
    body.append("<br>Please let me know if you have any questions or require further information regarding this "
                f"{NB}report.<br><br>Best Regards<br>Gyeom")
    main_html = "".join(body)
    parts.append('<div class="box" id="main">' + main_html + "</div>")
    reminders = []
    # ---- 리마인드
    parts.append("<h2>② 리마인드 메일 — 미제출 법인별</h2>")
    if not this:
        parts.append(f"<p>W{week} 취합본이 없어 리마인드를 만들지 않음</p>")
    else:
        # 그 주 구분만(담당자 2026-10-20): W43(CE 주)이면 CE 미제출 법인 · CE 범위만.
        # 끝 주 = 이번 주-1 (지난 일요일까지), 시작 주 = 그 법인의 마지막 제출 주 + 1
        pend = {this: missing_ranges(this, tabs[this][1], week)}
        subs = list(dict.fromkeys([s_ for k in ("MX", "CE") for s_ in pend.get(k, {})]))
        if not subs:
            parts.append("<p>미제출 법인 없음</p>")
        for sub in subs:
            ranges = {k: pend[k][sub] for k in ("MX", "CE") if sub in pend.get(k, {})}
            text = reminder_body(sub, ranges)
            subj = f"[Reminder] {sub} Weekly Report — " + " / ".join(f"{k} {weeks_text(*v)}" for k, v in ranges.items())
            reminders.append({"market": sub, "subject": subj, "text": text})
            rid = f"r{len(reminders)}"
            parts.append(f'<p class="meta">{html.escape(sub)} — 제목: <b>{html.escape(subj)}</b> '
                         f'<button onclick="cp(&quot;{rid}&quot;)">본문 복사</button></p>'
                         f'<div class="box" id="{rid}" style="white-space:pre-wrap">{html.escape(text)}</div>')
    parts.append("<script>function cp(id){var r=document.createRange();"
                 "r.selectNodeContents(document.getElementById(id));var s=getSelection();s.removeAllRanges();"
                 "s.addRange(r);document.execCommand('copy');s.removeAllRanges();}"
                 "</script></body></html>")
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = REPORT_DIR / f"메일_W{week}.html"
    out.write_text("".join(parts), encoding="utf-8")
    import json
    payload = json.dumps({"week": week, "kind": this, "subject": subject, "html": main_html,
                          "thread_query": THREAD_QUERY, "reminders": reminders}, ensure_ascii=False, indent=1)
    (REPORT_DIR / f"메일_W{week}.json").write_text(payload, encoding="utf-8")
    if DRAFT_DROP.parent.exists():               # GAS 자동 초안용(설정했을 때만 동작) — 내 드라이브
        DRAFT_DROP.mkdir(exist_ok=True)
        (DRAFT_DROP / f"메일_W{week}.json").write_text(payload, encoding="utf-8")
    return out


# 담당자 방식: 지난주 보낸 메일에 전체 회신 + 제목만 그 주 것으로 → 본문 아래에 지난 메일이 쭉 인용돼 기록이 쌓인다.
# 그래서 기본은 '보낸편지함에서 가장 최근 GMO Weekly Report' 검색 — 맨 위가 지난주 메일, 링크를 바꿀 필요 없음.
# 리마인드(제목에 Reminder) 는 빼고, PCG PIC 에게 간 것, 최근 3주만(검색이 관련성순이어도 2~3통이라 바로 보인다)
SENT_QUERY = ('in:sent subject:"GMO Weekly Report" -subject:Reminder to:<PCG 담당자 메일> '
              'newer_than:21d')
# 브라우저마다 u/0 이 다른 계정일 수 있어 계정을 이름으로 지정
MAIL_ACCOUNT = "<회사 Gmail>"
THREAD_URL = (f"https://mail.google.com/mail/?authuser={MAIL_ACCOUNT}#search/"
              + __import__("urllib.parse").parse.quote_plus(SENT_QUERY))
MAIL_SETTINGS = Path(__file__).resolve().parent.parent / "3_Config" / "mail_settings.json"


def thread_url() -> str:
    """PCG + 삼성 공유 스레드 링크 — GUI 에서 바꾸면 3_Config/mail_settings.json 에 저장된다."""
    import json
    try:
        return json.loads(MAIL_SETTINGS.read_text(encoding="utf-8")).get("thread_url") or THREAD_URL
    except (OSError, ValueError):
        return THREAD_URL


def save_thread_url(url: str) -> None:
    import json
    data = {}
    if MAIL_SETTINGS.exists():
        try:
            data = json.loads(MAIL_SETTINGS.read_text(encoding="utf-8"))
        except ValueError:
            data = {}
    data["thread_url"] = url.strip()
    MAIL_SETTINGS.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
# GAS 가 같은 스레드를 찾는 검색어(가장 최근) — 새 제목 · 옛 제목 둘 다
THREAD_QUERY = ('in:sent (subject:"GMO Weekly Report" OR subject:"GMO D2C Weekly Report" '
                'OR subject:"GMO D2C Pacing Report")')


def copy_html_to_clipboard(fragment: str) -> None:
    """Windows 클립보드에 HTML(CF_HTML) + 글자로 — Gmail 에 Ctrl+V 하면 표 서식이 그대로 들어간다."""
    import re
    import win32clipboard
    crlf = chr(13) + chr(10)
    head = crlf.join(["Version:0.9", "StartHTML:{0:010d}", "EndHTML:{1:010d}",
                      "StartFragment:{2:010d}", "EndFragment:{3:010d}", ""])
    pre = '<html><body><!--StartFragment-->'
    post = "<!--EndFragment--></body></html>"
    dummy = head.format(0, 0, 0, 0)
    start_html = len(dummy.encode("utf-8"))
    start_frag = start_html + len(pre.encode("utf-8"))
    end_frag = start_frag + len(fragment.encode("utf-8"))
    end_html = end_frag + len(post.encode("utf-8"))
    data = head.format(start_html, end_html, start_frag, end_frag) + pre + fragment + post
    text = re.sub(r"<br\s*/?>", chr(10), fragment)
    text = html.unescape(re.sub(r"<[^>]+>", chr(9), text))
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"), data.encode("utf-8"))
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--no-open", action="store_true")
    ap.add_argument("--copy", action="store_true", help="본문을 클립보드에 넣고 지난주 보낸 메일(검색)을 연다")
    ap.add_argument("--copy-subject", action="store_true", help="제목만 클립보드에")
    a = ap.parse_args()
    if a.copy_subject:
        import json
        import win32clipboard
        subj = json.loads((REPORT_DIR / f"메일_W{a.week}.json").read_text(encoding="utf-8"))["subject"]
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, subj)
        finally:
            win32clipboard.CloseClipboard()
        print(f"  ✔ 제목 복사: {subj}")
        raise SystemExit(0)
    p = build(a.week)
    print(f"■ 메일 — {p}")
    if a.copy:
        import json
        import webbrowser
        copy_html_to_clipboard(json.loads(p.with_suffix(".json").read_text(encoding="utf-8"))["html"])
        webbrowser.open(thread_url())
        print("  ✔ 본문을 클립보드에 복사, 보낸편지함의 지난주 메일 검색을 열었음")
        print("    맨 위 메일 → 전체 회신 → 제목은 GUI [제목 복사] 로 바꾸기 → 본문 Ctrl+V → 보내기(직접)")
    if not a.no_open:
        os.startfile(p)  # noqa: S606
