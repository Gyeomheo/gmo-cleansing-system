"""GMO 클렌징 GUI — 주차 고르기 → 수집 · 클렌징 → 취합 검증 → 취합본 만들기 → 메일.

    GMO클렌징_GUI.cmd 더블클릭 (또는 python tools/gui.py)

[▶ 전체 실행] = ① → ② → 검증에 오류가 없으면 ③ → ④ 메일(본문 클립보드 + Gmail 스레드 열기, 보내기는 직접).
디자인은 Weekly ROAS 데이터 준비 GUI 와 같다(흑백 · 한 글꼴 · 평평한 버튼 · 왼쪽 상태 / 오른쪽 실행 · 로그).
"""
from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from collect import WEEKLY_ROOT
from common import REPORT_DIR, ROOT, config
from consolidate import DASH_DIR, NAME_RE, latest_file
from weekly_mail import save_thread_url, thread_url

# 흑백 팔레트 (Weekly ROAS GUI 와 같다)
INK, SUB, FAINT, LINE, BG, PANEL = "#111111", "#6B6B6B", "#A3A3A3", "#E4E4E4", "#FFFFFF", "#F6F6F6"


def suggest_kind(week: int) -> str:
    """격주: 홀수 주 CE, 짝수 주 MX (담당자 2026-10-20)."""
    return "CE" if week % 2 else "MX"


class Btn(tk.Label):
    """평평한 버튼 — primary: 검정 바탕 흰 글자, 아니면 흰 바탕 검정 테두리"""

    def __init__(self, master, text, cmd, primary=False, small=False, **kw):
        self.primary, self.cmd, self.enabled = primary, cmd, True
        top = master.winfo_toplevel()
        super().__init__(master, text=text, padx=10 if small else 14, pady=3 if small else 5, cursor="hand2",
                         font=top.f_small if small else top.f_body,
                         bg=INK if primary else BG, fg=BG if primary else INK,
                         highlightthickness=1, highlightbackground=INK, highlightcolor=INK, **kw)
        self.bind("<Button-1>", lambda e: self.enabled and self.cmd())
        self.bind("<Enter>", lambda e: self.enabled and self.config(bg="#333333" if primary else PANEL))
        self.bind("<Leave>", lambda e: self.paint())

    def paint(self):
        if not self.enabled:
            self.config(bg=BG, fg=FAINT, highlightbackground=LINE)
        else:
            self.config(bg=INK if self.primary else BG, fg=BG if self.primary else INK, highlightbackground=INK)

    def set_enabled(self, on):
        self.enabled = on
        self.paint()


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("GMO D2C · 클렌징 · 취합")
        try:  # 작업표시줄을 python 대신 앱 아이콘으로
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Artience.WeeklyReport")
            self.iconbitmap(default=str(ROOT / "assets" / "icon_weekly_v8.ico"))
        except Exception:
            pass
        self.geometry("1280x800")
        self.minsize(1100, 680)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        self.configure(bg=BG)
        fam = next((f for f in ("Pretendard", "Pretendard Variable", "SamsungOneKorean 400", "Malgun Gothic")
                    if f in tkfont.families()), "Malgun Gothic")
        self.fam = fam
        self.f_body, self.f_small, self.f_bold = (fam, 10), (fam, 9), (fam, 10, "bold")
        self.f_head, self.f_title, self.f_mark = (fam, 11, "bold"), (fam, 16, "bold"), (fam, 12)
        self.option_add("*Font", self.f_body)
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.check_ok = False

        st = ttk.Style()
        st.theme_use("clam")
        st.configure("TCombobox", fieldbackground=BG, background=BG, foreground=INK, arrowcolor=INK, bordercolor=INK,
                     lightcolor=BG, darkcolor=BG, padding=4)
        st.map("TCombobox", fieldbackground=[("readonly", BG)], selectbackground=[("readonly", BG)],
               selectforeground=[("readonly", INK)])
        st.configure("Vertical.TScrollbar", background=PANEL, troughcolor=BG, bordercolor=BG, arrowcolor=SUB,
                     lightcolor=PANEL, darkcolor=PANEL)

        self._build_head()
        left, main = self._build_body()
        self._build_left(left)
        self._build_right(main)
        self.on_week()
        self.q.put(("say", "주차 · 취합 구분을 확인하고 ▶ 전체 실행.  보내기는 Gmail 에서 직접\n", "dim"))
        self.after(100, self.pump)

    # ───────── 화면 ─────────
    def _build_head(self) -> None:
        fam = self.fam
        head = tk.Frame(self, bg=BG)
        head.pack(fill="x", padx=28, pady=(22, 0))
        tk.Label(head, text="GMO D2C", font=self.f_title, bg=BG, fg=INK).pack(side="left")
        tk.Label(head, text="  클렌징 · 취합", font=(fam, 16), bg=BG, fg=SUB).pack(side="left")
        right = tk.Frame(head, bg=BG)
        right.pack(side="right")
        tk.Label(right, text="주차", bg=BG, fg=SUB, font=self.f_small).pack(side="left", padx=(0, 6))
        this_week = date.today().isocalendar()[1]
        self.wk = tk.StringVar(value=f"W{this_week}")
        box = ttk.Combobox(right, textvariable=self.wk, width=6, state="readonly", font=self.f_bold,
                           values=[f"W{w}" for w in range(this_week + 1, 0, -1)])
        box.pack(side="left")
        box.bind("<<ComboboxSelected>>", lambda e: self.on_week())
        tk.Label(right, text="취합", bg=BG, fg=SUB, font=self.f_small).pack(side="left", padx=(16, 6))
        self.kind = tk.StringVar(value=suggest_kind(this_week))
        self.kind_btns = {k: Btn(right, k, lambda k=k: self.set_kind(k), small=True) for k in ("CE", "MX")}
        for b in self.kind_btns.values():
            b.pack(side="left", padx=(0, 4))
        self.sub_lbl = tk.Label(self, text="", bg=BG, fg=SUB, font=self.f_small, anchor="w")
        self.sub_lbl.pack(fill="x", padx=28, pady=(2, 12))
        tk.Frame(self, bg=INK, height=2).pack(fill="x", padx=28)


    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=28, pady=(16, 20))
        left = tk.Frame(body, bg=BG, width=430)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        tk.Frame(body, bg=LINE, width=1).pack(side="left", fill="y", padx=22)
        main = tk.Frame(body, bg=BG)
        main.pack(side="left", fill="both", expand=True)

        return left, main

    def _build_left(self, left) -> None:
        self.section(left, "이번 주 원본", "● 들어옴  ○ 없음  – 제출 안 함")
        self.src_fr = tk.Frame(left, bg=BG)
        self.src_fr.pack(fill="x")
        tk.Frame(left, bg=BG, height=16).pack()
        self.section(left, "작업 순서", "파일로 자동 체크")
        self.step_fr = tk.Frame(left, bg=BG)
        self.step_fr.pack(fill="x")


    def _build_right(self, main) -> None:
        fam = self.fam
        self.section(main, "경로")
        self.root_dir = tk.StringVar()
        self.thread = tk.StringVar(value=thread_url())
        self.path_row(main, "원본 폴더", self.root_dir, [("폴더 선택…", self.pick)])
        self.path_row(main, "취합 폴더", tk.StringVar(value=str(DASH_DIR)), [("열기", lambda: self.open_path(None))],
                      readonly=True)
        self.path_row(main, "Gmail 열 곳", self.thread,
                      [("저장", self.save_thread), ("열기", lambda: webbrowser.open(self.thread.get().strip()))])
        tk.Label(main, text=("기본 = 보낸편지함의 가장 최근 'GMO Weekly Report'(리마인드 제외) — "
                             "지난주 메일에 전체 회신, 기록이 아래로 쌓임"),
                 bg=BG, fg=FAINT, font=self.f_small, anchor="w").pack(fill="x", pady=(2, 0))

        act = tk.Frame(main, bg=BG)
        act.pack(fill="x", pady=(18, 8))
        tk.Label(act, text="실행", font=self.f_head, bg=BG, fg=INK).pack(side="left")
        self.run_btn = Btn(act, "▶  전체 실행", self.run_all, primary=True)
        self.run_btn.pack(side="left", padx=(12, 10))
        self.btns = [self.run_btn]
        for t, fns in (("① 수집 · 클렌징", [self.step_clean]), ("② 검증", [self.step_check]),
                       ("③ 취합본", [self.step_make, self.step_status]), ("④ 메일", [self.step_mail])):
            b = Btn(act, t, lambda fns=fns: self.start(fns))
            b.pack(side="left", padx=(0, 6))
            self.btns.append(b)
        Btn(act, "제목 복사", self.copy_subject, small=True).pack(side="left", padx=(4, 0))
        tk.Label(act, text="Excel 은 모두 닫고", bg=BG, fg=SUB, font=self.f_small).pack(side="left", padx=8)
        srow = tk.Frame(main, bg=BG)
        srow.pack(fill="x", pady=(0, 6))
        self.status = tk.StringVar(value="대기")
        tk.Label(srow, textvariable=self.status, bg=BG, fg=SUB, font=self.f_small, anchor="w").pack(side="left")
        res = tk.Frame(srow, bg=BG)
        res.pack(side="right")
        tk.Label(res, text="결과", bg=BG, fg=SUB, font=self.f_small).pack(side="left", padx=(0, 6))
        for t, fn in (("실행요약", lambda: REPORT_DIR / "실행요약.md"),
                      ("취합검증", lambda: REPORT_DIR / f"취합검증_{self.kind.get()}.md"),
                      ("메일", lambda: REPORT_DIR / f"메일_W{self.week()}.html"),
                      ("OUTPUT", lambda: config.OUTPUT_DIR)):
            Btn(res, t, lambda fn=fn: self.open_path(fn()), small=True).pack(side="left", padx=2)
        lf = tk.Frame(main, bg=PANEL)
        lf.pack(fill="both", expand=True)
        self.log = tk.Text(lf, font=(fam, 9), bg=PANEL, fg=INK, wrap="none", bd=0, padx=12, pady=10,
                           highlightthickness=0, insertbackground=INK)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.log.yview, style="Vertical.TScrollbar")
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("bad", font=(fam, 9, "bold"))
        self.log.tag_configure("dim", foreground=SUB)

    # ───────── 그리기 도우미 ─────────
    def section(self, parent, title, sub=""):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=(0, 8))
        tk.Label(f, text=title, font=self.f_head, bg=BG, fg=INK).pack(side="left")
        if sub:
            tk.Label(f, text=sub, font=self.f_small, bg=BG, fg=FAINT).pack(side="right")

    def path_row(self, parent, label, var, buttons, readonly=False):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=2)
        tk.Label(row, text=label, bg=BG, fg=SUB, font=self.f_small, width=11, anchor="w").pack(side="left")
        e = tk.Entry(row, textvariable=var, bg=PANEL if readonly else BG, fg=INK, relief="flat", font=self.f_body,
                     highlightthickness=1, highlightbackground=LINE, highlightcolor=INK, readonlybackground=PANEL)
        if readonly:
            e.config(state="readonly")
        e.pack(side="left", fill="x", expand=True, ipady=4)
        for t, fn in buttons:
            Btn(row, t, fn, small=True).pack(side="left", padx=(6, 0))

    def week(self) -> int:
        m = re.fullmatch(r"W(\d+)", self.wk.get())
        return int(m.group(1)) if m else date.today().isocalendar()[1]

    # ───────── 상태 ─────────
    def set_kind(self, k: str) -> None:
        self.kind.set(k)
        self.refresh()

    def on_week(self) -> None:
        self.kind.set(suggest_kind(self.week()))
        self.root_dir.set(str(WEEKLY_ROOT / f"{self.week()}주차"))
        self.refresh()

    def refresh(self) -> None:
        for k, b in self.kind_btns.items():
            b.primary = k == self.kind.get()
            b.paint()
        w, k = self.week(), self.kind.get()
        self.sub_lbl.config(text=f"W{w} · {k} 주 (홀수 CE · 짝수 MX)   원본 {Path(self.root_dir.get()).name} \\ {k}")
        self.draw_sources()
        self.draw_steps()

    def draw_sources(self) -> None:
        for c in self.src_fr.winfo_children():
            c.destroy()
        k = self.kind.get()
        base = Path(self.root_dir.get()) / k
        folders = {}
        if base.exists():
            for d in sorted(p for p in base.iterdir() if p.is_dir()):
                m = re.match(r"\s*([A-Za-z]+)\s*(\d{3,4})?", d.name)
                if m:
                    folders.setdefault(m.group(1).upper(), []).append(d.name)
        try:
            import pandas as pd
            mk = pd.read_csv(config.CONFIG_DIR / "tracker_markets.csv", dtype=str, encoding="utf-8-sig").fillna("")
            mk = mk[mk["구분"].str.upper() == k]
            markets = [(r.Market, r.추적.upper() != "N") for r in mk.itertuples(index=False)]
        except Exception:
            markets = [(s, True) for s in folders]
        for s in folders:
            if s not in dict(markets):
                markets.append((s, True))
        if not base.exists():
            tk.Label(self.src_fr, text=f"폴더 없음 — {base}", bg=PANEL, fg=SUB, font=self.f_small, anchor="w",
                     justify="left", wraplength=400, padx=10, pady=8).pack(fill="x")
        for i, (s, tracked) in enumerate(markets):
            if i:
                tk.Frame(self.src_fr, bg=LINE, height=1).pack(fill="x")
            line = tk.Frame(self.src_fr, bg=BG)
            line.pack(fill="x", pady=3)
            have = s in folders
            mark, fg = ("●", INK) if have else (("○", INK) if tracked else ("–", FAINT))
            tk.Label(line, text=mark, font=self.f_mark, bg=BG, fg=fg, width=2).pack(side="left")
            tk.Label(line, text=s, font=self.f_bold if (tracked and not have) else self.f_body, bg=BG, fg=fg,
                     width=8, anchor="w").pack(side="left")
            detail = " · ".join(f.split(" ", 1)[-1] for f in folders.get(s, [])) if have else (
                "이번 주 없음" if tracked else "제출 안 함")
            tk.Label(line, text=detail, font=self.f_small, bg=BG, fg=SUB, anchor="w").pack(side="left")

    def step_states(self) -> list[tuple[str, str, bool]]:
        w, k = self.week(), self.kind.get()
        run_csv = config.OUTPUT_DIR / f"_이번실행_{k}.csv"
        check = REPORT_DIR / f"취합검증_{k}.md"
        f = latest_file(k)
        made = bool(f and NAME_RE.match(f.name) and int(NAME_RE.match(f.name).group(2)) == w)
        return [
            ("수집 · 클렌징", f"{w}주차 {k} 폴더 → 1_Input → run.py", run_csv.exists()),
            ("취합 검증", "최신 취합본 RAW 기준 · 법인별 교체 · 검증 12항목",
             check.exists() and f"W{w}_" in check.read_text(encoding="utf-8", errors="ignore")),
            ("취합본 만들기", (f.name if made else "새 버전 생성 · 이전 것은 old RAW files · Tracker"), made),
            ("메일", "본문 → 클립보드 + 지난주 보낸 메일 열기 · 리마인드 HTML",
             (REPORT_DIR / f"메일_W{w}.json").exists()),
            ("보내기 (사람)", "지난주 메일에 전체 회신 → [제목 복사] → 본문 Ctrl+V → 보내기 · 파일은 CPCex", False),
        ]

    def draw_steps(self) -> None:
        for c in self.step_fr.winfo_children():
            c.destroy()
        states = self.step_states()
        now = next((i for i, s in enumerate(states) if not s[2]), None)
        for i, (name, what, done) in enumerate(states):
            if i:
                tk.Frame(self.step_fr, bg=LINE, height=1).pack(fill="x")
            line = tk.Frame(self.step_fr, bg=BG)
            line.pack(fill="x", pady=4)
            tk.Label(line, text=f"{i + 1}", font=self.f_small, bg=BG, fg=INK if i == now else FAINT,
                     width=2).pack(side="left")
            tk.Label(line, text="■" if done else "□", font=self.f_mark, bg=BG,
                     fg=INK if (done or i == now) else SUB, width=2).pack(side="left")
            txt = tk.Frame(line, bg=BG)
            txt.pack(side="left", fill="x", expand=True)
            tk.Label(txt, text=name + ("   ← 지금" if i == now else ""), font=self.f_bold if i == now else self.f_body,
                     bg=BG, fg=INK, anchor="w").pack(fill="x")
            tk.Label(txt, text=what, font=self.f_small, bg=BG, fg=SUB, anchor="w", justify="left",
                     wraplength=330).pack(fill="x")

    # ───────── 경로 ─────────
    def pick(self) -> None:
        d = filedialog.askdirectory(initialdir=str(WEEKLY_ROOT), title="주차 폴더(안에 MX · CE 폴더) 선택")
        if d:
            m = re.search(r"(\d+)\s*주차", Path(d).name)
            if m:
                self.wk.set(f"W{int(m.group(1))}")
                self.kind.set(suggest_kind(self.week()))
            self.root_dir.set(d)
            self.refresh()

    def save_thread(self) -> None:
        """새 스레드로 보냈으면 그 링크를 붙여 넣고 저장 — 다음부터 [④ 메일] 이 이 스레드를 연다."""
        url = self.thread.get().strip()
        if not url.startswith("https://mail.google.com/"):
            messagebox.showwarning("Gmail", "Gmail 주소(https://mail.google.com/…)를 붙여 넣으세요.")
            return
        save_thread_url(url)
        self.status.set("✔ Gmail 열 곳 저장")

    def open_path(self, p: Path | None) -> None:
        p = DASH_DIR if p is None else p
        if p.exists():
            os.startfile(p)  # noqa: S606
        else:
            messagebox.showinfo("없음", f"아직 없습니다:\n{p}")

    # ───────── 실행 ─────────
    def call(self, args: list[str]) -> int:
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONPATH=str(ROOT))
        cmd = [sys.executable.replace("pythonw.exe", "python.exe"), "-X", "utf8"] + args
        self.q.put(("say", f"\n▸ {' '.join(Path(a).name if a.endswith('.py') else a for a in args)}\n", "dim"))
        p = subprocess.Popen(cmd, cwd=ROOT, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace",
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        for line in p.stdout:
            if "Could not determine dtype" not in line and " - INFO - " not in line:
                self.q.put(("say", line, "bad" if ("✖" in line or "Traceback" in line) else ""))
        return p.wait()

    def step_clean(self) -> bool:
        self.status.set("① 수집 · 클렌징 중…")
        return self.call([str(ROOT / "tools" / "run_all.py"), "--root", self.root_dir.get()]) == 0

    def step_check(self) -> bool:
        self.status.set("② 취합 검증 중…")
        rc = self.call([str(ROOT / "tools" / "consolidate.py"), self.kind.get(), "--week", str(self.week())])
        self.check_ok = rc == 0
        return self.check_ok

    def step_make(self) -> bool:
        f = latest_file(self.kind.get())
        m = NAME_RE.match(f.name) if f else None
        nl = chr(10)
        if m and int(m.group(2)) == self.week() and not messagebox.askyesno(
                "확인", f"W{self.week()} {self.kind.get()} 취합본이 이미 있습니다:{nl}{f.name}{nl}{nl}"
                        f"새 버전(v{int(m.group(3)) + 1})을 만들까요?"):
            self.q.put(("say", f"  · 취합본 만들기 건너뜀(이미 있음){nl}", "dim"))
            return True                               # 건너뛰고 ④ 메일로
        if not self.check_ok and not messagebox.askyesno("확인", "② 검증을 통과하지 않았습니다. 그래도 만들까요?"):
            return False
        self.status.set("③ 취합본 만드는 중… (Excel 1~2분)")
        return self.call([str(ROOT / "tools" / "consolidate.py"), self.kind.get(), "--apply",
                          "--week", str(self.week())]) == 0

    def step_status(self) -> bool:
        """국가별 weekly 폴더 처리 표시(_A · XXX · 허겸수정 클렌징본) + 백과사전 STATUS 시트. 실패해도 메일로 넘어간다."""
        self.status.set("③-2 처리 표시 · STATUS 기재 중… (Excel)")
        w, k = str(self.week()), self.kind.get()
        rc1 = self.call([str(ROOT / "tools" / "mark_done.py"), w, k, "--apply"])
        rc2 = self.call([str(ROOT / "tools" / "status_check.py"), w, k, "--apply"])
        if rc1 or rc2:
            self.q.put(("say", "  ✖ 처리 표시 · STATUS 중 실패 — 로그 확인(메일은 계속)\n", "bad"))
        return True

    def step_mail(self) -> bool:
        """주간 메일 본문 → 클립보드 + Gmail 스레드 열기, 리마인드는 HTML 에서 [본문 복사]. 보내기는 직접."""
        self.status.set("④ 메일 만드는 중…")
        return self.call([str(ROOT / "tools" / "weekly_mail.py"), "--week", str(self.week()), "--copy"]) == 0

    def copy_subject(self) -> None:
        """그 주 메일 제목 → 클립보드 (회신 창 제목 칸에 붙여넣기)."""
        import json
        f = REPORT_DIR / f"메일_W{self.week()}.json"
        if not f.exists():
            messagebox.showinfo("제목 복사", "먼저 ④ 메일을 실행하세요.")
            return
        subj = json.loads(f.read_text(encoding="utf-8"))["subject"]
        self.clipboard_clear()
        self.clipboard_append(subj)
        self.status.set(f"✔ 제목 복사: {subj}")

    def run_all(self) -> None:
        self.start([self.step_clean, self.step_check, self.step_make, self.step_status, self.step_mail])

    def start(self, steps) -> None:
        if self.busy:
            return
        self.busy = True
        for b in self.btns:
            b.set_enabled(False)

        def work():
            ok = True
            for s in steps:
                ok = s()
                if not ok:
                    break
            self.q.put(("done", ok))
        threading.Thread(target=work, daemon=True).start()

    def pump(self) -> None:
        try:
            while True:
                item = self.q.get_nowait()
                if item[0] == "done":
                    self.busy = False
                    for b in self.btns:
                        b.set_enabled(True)
                    self.status.set("✔ 끝" if item[1] else "✖ 멈춤 — 로그 확인")
                    self.refresh()
                else:
                    _, text, tag = item
                    self.log.insert("end", text, tag or ())
                    self.log.see("end")
        except queue.Empty:
            pass
        self.after(100, self.pump)


if __name__ == "__main__":
    App().mainloop()
