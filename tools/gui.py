"""GMO 클렌징 GUI — 폴더 고르기 → 수집 · 클렌징 → 취합 검증 → 취합본 만들기.

    GMO클렌징_GUI.cmd 더블클릭 (또는 python tools/gui.py)

[▶ 전체 실행] = ① → ② → 검증에 오류가 없으면 ③ 까지 자동. 각 단계는 tools/ 의 명령을 그대로 부른다.
"""
from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from collect import WEEKLY_ROOT
from common import REPORT_DIR, ROOT, config
from consolidate import DASH_DIR, NAME_RE, latest_file

FONT = ("Malgun Gothic", 10)
MONO = ("Consolas", 9)


def suggest_kind() -> str:
    """최신 취합본이 더 오래된 쪽(격주로 번갈아 나옴)."""
    weeks = {}
    for k in ("MX", "CE"):
        f = latest_file(k)
        m = NAME_RE.match(f.name) if f else None
        weeks[k] = int(m.group(2)) if m else 0
    return min(weeks, key=weeks.get)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("GMO 클렌징 · 취합")
        self.geometry("980x700")
        self.option_add("*Font", FONT)
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.check_ok = False

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="주차").grid(row=0, column=0, sticky="w")
        self.week = tk.IntVar(value=date.today().isocalendar()[1])
        ttk.Spinbox(top, from_=1, to=53, width=5, textvariable=self.week,
                    command=self.refresh_paths).grid(row=0, column=1, sticky="w", padx=(4, 16))
        ttk.Label(top, text="취합").grid(row=0, column=2, sticky="w")
        self.kind = tk.StringVar(value=suggest_kind())
        for i, k in enumerate(("CE", "MX")):
            ttk.Radiobutton(top, text=k, value=k, variable=self.kind).grid(row=0, column=3 + i, padx=2)
        ttk.Label(top, text="(최신 취합본이 오래된 쪽이 기본)").grid(row=0, column=5, sticky="w", padx=8)

        ttk.Label(top, text="원본 폴더").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.root_dir = tk.StringVar()
        ttk.Entry(top, textvariable=self.root_dir, width=90).grid(row=1, column=1, columnspan=5, sticky="we",
                                                                  pady=(8, 0))
        ttk.Button(top, text="폴더 선택…", command=self.pick).grid(row=1, column=6, padx=6, pady=(8, 0))
        ttk.Label(top, text="취합 폴더").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.dash = tk.StringVar(value=str(DASH_DIR))
        ttk.Entry(top, textvariable=self.dash, width=90, state="readonly").grid(row=2, column=1, columnspan=5,
                                                                                 sticky="we", pady=(4, 0))
        self.refresh_paths()

        steps = ttk.Frame(self, padding=(10, 4))
        steps.pack(fill="x")
        self.btns = [
            ttk.Button(steps, text="▶ 전체 실행", command=self.run_all),
            ttk.Button(steps, text="① 수집 · 클렌징", command=lambda: self.start([self.step_clean])),
            ttk.Button(steps, text="② 취합 검증", command=lambda: self.start([self.step_check])),
            ttk.Button(steps, text="③ 취합본 만들기", command=lambda: self.start([self.step_make])),
        ]
        for b in self.btns:
            b.pack(side="left", padx=4)
        ttk.Separator(steps, orient="vertical").pack(side="left", fill="y", padx=8)
        for label, path in (("OUTPUT", config.OUTPUT_DIR), ("Reports", REPORT_DIR), ("취합 폴더", None)):
            ttk.Button(steps, text=label, command=lambda p=path: self.open_path(p)).pack(side="left", padx=2)
        ttk.Button(steps, text="실행요약", command=lambda: self.open_path(REPORT_DIR / "실행요약.md")).pack(side="left")
        ttk.Button(steps, text="취합검증", command=lambda: self.open_path(
            REPORT_DIR / f"취합검증_{self.kind.get()}.md")).pack(side="left", padx=2)

        self.status = tk.StringVar(value="대기")
        ttk.Label(self, textvariable=self.status, padding=(12, 2)).pack(fill="x")
        self.log = tk.Text(self, font=MONO, wrap="none", bg="white", fg="black")
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.after(100, self.pump)

    # ------------------------------------------------------------ 경로
    def refresh_paths(self) -> None:
        self.root_dir.set(str(WEEKLY_ROOT / f"{self.week.get()}주차"))

    def pick(self) -> None:
        d = filedialog.askdirectory(initialdir=str(WEEKLY_ROOT), title="주차 폴더(안에 MX · CE 폴더) 선택")
        if d:
            self.root_dir.set(d)
            m = re.search(r"(\d+)\s*주차", Path(d).name)
            if m:
                self.week.set(int(m.group(1)))

    def open_path(self, p: Path | None) -> None:
        p = Path(self.dash.get()) if p is None else p
        if p.exists():
            os.startfile(p)  # noqa: S606
        else:
            messagebox.showinfo("없음", f"아직 없습니다:\n{p}")

    # ------------------------------------------------------------ 실행
    def call(self, args: list[str]) -> int:
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONPATH=str(ROOT))
        cmd = [sys.executable, "-X", "utf8"] + args
        self.q.put(f"\n$ {' '.join(Path(a).name if a.endswith('.py') else a for a in args)}\n")
        p = subprocess.Popen(cmd, cwd=ROOT, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace",
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        for line in p.stdout:
            if "Could not determine dtype" not in line and " - INFO - " not in line:
                self.q.put(line)
        return p.wait()

    def step_clean(self) -> bool:
        self.status.set("① 수집 · 클렌징 중…")
        return self.call([str(ROOT / "tools" / "run_all.py"), "--root", self.root_dir.get()]) == 0

    def step_check(self) -> bool:
        self.status.set("② 취합 검증 중…")
        rc = self.call([str(ROOT / "tools" / "consolidate.py"), self.kind.get(), "--week", str(self.week.get())])
        self.check_ok = rc == 0
        return self.check_ok

    def step_make(self) -> bool:
        if not self.check_ok and not messagebox.askyesno("확인", "② 검증을 통과하지 않았습니다. 그래도 만들까요?"):
            return False
        self.status.set("③ 취합본 만드는 중… (Excel 1~2분)")
        return self.call([str(ROOT / "tools" / "consolidate.py"), self.kind.get(), "--apply",
                          "--week", str(self.week.get())]) == 0

    def run_all(self) -> None:
        self.start([self.step_clean, self.step_check, self.step_make])

    def start(self, steps) -> None:
        if self.busy:
            return
        self.busy = True
        for b in self.btns:
            b.state(["disabled"])

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
                if isinstance(item, tuple):
                    self.busy = False
                    for b in self.btns:
                        b.state(["!disabled"])
                    self.status.set("✔ 끝" if item[1] else "✖ 멈춤 — 로그 확인")
                else:
                    self.log.insert("end", item)
                    self.log.see("end")
        except queue.Empty:
            pass
        self.after(100, self.pump)


if __name__ == "__main__":
    App().mainloop()
