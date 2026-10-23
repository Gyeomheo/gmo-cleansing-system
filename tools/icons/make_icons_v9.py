"""바로가기 아이콘 생성 (W / R) — v9 (2026-10-23): v8 그림 그대로, 작업표시줄에서 깨져 보이던 것만 고침.
   - v8 은 32 · 48px 도 픽셀 사이 틈이 있는 큰 그림을 줄여 썼다 → 1px 틈이 회색으로 뭉개짐(Windows 11 작업표시줄은 32px 를 24px 로 또 줄인다).
   - v9: 64px 미만은 틈 없이, 그리고 1024px 를 줄이지 않고 그 크기에서 직접 그린다 — 칸 크기를 정수 px 로 맞춰 글자 가장자리가 흐려지지 않게.
     타일(둥근 사각형)만 8배로 그려 줄인다. 64px 이상은 v8 과 같다(틈 있는 그림을 LANCZOS 로 줄임).
   - 바꾼 뒤엔 파일 이름 버전을 올려야 Windows 아이콘 캐시가 새로 읽는다(gui.py · gui_prep.py 도 같이).
   이전 생성기: make_icons_v8.py(담당자 선택 시안 — 글꼴 · 빨간 픽셀 위치는 거기서 정함)."""
import os
from PIL import Image, ImageDraw

from make_icons_v8 import ACC, GLYPHS, INK, WHITE, draw

# 크기별 정수 칸 크기 — 글자 높이 7칸이 타일의 44~66% (16px 은 2px 칸이면 14px 로 꽉 차서 1px)
CELL = {16: 1, 24: 2, 32: 3, 48: 4}


def draw_small(ch, s):
    k = 8  # 타일 가장자리만 부드럽게
    tile = Image.new("RGBA", (s * k, s * k), (0, 0, 0, 0))
    ImageDraw.Draw(tile).rounded_rectangle((0, 0, s * k - 1, s * k - 1), radius=round(s * k * 0.17), fill=INK)
    im = tile.resize((s, s), Image.LANCZOS); d = ImageDraw.Draw(im)
    rows = GLYPHS[ch]; ny, nx = len(rows), len(rows[0]); cell = CELL[s]
    x0, y0 = (s - nx * cell) // 2, (s - ny * cell) // 2
    for r, line in enumerate(rows):
        for c, v in enumerate(line):
            if v != ".":
                x, y = x0 + c * cell, y0 + r * cell
                d.rectangle((x, y, x + cell - 1, y + cell - 1), fill=ACC if v == "A" else WHITE)
    return im


def save(ch, path):
    big = draw(ch, False)
    imgs = {s: draw_small(ch, s) if s in CELL else big.resize((s, s), Image.LANCZOS) for s in (16, 24, 32, 48, 64, 128, 256)}
    imgs[256].save(path, format="ICO", sizes=[(s, s) for s in imgs], append_images=[imgs[s] for s in (16, 24, 32, 48, 64, 128)])
    return imgs


if __name__ == "__main__":
    HERE = os.path.dirname(os.path.abspath(__file__))
    a = save("W", r"C:\GMO_Cleansing\assets\icon_weekly_v9.ico")
    b = save("R", r"C:\Weekly ROAS\draft_tools\assets\icon_roas_v9.ico")
    sh = Image.new("RGB", (620, 330), (240, 240, 240))
    for i, im in enumerate((a, b)):
        X = 10 + i * 300; sh.paste(im[256], (X, 10), im[256]); x = X
        for s in (48, 32, 24, 16):
            sh.paste(im[s], (x, 320 - s), im[s]); x += s + 16
    sh.save(os.path.join(HERE, "pv.png"))
    print(os.path.join(HERE, "pv.png"))
