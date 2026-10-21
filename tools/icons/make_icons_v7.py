"""바로가기 아이콘 생성 (W. / R.) — v7 (2026-10-21): v5 의 타일 · 글자 높이 · 광학 중앙(2% 위) · 우하단 빨간 점을 그대로 두고,
   글자만 손으로 그린 픽셀 글꼴(굵은 획 2칸, 7줄)로 바꾼다. 픽셀 사이 틈은 두지 않는다(틈이 있으면 격자로 보여 반려).
   - 글자 높이는 두 글자 모두 v5 R 높이, 가로는 타일 가운데.
   - 자동 픽셀화(Pretendard 를 격자로 샘플링)와 v6(5 × 5 격자)는 글자가 R · W 로 안 읽혀 반려(2026-10-21).
   - 바꾼 뒤엔 파일 이름 버전을 올려야 Windows 아이콘 캐시가 새로 읽는다(_v7 → _v8, gui.py · gui_prep.py · 바로가기도 같이).
   이전 생성기: make_icons_v5.py(Pretendard Black 글자) · make_icons_v6.py(5 × 5 격자)."""
import math, os
import numpy as np
from PIL import Image, ImageDraw, ImageFont

PF = r"C:\Users\GyeomHeo\AppData\Local\Microsoft\Windows\Fonts\Pretendard-Black.ttf"
INK, WHITE, ACC = (17, 17, 17, 255), (255, 255, 255, 255), (255, 77, 46, 255)
S = 1024

GLYPHS = {  # X = 흰 픽셀
    "R": ["XXXXX.",
          "XX..XX",
          "XX..XX",
          "XXXXX.",
          "XX.XX.",
          "XX..XX",
          "XX..XX"],
    "W": ["XX...XX",
          "XX...XX",
          "XX.X.XX",
          "XX.X.XX",
          "XXXXXXX",
          ".XX.XX.",
          ".XX.XX."],
}


def v5_height(small):
    """v5 에서 R 글자가 차지하던 높이(잉크 면적 기준 크기)."""
    def glyph(size):
        m = Image.new("L", (S * 2, S * 2), 0)
        ImageDraw.Draw(m).text((S // 2, S // 2), "R", font=ImageFont.truetype(PF, size), fill=255)
        a = np.array(m) > 127; ys, xs = np.where(a)
        return a[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
    area = (S * (0.44 if small else 0.33)) ** 2
    g = glyph(400)
    g = glyph(int(400 * math.sqrt(area / g.sum())))
    return g.shape[0]


def draw(ch, small):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    pad, rad = (0, 170) if small else (32, 200)
    d.rounded_rectangle((pad, pad, S - pad, S - pad), radius=rad, fill=INK)
    rows = GLYPHS[ch]; ny, nx = len(rows), len(rows[0])
    gh = v5_height(small); cell = gh / ny
    x0, y0 = S / 2 - nx * cell / 2, S / 2 - gh / 2 - S * 0.02
    for r, line in enumerate(rows):
        for c, v in enumerate(line):
            if v == "X":
                x, y = x0 + c * cell, y0 + r * cell
                d.rectangle((round(x), round(y), round(x + cell) - 1, round(y + cell) - 1), fill=WHITE)
    # 점: v5 와 같다 — 모서리 원 중심에서 대각선 방향, 둥근 변과 등간격
    dot = int(S * (0.15 if small else 0.10)); inset = S * (0.075 if small else 0.085)
    c0 = S - pad - rad; arc = c0 + rad / math.sqrt(2)
    corner = arc - inset / math.sqrt(2) - inset * 0.5
    d.rectangle((corner - dot, corner - dot, corner, corner), fill=ACC)
    return im


def save(ch, path):
    big, sm = draw(ch, False), draw(ch, True)
    imgs = {s: (sm if s <= 24 else big).resize((s, s), Image.LANCZOS) for s in (16, 24, 32, 48, 64, 128, 256)}
    imgs[256].save(path, format="ICO", sizes=[(s, s) for s in imgs], append_images=[imgs[s] for s in (16, 24, 32, 48, 64, 128)])
    return imgs


if __name__ == "__main__":
    HERE = os.path.dirname(os.path.abspath(__file__))
    a = save("W", r"C:\GMO_Cleansing\assets\icon_weekly_v7.ico")
    b = save("R", r"C:\Weekly ROAS\draft_tools\assets\icon_roas_v7.ico")
    sh = Image.new("RGB", (620, 330), (240, 240, 240))
    for i, im in enumerate((a, b)):
        X = 10 + i * 300; sh.paste(im[256], (X, 10), im[256]); x = X
        for s in (48, 32, 24, 16):
            sh.paste(im[s], (x, 320 - s), im[s]); x += s + 16
    sh.save(os.path.join(HERE, "pv.png"))
    print(os.path.join(HERE, "pv.png"))
