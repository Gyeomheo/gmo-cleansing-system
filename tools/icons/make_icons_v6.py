"""바로가기 아이콘 생성 (W. / R.) — v6 (2026-10-21): Artience 로고의 픽셀 격자를 따라 글자를 흰 정사각 픽셀로 쌓는다.
   - INK 둥근 타일 위에 5 × 5 픽셀 글자, 픽셀 사이 좁은 틈(로고의 인물 격자처럼). 빨간 픽셀 하나는 글자 안에 둔다(로고의 가슴 픽셀).
   - 작은 크기(16 · 24px)는 틈 없이 · 여백 0 · 글자 크게 — 틈은 그 크기에서 뭉개진다.
   - 바꾼 뒤엔 파일 이름 버전을 올려야 Windows 아이콘 캐시가 새로 읽는다(_v6 → _v7, gui.py · gui_prep.py · 바로가기도 같이).
   v5(Pretendard Black 글자 + 우하단 빨간 점) 생성기는 make_icons_v5.py, 아이콘은 assets 의 *_v5.ico 로 남아 있다."""
import os
from PIL import Image, ImageDraw

INK, WHITE, ACC = (17, 17, 17, 255), (255, 255, 255, 255), (255, 77, 46, 255)
S = 1024

# 5 × 5 픽셀 글자. X = 흰 픽셀, A = 빨간 픽셀(로고의 포인트), . = 빈 칸
GLYPHS = {
    "R": ["XXXA.",
          "X...X",
          "XXXX.",
          "X..X.",
          "X...X"],
    "W": ["X...A",
          "X...X",
          "X.X.X",
          "X.X.X",
          ".X.X."],
}


def draw(ch, small):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    pad, rad = (0, 170) if small else (32, 200)
    d.rounded_rectangle((pad, pad, S - pad, S - pad), radius=rad, fill=INK)
    rows = GLYPHS[ch]; n = len(rows)
    span = S * (0.72 if small else 0.58)          # 글자 격자 한 변
    gap = 0 if small else span * 0.035            # 픽셀 사이 틈
    cell = (span - gap * (n - 1)) / n
    x0 = (S - span) / 2
    y0 = (S - span) / 2 - S * (0.0 if small else 0.01)
    for r, line in enumerate(rows):
        for c, v in enumerate(line):
            if v == ".":
                continue
            x, y = x0 + c * (cell + gap), y0 + r * (cell + gap)
            d.rectangle((round(x), round(y), round(x + cell) - 1, round(y + cell) - 1), fill=ACC if v == "A" else WHITE)
    return im


def save(ch, path):
    big, sm = draw(ch, False), draw(ch, True)
    imgs = {s: (sm if s <= 24 else big).resize((s, s), Image.LANCZOS) for s in (16, 24, 32, 48, 64, 128, 256)}
    imgs[256].save(path, format="ICO", sizes=[(s, s) for s in imgs], append_images=[imgs[s] for s in (16, 24, 32, 48, 64, 128)])
    return imgs, big


if __name__ == "__main__":
    HERE = os.path.dirname(os.path.abspath(__file__))
    (a, A) = save("W", r"C:\GMO_Cleansing\assets\icon_weekly_v6.ico")
    (b, B) = save("R", r"C:\Weekly ROAS\draft_tools\assets\icon_roas_v6.ico")
    sh = Image.new("RGB", (620, 330), (240, 240, 240)); d = ImageDraw.Draw(sh)
    for i, im in enumerate((a, b)):
        X = 10 + i * 300; sh.paste(im[256], (X, 10), im[256])
        x = X
        for s in (48, 32, 24, 16):
            sh.paste(im[s], (x, 320 - s), im[s]); x += s + 16
    sh.save(os.path.join(HERE, "pv.png"))
    print(os.path.join(HERE, "pv.png"))
