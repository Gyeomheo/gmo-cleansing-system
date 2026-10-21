"""바로가기 아이콘 생성 (W / R) — v8 (2026-10-21 담당자 선택, 시안 'PV3 왼쪽'):
   v7 의 손으로 그린 픽셀 글꼴(7줄 · 굵은 획 2칸 · 픽셀 사이 좁은 틈) + v6 처럼 빨간 점을 따로 두지 않고 글자의 한 픽셀을 빨강으로.
   - 빨간 픽셀(A): R = 윗줄 끝, W = 오른쪽 위 끝. 글자는 타일 가운데(1% 위), 높이 = 타일의 50%.
   - 작은 그림(16 · 24px)은 틈 없이 · 여백 0 · 글자 66% — 틈은 그 크기에서 뭉개진다.
   - 바꾼 뒤엔 파일 이름 버전을 올려야 Windows 아이콘 캐시가 새로 읽는다(_v8 → _v9, gui.py · gui_prep.py · 바로가기도 같이).
   이전 생성기: make_icons_v5.py(Pretendard 글자 + 우하단 점) · make_icons_v6.py(5 × 5 격자) · make_icons_v7.py(v5 배치 + 픽셀 글꼴) — v6 · v7 반려."""
import os
from PIL import Image, ImageDraw

INK, WHITE, ACC = (17, 17, 17, 255), (255, 255, 255, 255), (255, 77, 46, 255)
S = 1024

GLYPHS = {  # X = 흰 픽셀, A = 빨간 픽셀, . = 빈 칸
    "R": ["XXXXA.",
          "XX..XX",
          "XX..XX",
          "XXXXX.",
          "XX.XX.",
          "XX..XX",
          "XX..XX"],
    "W": ["XX...XA",
          "XX...XX",
          "XX.X.XX",
          "XX.X.XX",
          "XXXXXXX",
          ".XX.XX.",
          ".XX.XX."],
}


def draw(ch, small):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    pad, rad = (0, 170) if small else (32, 200)
    d.rounded_rectangle((pad, pad, S - pad, S - pad), radius=rad, fill=INK)
    rows = GLYPHS[ch]; ny, nx = len(rows), len(rows[0])
    gh = S * (0.66 if small else 0.50); cell = gh / ny
    gap = 0 if small else cell * 0.10
    x0, y0 = S / 2 - nx * cell / 2, S / 2 - gh / 2 - S * 0.01
    for r, line in enumerate(rows):
        for c, v in enumerate(line):
            if v == ".":
                continue
            x, y = x0 + c * cell, y0 + r * cell
            d.rectangle((round(x + gap / 2), round(y + gap / 2), round(x + cell - gap / 2) - 1, round(y + cell - gap / 2) - 1),
                        fill=ACC if v == "A" else WHITE)
    return im


def save(ch, path):
    big, sm = draw(ch, False), draw(ch, True)
    imgs = {s: (sm if s <= 24 else big).resize((s, s), Image.LANCZOS) for s in (16, 24, 32, 48, 64, 128, 256)}
    imgs[256].save(path, format="ICO", sizes=[(s, s) for s in imgs], append_images=[imgs[s] for s in (16, 24, 32, 48, 64, 128)])
    return imgs


if __name__ == "__main__":
    HERE = os.path.dirname(os.path.abspath(__file__))
    a = save("W", r"C:\GMO_Cleansing\assets\icon_weekly_v8.ico")
    b = save("R", r"C:\Weekly ROAS\draft_tools\assets\icon_roas_v8.ico")
    sh = Image.new("RGB", (620, 330), (240, 240, 240))
    for i, im in enumerate((a, b)):
        X = 10 + i * 300; sh.paste(im[256], (X, 10), im[256]); x = X
        for s in (48, 32, 24, 16):
            sh.paste(im[s], (x, 320 - s), im[s]); x += s + 16
    sh.save(os.path.join(HERE, "pv.png"))
    print(os.path.join(HERE, "pv.png"))
