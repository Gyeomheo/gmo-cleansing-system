"""바로가기 아이콘 생성 (W. / R.) — 글자: 광학 중앙(면적 기준 크기, 무게중심 보정, 2% 위) · 점: 타일 우하단 대각선, 둥근 모서리와 등간격.
   바꾼 뒤엔 파일 이름 버전을 올려야 Windows 아이콘 캐시가 새로 읽는다(_v5 → _v6, gui.py · 바로가기도 같이)."""
import math, os
import numpy as np
from PIL import Image, ImageDraw, ImageFont
PF = r"C:\Users\GyeomHeo\AppData\Local\Microsoft\Windows\Fonts\Pretendard-Black.ttf"
INK, WHITE, ACC = (17,17,17,255), (255,255,255,255), (255,77,46,255)
S = 1024

def glyph(ch, size):
    f = ImageFont.truetype(PF, size)
    m = Image.new("L", (S*2, S*2), 0); ImageDraw.Draw(m).text((S//2, S//2), ch, font=f, fill=255)
    a = np.array(m); ys, xs = np.where(a > 127)
    return m.crop((xs.min(), ys.min(), xs.max()+1, ys.max()+1))

def draw(ch, small):
    im = Image.new("RGBA", (S,S), (0,0,0,0)); d = ImageDraw.Draw(im)
    pad, rad = (0, 170) if small else (32, 200)
    d.rounded_rectangle((pad,pad,S-pad,S-pad), radius=rad, fill=INK)
    # 크기: 잉크 면적을 같게 (W 와 R 이 같은 무게로 보이게), 너무 넓으면 폭 제한
    area = (S * (0.44 if small else 0.33)) ** 2          # 목표 잉크 면적
    g = glyph(ch, 400); ink = (np.array(g) > 127).sum()
    size = int(400 * math.sqrt(area / ink))
    g = glyph(ch, size)
    maxw = S * (0.82 if small else 0.62)
    if g.width > maxw: g = glyph(ch, int(size * maxw / g.width))
    a = np.array(g) > 127; gw, gh = g.size
    cx_ink = np.where(a)[1].mean()                         # 무게중심 x
    ox = 0.5 * (gw / 2) + 0.5 * cx_ink                      # 박스 중심과 무게중심 절반씩
    x0 = round(S/2 - ox)
    y0 = round(S/2 - gh/2 - S*0.02)                         # 광학 중심: 2% 위
    im.paste(Image.new("RGBA", g.size, WHITE), (x0, y0), g)
    # 점: 모서리 원 중심 → 대각선 방향 둥근 변 위 점에서 안쪽으로 inset
    dot = int(S * (0.15 if small else 0.10))
    inset = S * (0.075 if small else 0.085)
    c = S - pad - rad
    arc = c + rad / math.sqrt(2)                            # 대각선과 둥근 변이 만나는 점
    corner = arc - inset / math.sqrt(2) * 1.0 - inset * 0.5  # 점의 바깥 모서리 (대각선 위)
    d.rectangle((corner - dot, corner - dot, corner, corner), fill=ACC)
    # 겹침 검사
    gx1, gy1 = x0 + gw, y0 + gh
    if not small and corner - dot < gx1 and corner - dot < gy1:
        sub = a[max(0, int(corner-dot-y0)):, max(0, int(corner-dot-x0)):]
        if sub.any(): print("⚠ 겹침", ch)
    return im

def save(ch, path):
    big, sm = draw(ch, False), draw(ch, True)
    imgs = {s: (sm if s <= 24 else big).resize((s,s), Image.LANCZOS) for s in (16,24,32,48,64,128,256)}
    imgs[256].save(path, format="ICO", sizes=[(s,s) for s in imgs], append_images=[imgs[s] for s in (16,24,32,48,64,128)])
    return imgs, big

HERE = os.path.dirname(os.path.abspath(__file__))
(a, A) = save("W", r"C:\GMO_Cleansing\assets\icon_weekly_v5.ico")
(b, B) = save("R", r"C:\Weekly ROAS\draft_tools\assets\icon_roas_v5.ico")
sh = Image.new("RGB", (620, 320), (240,240,240)); d = ImageDraw.Draw(sh)
for i, im in enumerate((a, b)):
    X = 10 + i*300; sh.paste(im[256], (X, 10), im[256])
    d.line((X+128, 10, X+128, 266), fill=(90,90,90)); d.line((X, 138, X+256, 138), fill=(90,90,90))   # 중심선
    x = X
    for s in (48,32,24,16): sh.paste(im[s], (x, 310 - s), im[s]); x += s + 16
sh.save(os.path.join(HERE, "pv.png"))
