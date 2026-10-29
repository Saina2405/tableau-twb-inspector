"""
Draws the TWB Inspector icon (a small lineage flow: table -> calculations -> worksheet)
and writes:
  twb_inspector.ico         Windows icon with all sizes (used for the .exe and the window)
  twb_inspector.png         256 px PNG (window icon on other systems, README)
  twb_inspector_header.png  96 px version with a light outline for dark backgrounds
  favicon.png               64 px PNG embedded in the HTML lineage diagram

Run:  python assets/make_icon.py      (needs Pillow: pip install pillow)
"""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
S = 1024
NAVY = "#14213D"
PINK, BLUE, ORANGE, TEAL, LINE = "#F7A8C4", "#8AB4F8", "#F6AD55", "#4FD1C5", "#A9B8D6"


def bez(p0, p1, p2, p3, n):
    for i in range(n + 1):
        t = i / n
        yield ((1-t)**3*p0[0] + 3*(1-t)**2*t*p1[0] + 3*(1-t)*t**2*p2[0] + t**3*p3[0],
               (1-t)**3*p0[1] + 3*(1-t)**2*t*p1[1] + 3*(1-t)*t**2*p2[1] + t**3*p3[1])


def curve(d, a, b, w):
    mx = (a[0] + b[0]) / 2
    r = w / 2
    for x, y in bez(a, (mx, a[1]), (mx, b[1]), b, 600):
        d.ellipse((x - r, y - r, x + r, y + r), fill=LINE)


def node(d, c, r, fill):
    x, y = c
    d.rounded_rectangle((x - r*1.25, y - r*0.8, x + r*1.25, y + r*0.8), radius=r*0.35, fill=fill)


def draw(outline=None, bold=False):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    if outline:
        d.rounded_rectangle((8, 8, S - 8, S - 8), radius=226, fill=outline)
    d.rounded_rectangle((24, 24, S - 24, S - 24), radius=210, fill=NAVY)
    k = 1.25 if bold else 1.0           # thicker shapes for the tiny sizes
    L, M1, M2, R = (215, 512), (512, 300), (512, 724), (809, 512)
    for a, b in [(L, M1), (L, M2), (M1, R), (M2, R)]:
        curve(d, a, b, 34 * k * (1.3 if bold else 1))
    node(d, L, 74 * k, PINK); node(d, M1, 66 * k, ORANGE); node(d, M2, 66 * k, BLUE); node(d, R, 74 * k, TEAL)
    return im


if __name__ == "__main__":
    big, small = draw(), draw(bold=True)
    sizes = [256, 128, 64, 48, 32, 24, 16]
    frames = [(small if s <= 32 else big).resize((s, s), Image.LANCZOS) for s in sizes]
    frames[0].save(os.path.join(HERE, "twb_inspector.ico"), format="ICO",
                   sizes=[(s, s) for s in sizes], append_images=frames[1:])
    big.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, "twb_inspector.png"))
    draw(outline="#3A4C78").resize((96, 96), Image.LANCZOS).save(os.path.join(HERE, "twb_inspector_header.png"))
    big.resize((64, 64), Image.LANCZOS).save(os.path.join(HERE, "favicon.png"))
    print("Icons written to", HERE)
