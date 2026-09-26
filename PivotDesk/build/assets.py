"""
The Desk's artwork: icons, the logo mark, and the hero panel.

Everything is authored as SVG here and rasterised by Chromium, so what ships
is exactly what was designed. Icons go into the workbook as SVG (Office 2016+
draws them crisp at any zoom) with a PNG fallback for older Office; the hero
is PNG only, because it uses blur and masks that Office's SVG renderer does
not implement.
"""

from __future__ import annotations

import math
import os
import subprocess

from design_tokens import EM, CANVAS, TX_1

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "design", "assets")
RENDER = os.path.join(HERE, "render.js")
PW = os.environ.get("PLAYWRIGHT_MODULE", "/opt/node22/lib/node_modules/playwright")

# ---------------------------------------------------------------------------
#  Icons: one hand - 24 grid, 1.75 stroke, round caps and joins.
# ---------------------------------------------------------------------------
ICONS = {
    "upload": '<path d="M12 15V4"/><path d="M7.5 8.5 12 4l4.5 4.5"/><path d="M4 15v3.5A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5V15"/>',
    "folder": '<path d="M3.5 7.5A1.5 1.5 0 0 1 5 6h4.2l1.8 2H19a1.5 1.5 0 0 1 1.5 1.5v8A1.5 1.5 0 0 1 19 19H5a1.5 1.5 0 0 1-1.5-1.5z"/>',
    "files": '<path d="M8 3.5h6.5L19 8v11a1.5 1.5 0 0 1-1.5 1.5h-9A1.5 1.5 0 0 1 7 19V5a1.5 1.5 0 0 1 1-1.5z"/><path d="M14 3.5V8h5"/><path d="M10 12.5h6M10 16h6"/>',
    "pivot": '<rect x="3.5" y="3.5" width="17" height="17" rx="2"/><path d="M3.5 9h17M9 9v11.5"/><path d="M13 13.5h4.5M13 16.8h3"/>',
    "balance": '<path d="M12 4v16"/><path d="M6 7h12"/><path d="M8 20h8"/><path d="M6 7l-3 6.5a3 3 0 0 0 6 0z"/><path d="M18 7l-3 6.5a3 3 0 0 0 6 0z"/>',
    "pulse": '<path d="M3 12h4l2.5-6 5 12 2.5-6h4"/>',
    "check": '<path d="M5 12.5l4.2 4.2L19 7"/>',
    "arrow": '<path d="M5 12h14"/><path d="M13.5 6.5 19 12l-5.5 5.5"/>',
    "reset": '<path d="M4 12a8 8 0 1 0 2.4-5.7"/><path d="M4 4.5V9h4.5"/>',
    "expand": '<path d="M14 4h6v6"/><path d="M10 20H4v-6"/><path d="M20 4l-6.5 6.5M4 20l6.5-6.5"/>',
    "info": '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5.5"/><path d="M12 7.6v.2"/>',
    "warn": '<path d="M12 4 21 19.5H3z"/><path d="M12 10v4.5"/><path d="M12 17.2v.2"/>',
    "error": '<circle cx="12" cy="12" r="8.5"/><path d="M9 9l6 6M15 9l-6 6"/>',
    "clock": '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
    "database": '<ellipse cx="12" cy="6" rx="7" ry="2.8"/><path d="M5 6v12c0 1.5 3.1 2.8 7 2.8s7-1.3 7-2.8V6"/><path d="M5 12c0 1.5 3.1 2.8 7 2.8s7-1.3 7-2.8"/>',
    "calendar": '<rect x="4" y="5.5" width="16" height="14.5" rx="2"/><path d="M4 10h16M8.5 3.5v4M15.5 3.5v4"/>',
    "spark": '<path d="M12 3.5l1.9 5.1 5.1 1.9-5.1 1.9L12 17.5l-1.9-5.1L5 10.5l5.1-1.9z"/><path d="M18.5 16.5l.7 1.8 1.8.7-1.8.7-.7 1.8-.7-1.8-1.8-.7 1.8-.7z"/>',
    "chevron": '<path d="M9.5 6 15.5 12l-6 6"/>',
    "external": '<path d="M14 4.5h5.5V10"/><path d="M19.5 4.5 11 13"/><path d="M18 14v4.5a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 4 18.5v-11A1.5 1.5 0 0 1 5.5 6H10"/>',
    "layers": '<path d="M12 3.5 20.5 8 12 12.5 3.5 8z"/><path d="M3.5 12 12 16.5 20.5 12"/><path d="M3.5 16 12 20.5 20.5 16"/>',
}


def icon_svg(name: str, color: str, stroke: float = 1.75) -> str:
    return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24" '
            'fill="none" stroke="#%s" stroke-width="%s" stroke-linecap="round" '
            'stroke-linejoin="round">%s</svg>') % (color, stroke, ICONS[name])


# ---------------------------------------------------------------------------
#  The khatam: an eight-point star, two squares turned 45 degrees apart.
# ---------------------------------------------------------------------------
def khatam_points(cx, cy, r):
    pts = []
    # an {8/2} star as a single outline: alternate outer tips and inner notches
    r_in = r * math.cos(math.radians(45)) / math.cos(math.radians(22.5))
    for k in range(16):
        ang = math.radians(-90 + k * 22.5)
        rr = r if k % 2 == 0 else r_in
        pts.append((cx + rr * math.cos(ang), cy + rr * math.sin(ang)))
    return pts


def logo_svg(size=64) -> str:
    s = size
    pts = khatam_points(s / 2, s / 2, s * 0.30)
    poly = " ".join("%.2f,%.2f" % p for p in pts)
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {s} {s}" width="{s}" height="{s}">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#{EM[300]}"/><stop offset=".55" stop-color="#{EM[500]}"/><stop offset="1" stop-color="#{EM[700]}"/>
    </linearGradient>
    <linearGradient id="sheen" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#fff" stop-opacity=".28"/><stop offset=".5" stop-color="#fff" stop-opacity="0"/>
    </linearGradient>
  </defs>
  <rect x="0" y="0" width="{s}" height="{s}" rx="{s*0.26:.1f}" fill="url(#g)"/>
  <rect x="0" y="0" width="{s}" height="{s}" rx="{s*0.26:.1f}" fill="url(#sheen)"/>
  <polygon points="{poly}" fill="none" stroke="#FFFFFF" stroke-width="{s*0.055:.2f}" stroke-linejoin="round"/>
  <circle cx="{s/2}" cy="{s/2}" r="{s*0.075:.2f}" fill="#FFFFFF"/>
</svg>'''


# ---------------------------------------------------------------------------
#  The lattice: stars at every grid point whose axial tips touch, so the
#  gaps between four stars read as crosses - the star-and-cross of Cairo's
#  mashrabiya and tilework.
# ---------------------------------------------------------------------------
def lattice_paths(w, h, step):
    r = step / 2.0
    out = []
    y = -step
    row = 0
    while y < h + step:
        x = -step
        while x < w + step:
            pts = khatam_points(x, y, r)
            out.append("M" + " L".join("%.1f %.1f" % p for p in pts) + " Z")
            # a smaller star nested inside, for the interlaced look
            pts2 = khatam_points(x, y, r * 0.52)
            out.append("M" + " L".join("%.1f %.1f" % p for p in pts2) + " Z")
            x += step
        y += step
        row += 1
    return " ".join(out)


def hero_html(w_pt, h_pt, radius_pt) -> str:
    """The hero panel as a page, sized in CSS px (1 pt = 4/3 px)."""
    W = round(w_pt * 4 / 3)
    H = round(h_pt * 4 / 3)
    R = radius_pt * 4 / 3
    lat = lattice_paths(W, H, 64)
    return f'''<!doctype html><html><head><meta charset="utf-8"><style>
html,body{{margin:0;background:transparent;}}
.p{{position:relative;width:{W}px;height:{H}px;border-radius:{R}px;overflow:hidden;
   background:linear-gradient(100deg,#050C09 0%,#061410 38%,#072217 70%,#08301F 100%);}}
.a1{{position:absolute;left:{W*0.52}px;top:-{H*0.9}px;width:{W*0.62}px;height:{H*2.1}px;border-radius:50%;
   background:radial-gradient(closest-side,rgba(0,144,96,.55),rgba(0,144,96,0));filter:blur(18px);}}
.a2{{position:absolute;left:{W*0.80}px;top:{H*0.35}px;width:{W*0.36}px;height:{H*1.4}px;border-radius:50%;
   background:radial-gradient(closest-side,rgba(22,176,127,.40),rgba(22,176,127,0));filter:blur(14px);}}
.a3{{position:absolute;left:-{W*0.10}px;top:-{H*0.6}px;width:{W*0.45}px;height:{H*1.4}px;border-radius:50%;
   background:radial-gradient(closest-side,rgba(0,97,65,.30),rgba(0,97,65,0));filter:blur(20px);}}
svg.l{{position:absolute;left:0;top:0;
   -webkit-mask-image:linear-gradient(90deg,transparent 0%,transparent 34%,rgba(0,0,0,.55) 58%,#000 82%);
   mask-image:linear-gradient(90deg,transparent 0%,transparent 34%,rgba(0,0,0,.55) 58%,#000 82%);}}
.g{{position:absolute;inset:0;opacity:.07;mix-blend-mode:overlay;}}
.shine{{position:absolute;left:0;right:0;top:0;height:1px;
   background:linear-gradient(90deg,rgba(255,255,255,0),rgba(255,255,255,.14) 40%,rgba(143,219,190,.22) 75%,rgba(255,255,255,0));}}
.edge{{position:absolute;inset:0;border-radius:{R}px;box-shadow:inset 0 0 0 1px rgba(79,199,156,.16);}}
</style></head><body><div class="p">
<div class="a3"></div><div class="a1"></div><div class="a2"></div>
<svg class="l" width="{W}" height="{H}" viewBox="0 0 {W} {H}">
  <path d="{lat}" fill="none" stroke="rgba(79,199,156,.26)" stroke-width="1"/>
</svg>
<svg class="g" width="{W}" height="{H}"><filter id="n"><feTurbulence type="fractalNoise" baseFrequency=".9" numOctaves="2" stitchTiles="stitch"/></filter>
  <rect width="100%" height="100%" filter="url(#n)"/></svg>
<div class="shine"></div><div class="edge"></div>
</div></body></html>'''


def empty_art_html(w_pt, h_pt) -> str:
    """A quiet lattice medallion for empty states."""
    W = round(w_pt * 4 / 3)
    H = round(h_pt * 4 / 3)
    lat = lattice_paths(W, H, 36)
    return f'''<!doctype html><html><head><meta charset="utf-8"><style>
html,body{{margin:0;background:transparent;}}
svg{{-webkit-mask-image:radial-gradient(closest-side,#000 30%,transparent 100%);}}
</style></head><body>
<svg width="{W}" height="{H}" viewBox="0 0 {W} {H}"><path d="{lat}" fill="none" stroke="rgba(79,199,156,.30)" stroke-width="1"/></svg>
</body></html>'''


def render(html: str, out_png: str, w_px: int, h_px: int, scale=2, transparent=True):
    tmp = out_png + ".html"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html)
    env = dict(os.environ, PLAYWRIGHT_MODULE=PW)
    subprocess.run(["node", RENDER, tmp, out_png, str(w_px), str(h_px), str(scale),
                    "1" if transparent else "0"], check=True, env=env)
    os.remove(tmp)


def build_icons(colors: dict[str, str]):
    """colors: variant suffix -> hex. Writes <name>-<variant>.svg/.png."""
    os.makedirs(OUT, exist_ok=True)
    made = {}
    for name in ICONS:
        for variant, col in colors.items():
            base = os.path.join(OUT, "icon-%s-%s" % (name, variant))
            svg = icon_svg(name, col)
            with open(base + ".svg", "w", encoding="utf-8") as f:
                f.write(svg)
            html = ('<!doctype html><html><body style="margin:0;background:transparent">'
                    '<div style="width:24px;height:24px">%s</div></body></html>') % svg
            render(html, base + ".png", 24, 24, scale=4)
            made[(name, variant)] = base
    return made


def build_all(hero_w, hero_h, hero_r):
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, "logo.svg"), "w", encoding="utf-8") as f:
        f.write(logo_svg())
    render('<!doctype html><html><body style="margin:0;background:transparent">%s</body></html>' % logo_svg(),
           os.path.join(OUT, "logo.png"), 64, 64, scale=2)
    render(hero_html(hero_w, hero_h, hero_r), os.path.join(OUT, "hero.png"),
           round(hero_w * 4 / 3), round(hero_h * 4 / 3), scale=2)
    render(empty_art_html(240, 78), os.path.join(OUT, "lattice-medallion.png"),
           round(240 * 4 / 3), round(78 * 4 / 3), scale=2)


if __name__ == "__main__":
    build_all(1064, 152, 16)
    build_icons({"em": EM[300], "white": "FFFFFF", "ink": "0C1512", "muted": "7E9388", "dark": "18241F"})
