"""
The Avati logo, made ready for the tool.

The logo supplied was cropped after the T (AVAT). The full mark, AVATI, ends
in a plain vertical I in the gradient's deepest blue. It is completed here from
the T's own stem - the same width, the same height, the same blue - set one
letter-gap after the T's crossbar. Then the white ground is taken out as
alpha (colour-to-alpha, so the anti-aliased edges stay clean on a dark field)
and the result is trimmed and written at the sizes the build uses.
"""

from __future__ import annotations

import base64
import io
import os

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "design", "brand", "avati-supplied-avat.png")
OUT = os.path.join(ROOT, "design", "brand")

STEM_X0, STEM_X1 = 635, 671          # the T's stem, measured on the supplied image
CAP_Y0, CAP_Y1 = 9, 198
CROSSBAR_END = 727
GAP = 25                             # 0.13 of the cap height, as in the full mark


def _lerp(a, b, t):
    return tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def compose() -> Image.Image:
    src = Image.open(SRC).convert("RGB")
    w, h = src.size
    stem_w = STEM_X1 - STEM_X0 + 1
    x0 = CROSSBAR_END + GAP
    out = Image.new("RGB", (x0 + stem_w + 14, h), (255, 255, 255))
    out.paste(src, (0, 0))
    # The I: the T stem's shading, a shade deeper - it ends the gradient.
    top_l, top_r = (18, 108, 170), (14, 92, 156)
    bot_l, bot_r = (12, 96, 166), (10, 84, 150)
    for y in range(CAP_Y0, CAP_Y1 + 1):
        ty = (y - CAP_Y0) / (CAP_Y1 - CAP_Y0)
        left, right = _lerp(top_l, bot_l, ty), _lerp(top_r, bot_r, ty)
        for x in range(stem_w):
            out.putpixel((x0 + x, y), _lerp(left, right, x / (stem_w - 1)))
    return out


def white_to_alpha(im: Image.Image) -> Image.Image:
    """GIMP's colour-to-alpha against white: every pixel becomes the least
    transparent colour that, laid on white, gives back the original."""
    im = im.convert("RGB")
    out = Image.new("RGBA", im.size)
    px, po = im.load(), out.load()
    for y in range(im.size[1]):
        for x in range(im.size[0]):
            r, g, b = px[x, y]
            a = max(255 - r, 255 - g, 255 - b) / 255
            if a <= 0.02:
                po[x, y] = (0, 0, 0, 0)
                continue
            c = tuple(max(0, min(255, round((v - 255 * (1 - a)) / a))) for v in (r, g, b))
            po[x, y] = c + (round(a * 255),)
    return out


def trimmed(im: Image.Image, pad=2) -> Image.Image:
    box = im.getbbox()
    box = (max(0, box[0] - pad), max(0, box[1] - pad), min(im.width, box[2] + pad), min(im.height, box[3] + pad))
    return im.crop(box)


def build():
    logo = trimmed(white_to_alpha(compose()))
    os.makedirs(OUT, exist_ok=True)
    full = os.path.join(OUT, "avati-logo.png")
    logo.save(full, optimize=True)
    small = logo.resize((round(logo.width * 56 / logo.height), 56), Image.LANCZOS)
    small.save(os.path.join(OUT, "avati-logo-56.png"), optimize=True)
    return full, logo.size


def small_base64() -> str:
    """The small logo as base64, for the VBA to write out and place in every
    workbook it builds (a macro cannot read the images inside its own file)."""
    path = os.path.join(OUT, "avati-logo-56.png")
    if not os.path.exists(path):
        build()
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


if __name__ == "__main__":
    path, size = build()
    print(path, size, "base64 chars:", len(small_base64()))
