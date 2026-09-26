"""
Contrast, measured rather than eyeballed.

For every run of text on the Desk, find what is actually behind it - the
shapes beneath it in drawing order, composited with their transparency over
the canvas, and for the hero the rendered artwork itself, sampled at the
text's position - and compute the WCAG contrast ratio.

  AA body text   4.5 : 1
  AA large text  3.0 : 1   (18 pt and up, or 14 pt semibold and up)

Text that is deliberately quiet - a disabled control, a placeholder - is
exempt under WCAG 1.4.3 and is reported as such, not hidden.
"""

from __future__ import annotations

import os
import sys

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import desk  # noqa: E402
from design_tokens import CANVAS, TX_4  # noqa: E402
from shapes import Linear, Solid  # noqa: E402

DISABLED = {TX_4.upper(), "4A5E55", "3A4B43"}


def rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def lum(c):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = c
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def ratio(a, b):
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def over(top, alpha, bottom):
    return tuple(round(t * alpha + b * (1 - alpha)) for t, b in zip(top, bottom))


def fill_at(fill, fy):
    if isinstance(fill, Solid):
        return rgb(fill.color), fill.alpha
    if isinstance(fill, Linear):
        # vertical gradients vary with y; horizontal ones - take the middle
        stops = fill.stops
        t = fy if 45 <= fill.angle % 180 <= 135 else 0.5
        for (p0, c0, a0), (p1, c1, a1) in zip(stops, stops[1:]):
            if p0 <= t <= p1:
                k = 0 if p1 == p0 else (t - p0) / (p1 - p0)
                c = tuple(round(x + (y - x) * k) for x, y in zip(rgb(c0), rgb(c1)))
                return c, a0 + (a1 - a0) * k
        return rgb(stops[-1][1]), stops[-1][2]
    return None, 0


def background(shapes, idx, x, y, images):
    """Composite everything below shapes[idx] at (x, y)."""
    col = rgb(CANVAS)
    for s in shapes[:idx]:
        if s.hidden:
            continue
        if not (s.x <= x <= s.x + s.w and s.y <= y <= s.y + s.h):
            continue
        if s.image:
            img = images.get(s.image)
            if img is None:
                path = s.image + ".png"
                img = Image.open(path).convert("RGBA")
                images[s.image] = img
            px = int((x - s.x) / s.w * (img.width - 1))
            py = int((y - s.y) / s.h * (img.height - 1))
            r, g, b, a = img.getpixel((px, py))
            col = over((r, g, b), a / 255 * s.image_alpha, col)
            continue
        c, a = fill_at(s.fill, (y - s.y) / max(s.h, 1e-6))
        if c is not None and a > 0:
            col = over(c, a, col)
    return col


def check(state, label):
    shapes = desk.desk(state)
    images = {}
    rows = []
    for i, s in enumerate(shapes):
        if s.hidden or not s.body:
            continue
        cx, cy = s.x + min(s.w, 60) / 2, s.y + s.h / 2
        # a shape with its own fill is its own background
        own, a = fill_at(s.fill, 0.5) if s.fill is not None else (None, 0)
        bg = background(shapes, i, cx, cy, images)
        if own is not None and a > 0:
            bg = over(own, a, bg)
        for p in s.body.paras:
            for r in p.runs:
                if not r.text.strip():
                    continue
                fg = over(rgb(r.color), r.alpha, bg)
                large = r.size >= 18 or (r.size >= 14 and "Semibold" in r.font)
                need = 3.0 if large else 4.5
                cr = ratio(fg, bg)
                exempt = r.color.upper() in DISABLED
                rows.append((cr, need, exempt, s.name, r.text[:40], r.color, "%02X%02X%02X" % bg, r.size))
    fails = [r for r in rows if r[0] < r[1] and not r[2]]
    exempt = [r for r in rows if r[0] < r[1] and r[2]]
    lines = ["## %s" % label, "",
             "%d text runs checked, %d below AA, %d below AA but exempt (disabled / placeholder)." % (
                 len(rows), len(fails), len(exempt)), "",
             "| ratio | needs | shape | text | fg | bg | pt |", "|---:|---:|---|---|---|---|---:|"]
    for cr, need, ex, name, txt, fg, bg, sz in sorted(rows):
        flag = " (exempt)" if ex and cr < need else (" **FAIL**" if cr < need else "")
        lines.append("| %.2f%s | %.1f | %s | %s | #%s | #%s | %.1f |" % (cr, flag, need, name, txt.replace("|", "/"),
                                                                           fg, bg, sz))
    return fails, "\n".join(lines)


# The table sheets and the built workbooks: every text / background pair
# modPD_Theme and modPD_Pivot put together, with the size it is set at.
SHEET_PAIRS = [
    # every tool sheet: modPD_Theme.Dress, Rail, Toolbar, Head, DressTable, SetStatus
    ("overline over the title", "4FC79C", "000000", 7.5),
    ("title on the black masthead", "F2F7F4", "000000", 19),
    ("description under the title", "A9BDB3", "000000", 9),
    ("app-bar ALM DESK", "F2F7F4", "000000", 7.5),
    ("app-bar bank name", "7E9388", "000000", 6),
    ("nav: other sheets, in the well", "A9BDB3", "0B1310", 8.5),
    ("nav: this sheet", "F2F7F4", "003323", 8.5),
    ("action pills", "C2EBDA", "001D14", 8.5),
    ("band over the columns", "4FC79C", "001D14", 7.5),
    ("table header", "C2EBDA", "000000", 8.5),
    ("body on a row", "F2F7F4", "0C1512", 9.5),
    ("body on the other row", "F2F7F4", "111F19", 9.5),
    ("muted on a row", "7E9388", "0C1512", 9.5),
    ("muted on the other row", "7E9388", "111F19", 9.5),
    ("What to fix on a row", "A9BDB3", "111F19", 9.5),
    ("status OK", "2FC48D", "0D2A20", 9),
    ("status Check", "F2B544", "2A2310", 9),
    ("status Break", "FF6B5E", "2E1614", 9),
    ("status Idle", "7E9388", "121D19", 9),
    # the gallery's cards
    ("card kind", "4FC79C", "0C1512", 7),
    ("card title", "F2F7F4", "0C1512", 12.5),
    ("card description", "A9BDB3", "0C1512", 9),
    ("card ingredients", "7E9388", "0C1512", 7.5),
    ("card IN THE NEXT BUILD", "16B07F", "0C1512", 7),
    # a built workbook: Start here, the pivots, the charts
    ("guide links", "4FC79C", "0C1512", 9.5),
    ("guide section label", "4FC79C", "060B09", 8),
    ("tile label", "7E9388", "0C1512", 6.5),
    ("tile figure", "F2F7F4", "0C1512", 17),
    ("pivot header", "C2EBDA", "000000", 10),
    ("pivot grand total", "F2F7F4", "001D14", 10),
    ("pivot filter labels", "4FC79C", "111F19", 10),
    ("negative figure on a row", "FF0000", "0C1512", 10),
    ("figure on a data bar", "F2F7F4", "00794F", 10),
    ("figure on the heatmap's top", "F2F7F4", "00794F", 10),
    ("figure in a top-N cell", "F2F7F4", "003323", 10),
    ("negative under Highlight=Negatives", "FF6B5E", "2E1614", 10),
    ("slicer: chosen item", "F2F7F4", "006141", 8.5),
    ("slicer: other item", "A9BDB3", "111F19", 8.5),
    ("chart title", "F2F7F4", "0C1512", 11),
    ("chart axis figures", "7E9388", "0C1512", 8.5),
    ("chart category labels and legend", "A9BDB3", "0C1512", 8.5),
    ("chart data labels", "F2F7F4", "0C1512", 8),
    # graphics, not text: WCAG 1.4.11 asks 3:1 against what they sit on
    ("data bar (graphic)", "00794F", "0C1512", None),
    ("data bar on the other row (graphic)", "00794F", "111F19", None),
    ("tile edge (graphic)", "009060", "0C1512", None),
    ("card edge when on (graphic)", "009060", "060B09", None),
]


def check_sheets():
    lines = ["## Table sheets and built workbooks", "",
             "| ratio | needs | where | fg | bg | pt |", "|---:|---:|---|---|---|---:|"]
    fails = []
    for where, fg, bg, pt in SHEET_PAIRS:
        need = 3.0 if pt is None or pt >= 18 else 4.5
        cr = ratio(rgb(fg), rgb(bg))
        if cr < need:
            fails.append((cr, need, where, fg, bg))
        lines.append("| %.2f%s | %.1f | %s | #%s | #%s | %s |" % (cr, " **FAIL**" if cr < need else "",
                                                                 need, where, fg, bg,
                                                                 "graphic" if pt is None else "%.1f" % pt))
    return fails, "\n".join(lines)


def main():
    out = ["# Desk contrast report", "",
           "Generated by build/contrast.py. WCAG 2.1 AA: 4.5:1 for body text, 3:1 for large text.", ""]
    total = 0
    for label, st in (("Shipped state", desk.empty_state()), ("Working state", desk.showcase_state())):
        st = dict(st)
        st["macros_banner"] = label == "Shipped state"
        fails, md = check(st, label)
        total += len(fails)
        out += [md, ""]
        for f in fails:
            print("FAIL %.2f < %.1f  %s  %r  #%s on #%s" % (f[0], f[1], f[3], f[4], f[5], f[6]))
    fails, md = check_sheets()
    total += len(fails)
    out += [md, ""]
    for f in fails:
        print("FAIL %.2f < %.1f  %s  #%s on #%s" % f)
    path = os.path.join(os.path.dirname(HERE), "docs", "CONTRAST.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print("%d failure(s); report in docs/CONTRAST.md" % total)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
