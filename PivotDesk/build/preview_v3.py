"""
Previews of what Avati 3.0 adds, drawn with the rules the VBA applies:

  sheet-chart-config     Chart config, its defaults read out of modPD_Charts
  sheet-workbooks        Workbooks, its defaults read out of modPD_Books
  sheet-gallery          the Gallery's cards, read out of modPD_Gallery
  built-start-here       Start here of an LCR build, with the default charts
  built-top-counterparties  the top counterparties pivot of that build

The built-workbook pictures are computed from the LCR sample the bank
supplied (LCR_SAMPLE), grouped the way the staging pass and the pivots
group it. They are pictures for review off Windows, not screenshots.
"""

from __future__ import annotations

import collections
import html
import math
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import assets  # noqa: E402
import labels  # noqa: E402
import preview_sheets as P  # noqa: E402

C = P.C
px = P.px
SAMPLE = os.environ.get("LCR_SAMPLE", "")
SURFACE = "0C1512"
PALETTES = {"avati": ["16B07F", "4FA3D9", "E0A43B", "8FDBBE", "E07A6B", "A58BE0", "C9B98A", "7E9388"],
            "emerald": ["16B07F", "8FDBBE", "009060", "4FC79C", "C2EBDA", "2E9E78"],
            "two tone": ["16B07F", "4FA3D9"]}


# ---------------------------------------------------------------------------
#  reading the VBA
# ---------------------------------------------------------------------------

def _joined(module):
    return re.sub(r" _\n\s*", " ", P._vba(module))


def _consts():
    k = P._consts()
    k.update({"MAX_BOOKS_DEFAULT": "40", "MAX_SHEETS_DEFAULT": "120"})
    return k


def _args(text):
    text = text.replace("ChrW(183)", '"·"').replace("ChrW(9679)", '"●"')
    text = re.sub(r"FwLabel\(FW_LCR\)", '"LCR"', text)
    text = re.sub(r"FwLabel\(FW_NSFR\)", '"NSFR"', text)
    text = re.sub(r"FwLabel\(FW_ML\)", '"Maturity Ladder"', text)
    return P._args(text, _consts())


def heads_of(module, fn):
    src = _joined(module)
    return re.findall(r'"([^"]*)"', re.search(r"%s = Array\((.*?)\)\s*\n" % fn, src).group(1))


def widths_of(module, fn):
    src = _joined(module)
    return [int(w) for w in re.search(r"%s = Array\(([^)]*)\)" % fn, src).group(1).split(",")]


def bands_of(module, prefix):
    src = _joined(module)
    k = {n: int(v) for n, v in re.findall(r"(?:Private|Public) Const (%s\w+) As Long = (\d+)" % prefix, src)}
    return [(k[a], k[b], lab) for a, b, lab in
            re.findall(r'Band ws, (%s\w+), (%s\w+), "([^"]+)"' % (prefix, prefix), src)]


def rows_by_heading(module, call, heads):
    src = _joined(module)
    out = []
    for m in re.finditer(r"^\s*%s ws, r(?: \+ \d)?, (.*)$" % call, src, re.M):
        d = {}
        for kv in _args(m.group(1)):
            key, _, val = kv.partition("=")
            d[key.lower()] = val
        out.append([d.get(h.lower(), "") for h in heads])
    return out


def templates():
    src = _joined("modPD_Gallery")
    src = src[src.index("Private Function Templates"):src.index("Set Templates = c")]
    out = []
    for chunk in src.split("c.Add Array(")[1:]:
        head = chunk[:chunk.index(", Array(")]
        key, kind, title, desc, uses = _args(head)[:5]
        spec = chunk[chunk.index(", Array(") + 8:]
        name = re.search(r'"(?:Pivot|Chart)=([^"]*)"', spec).group(1)
        out.append(dict(key=key, kind=kind, title=title, desc=desc, uses=uses, name=name))
    return out


# ---------------------------------------------------------------------------
#  the tool sheets
# ---------------------------------------------------------------------------

def chart_config_sheet():
    heads = heads_of("modPD_Charts", "ChartHeads")
    widths = widths_of("modPD_Charts", "ChartWidths")
    rows = []
    for r in rows_by_heading("modPD_Charts", "ChartRec", heads):
        rows.append(r[:-2] + (["OK", ""] if r[0] == "Yes" else ["Off", ""]))
    n_on = sum(r[0] == "Yes" for r in rows)
    rows += [[""] * len(heads) for _ in range(3)]
    t = P.table(heads, widths, rows, verdict_col=len(heads) - 2, muted_cols=(len(heads) - 3, len(heads) - 1),
                semi_cols=(1,), pre_rows=P.band_row(widths, bands_of("modPD_Charts", "X_")))
    body = (P.chrome("reports", "Chart config", "What a build draws. One row is one chart, read like a pivot: the "
                     "same fields, values and filters. It goes on Start here or on a sheet of its own, and redraws "
                     "whenever the book refreshes.", "%d chart(s) switched on, every one ready to draw." % n_on, "OK",
                     ["Add a chart", "Check", "Restore defaults"], tab="Charts",
                     overline="REPORTS &nbsp;&#183;&nbsp; CHARTS", status_w=1400) + t)
    return P.page("Chart config", body, 3400)


def workbooks_sheet():
    heads = heads_of("modPD_Books", "BookHeads")
    widths = widths_of("modPD_Books", "BookWidths")
    src = _joined("modPD_Books")
    body_src = src[src.index("Private Sub WriteDefaultBooks"):src.index("Private Sub DressBooks")]
    rows = []
    for m in re.finditer(r"\.Value2 = Array\((.*?)\)\s*$", body_src, re.M):
        vals = _args(m.group(1))
        rows.append(vals + ["OK", ""])
    t = P.table(heads, widths, rows, verdict_col=len(heads) - 2, muted_cols=(5, 7), semi_cols=(0,))
    body = (P.chrome("reports", "Workbooks", "How many files each framework is built into. Blank: one workbook. Name "
                     "a field, and each of its values gets a workbook of its own - with every sheet Pivot config "
                     "asks for inside it.", "1 framework(s) split into one workbook per value; the rest one "
                     "workbook each.", "OK", ["Check", "Restore defaults"], tab="Workbooks",
                     overline="REPORTS &nbsp;&#183;&nbsp; WORKBOOKS") + t)
    return P.page("Workbooks", body, 1700)


def _default_states():
    """Which gallery rows are already on the sheets, and on or off."""
    st = {}
    for r in P.default_recipes():
        st[("pivot", r[1].lower())] = "on" if r[0] == "Yes" else "off"
    for r in rows_by_heading("modPD_Charts", "ChartRec", heads_of("modPD_Charts", "ChartHeads")):
        st[("chart", r[1].lower())] = "on" if r[0] == "Yes" else "off"
    return st


def gallery_sheet():
    card_w, card_h, gap, left0 = 300, 164, 14, 14
    tpl = templates()
    states = _default_states()
    out = []
    y = 0.0
    kind_prev = None
    col = 0
    n_on = 0
    for t in tpl:
        if t["kind"] != kind_prev:
            if kind_prev is not None:
                y += card_h + gap
            out.append("<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";"
                       "font-size:%.1fpx;letter-spacing:1.6px;color:#%s'>%s</div>" % (
                           px(left0 + 2), px(y + 12), px(8), C["M300"],
                           "PIVOT REPORTS" if t["kind"] == "pivot" else "CHARTS"))
            y += 30
            col = 0
            kind_prev = t["kind"]
        if col == 3:
            y += card_h + gap
            col = 0
        x = left0 + col * (card_w + gap)
        state = states.get((t["kind"], t["name"].lower()), "")
        n_on += state == "on"
        out.append(card(t, x, y, card_w, card_h, state))
        col += 1
    y += card_h + gap
    grid = "<div style='position:relative;height:%.1fpx'>%s</div>" % (px(y + 10), "".join(out))
    npiv = sum(t["kind"] == "pivot" for t in tpl)
    body = (P.chrome("reports", "Gallery", "Reports worth having, ready to build. Add puts the report on Pivot config "
                     "or Chart config, switched on; from there it is a row like any other - change what you like.",
                     "%d pivot reports and %d charts. %d already in the next build." % (npiv, len(tpl) - npiv, n_on),
                     "Idle", [], tab="Gallery", overline="REPORTS &nbsp;&#183;&nbsp; GALLERY") + grid)
    return P.page("Gallery", body, 1320), px(y + 10) + 250


def card(t, x, y, w, h, state):
    glyph = ""
    gx, gy = x + w - 64, y + 16
    if t["kind"] == "chart":
        for i, bh in enumerate((14, 24, 19, 30, 11)):
            glyph += "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:#%s'>" \
                     "</div>" % (px(gx + i * 10), px(gy + 32 - bh), px(7), px(bh), "16B07F" if i == 3 else C["DEEP"])
    else:
        for i, bw in enumerate((48, 40, 44, 30)):
            glyph += ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:#%s'>"
                      "</div><div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;"
                      "background:#%s'></div>") % (px(gx), px(gy + 2 + i * 8), px(14), px(4), C["LINE2"],
                                                   px(gx + 18), px(gy + 2 + i * 8), px(bw - 18), px(4),
                                                   "16B07F" if i == 0 else C["DEEP"])
    cap, kind = {"on": ("Open", 2), "off": ("Switch on", 1)}.get(state, ("Add", 1))
    btn_bg, btn_line, btn_tx = ((C["E900"], C["DEEP"], C["TX1"]) if kind == 2 else (C["E950"], C["DEEP"], C["SOFT"]))
    note = ""
    if state == "on":
        note = ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
                "letter-spacing:1px;color:#16B07F'>&#9679;&nbsp;&nbsp;IN THE NEXT BUILD</div>") % (
            px(x + 112), px(y + h - 29), px(7))
    elif state == "off":
        note = ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
                "letter-spacing:1px;color:#%s'>ON THE SHEET, SWITCHED OFF</div>") % (
            px(x + 112), px(y + h - 29), px(7), C["TX3"])
    return ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:#%s;"
            "border-radius:9px;box-shadow:inset 0 0 0 1px #%s'></div>%s"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "letter-spacing:1.2px;color:#%s'>%s</div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "color:#%s'>%s</div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;font-size:%.1fpx;line-height:1.3;"
            "color:#%s'>%s</div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;font-size:%.1fpx;color:#%s'>%s</div>"
            "<span class='pill' style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;"
            "background:#%s;box-shadow:inset 0 0 0 1px #%s;color:#%s'>%s</span>%s") % (
        px(x), px(y), px(w), px(h), SURFACE, C["BRAND"] if state == "on" else C["LINE2"], glyph,
        px(x + 16), px(y + 16), px(7), C["M300"], "PIVOT" if t["kind"] == "pivot" else "CHART",
        px(x + 16), px(y + 30), px(12.5), C["TX1"], html.escape(t["title"]),
        px(x + 16), px(y + 54), px(w - 32), px(9), C["TX2"], html.escape(t["desc"]),
        px(x + 16), px(y + 104), px(w - 32), px(7.5), C["TX3"], html.escape(t["uses"]),
        px(x + 16), px(y + h - 36), px(86), px(24), btn_bg, btn_line, btn_tx, cap, note)


# ---------------------------------------------------------------------------
#  the sample, staged
# ---------------------------------------------------------------------------

def staged():
    import openpyxl
    wb = openpyxl.load_workbook(SAMPLE, read_only=True)
    it = wb.worksheets[0].iter_rows(values_only=True)
    hdr = next(it)
    ix = {}
    for i, h in enumerate(hdr):
        ix.setdefault(h, i)
    rows = []
    for r in it:
        pre = float(r[ix["CASHFLOW_AMOUNT_LCY_PRE_FACTOR"]] or 0)
        post = float(r[ix["CASHFLOW_AMOUNT_LCY_POST_FACTOR"]] or 0)
        rows.append(dict(
            type=r[ix["COA_BALANCESHEET_CATEGORY"]] or "(no type)",
            cp=labels.tidy(r[ix["COUNTERPARTY_NAME"]], 3) if r[ix["COUNTERPARTY_NAME"]] else "(no counterparty)",
            rule=r[ix["ALM_PORTFOLIO_SEGMENTATION_RULE_NAME"]] or "(no rule)",
            ccy=r[ix["CURRENCY_NAME"]], pre=pre, post=post))
    return rows


def top(rows, key, val, n):
    agg = collections.defaultdict(float)
    for r in rows:
        agg[r[key]] += val(r)
    return sorted(agg.items(), key=lambda kv: -kv[1])[:n]


# ---------------------------------------------------------------------------
#  charts, as modPD_Charts draws them
# ---------------------------------------------------------------------------

def nice_ticks(vmax, n=5):
    if vmax <= 0:
        return [0, 1]
    raw = vmax / n
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    ticks = [0]
    while ticks[-1] < vmax:
        ticks.append(ticks[-1] + step)
    return ticks


def auto_units(vmax):
    """modPD_Charts.UnitsFor: the unit the biggest figure reads best in."""
    for lim, div, suffix in ((1e9, 1e9, "bn"), (1e6, 1e6, "m"), (1e3, 1e3, "k")):
        if vmax >= lim:
            return div, suffix
    return 1, ""


def fmt_units(v, unit, dec=0):
    div, suffix = unit
    return ("{:,.%df}%s" % (dec, " " + suffix if suffix else "")).format(v / div)


def chart_card(w, h, title, inner):
    return ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:#%s;"
            "box-shadow:inset 0 0 0 1px #%s;box-sizing:border-box'>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "color:#%s'>%s</div>%s</div>") % (0, 0, w, h, SURFACE, C["LINE"], 13, 8, px(11), C["TX1"],
                                            html.escape(title), inner)


def doughnut(w, h, title, items, palette):
    total = sum(v for _, v in items) or 1
    cx, cy, r_out = w * 0.32, h * 0.56, min(w * 0.26, h * 0.36)
    r_in = r_out * 0.6
    a0 = -math.pi / 2
    paths, lbls = [], []
    for i, (name, v) in enumerate(items):
        a1 = a0 + 2 * math.pi * v / total
        large = 1 if a1 - a0 > math.pi else 0
        p = (cx + r_out * math.cos(a0), cy + r_out * math.sin(a0), cx + r_out * math.cos(a1), cy + r_out * math.sin(a1),
             cx + r_in * math.cos(a1), cy + r_in * math.sin(a1), cx + r_in * math.cos(a0), cy + r_in * math.sin(a0))
        paths.append("<path d='M%.1f %.1f A%.1f %.1f 0 %d 1 %.1f %.1f L%.1f %.1f A%.1f %.1f 0 %d 0 %.1f %.1f Z' "
                     "fill='#%s' stroke='#%s' stroke-width='2'/>" % (
                         p[0], p[1], r_out, r_out, large, p[2], p[3], p[4], p[5], r_in, r_in, large, p[6], p[7],
                         palette[i % len(palette)], SURFACE))
        if v / total >= 0.05:
            am = (a0 + a1) / 2
            rm = (r_out + r_in) / 2
            lbls.append("<text x='%.1f' y='%.1f' fill='#%s' font-size='%.1f' text-anchor='middle' "
                        "dominant-baseline='middle' font-family='Selawik Semibold'>%d%%</text>" % (
                            cx + rm * math.cos(am), cy + rm * math.sin(am), "000000" if i % len(palette) in (2, 3)
                            else C["TX1"], px(8), round(100 * v / total)))
        a0 = a1
    legend = "".join(
        "<rect x='%.1f' y='%.1f' width='9' height='9' rx='2' fill='#%s'/><text x='%.1f' y='%.1f' fill='#%s' "
        "font-size='%.1f' dominant-baseline='middle' font-family='Selawik'>%s</text>" % (
            w * 0.64, h * 0.34 + i * 20 - 4.5, palette[i % len(palette)], w * 0.64 + 15, h * 0.34 + i * 20,
            C["TX2"], px(8.5), html.escape(name)) for i, (name, _) in enumerate(items))
    svg = "<svg width='%.1f' height='%.1f' style='position:absolute;left:0;top:0'>%s%s%s</svg>" % (
        w, h, "".join(paths), "".join(lbls), legend)
    return chart_card(w, h, title, svg)


def hbars(w, h, title, cats, series, palette, legend=None, label_w=None):
    """Horizontal bars, the first category on top; series side by side."""
    nser = len(series)
    vmax = max(max(s) for s in series)
    unit = auto_units(vmax)
    ticks = nice_ticks(vmax)
    top_y = 40 if not legend else 58
    bottom = 26
    lw = label_w or min(w * 0.42, max(len(c) for c in cats) * 6.3 + 16)
    x0, x1 = lw, w - 64
    band = (h - top_y - bottom) / len(cats)
    bar = min(18, band * 0.62 / nser)
    out = []
    for t in ticks:
        x = x0 + (x1 - x0) * t / ticks[-1]
        out.append("<line x1='%.1f' x2='%.1f' y1='%.1f' y2='%.1f' stroke='#%s' stroke-width='1'/>" % (
            x, x, top_y - 4, h - bottom, C["LINE"]))
        out.append("<text x='%.1f' y='%.1f' fill='#%s' font-size='%.1f' text-anchor='middle' "
                   "font-family='Selawik'>%s</text>" % (x, h - bottom + 15, C["TX3"], px(8.5), fmt_units(t, unit, 0 if all(abs(k / unit[0] - round(k / unit[0])) < 1e-9 for k in ticks) else 1)))
    for i, c in enumerate(cats):
        yc = top_y + band * (i + 0.5)
        name = c if len(c) <= (lw - 16) / 6.3 else c[:int((lw - 16) / 6.3) - 1] + "…"
        out.append("<text x='%.1f' y='%.1f' fill='#%s' font-size='%.1f' text-anchor='end' dominant-baseline='middle' "
                   "font-family='Selawik'>%s</text>" % (x0 - 8, yc, C["TX2"], px(8.5), html.escape(name)))
        for k, s in enumerate(series):
            y = yc - nser * bar / 2 + k * bar
            bw = (x1 - x0) * s[i] / ticks[-1]
            out.append("<rect x='%.1f' y='%.1f' width='%.1f' height='%.1f' fill='#%s'/>" % (
                x0, y + 1, max(bw, 1), bar - 2, palette[k % len(palette)]))
            out.append("<text x='%.1f' y='%.1f' fill='#%s' font-size='%.1f' dominant-baseline='middle' "
                       "font-family='Selawik'>%s</text>" % (x0 + bw + 5, y + bar / 2, C["TX1"], px(8),
                                                          fmt_units(s[i], unit, 2 if unit[1] == "bn" else 1)))
    out.append("<line x1='%.1f' x2='%.1f' y1='%.1f' y2='%.1f' stroke='#%s' stroke-width='1'/>" % (
        x0, x0, top_y - 4, h - bottom, C["LINE2"]))
    if legend:
        lx = 13
        for k, name in enumerate(legend):
            out.append("<rect x='%.1f' y='34' width='9' height='9' rx='2' fill='#%s'/><text x='%.1f' y='39' fill='#%s' "
                       "font-size='%.1f' dominant-baseline='middle' font-family='Selawik'>%s</text>" % (
                           lx, palette[k], lx + 14, C["TX2"], px(8.5), name))
            lx += 14 + len(name) * 7 + 18
    svg = "<svg width='%.1f' height='%.1f' style='position:absolute;left:0;top:0'>%s</svg>" % (w, h, "".join(out))
    return chart_card(w, h, title, svg)


# ---------------------------------------------------------------------------
#  the built workbook
# ---------------------------------------------------------------------------

def start_here(rows):
    wide = P.colw(38) + P.colw(110) - 18
    gap = px(12)
    ch_h = px(250)
    third = (wide - 2 * gap) / 3
    two = third * 2 + gap

    by_type = top(rows, "type", lambda r: abs(r["pre"]), 8)
    named = [r for r in rows if r["cp"] != "(no counterparty)"]
    cps = top(named, "cp", lambda r: abs(r["pre"]), 10)
    ruled = [r for r in rows if r["rule"] != "(no rule)"]
    rules = top(ruled, "rule", lambda r: abs(r["pre"]), 8)
    kept = {k: 0.0 for k, _ in rules}
    for r in rows:
        if r["rule"] in kept:
            kept[r["rule"]] += abs(r["post"])
    c1 = doughnut(third, ch_h, "Where the balance sits", by_type, PALETTES["avati"])
    c2 = hbars(two, ch_h, "Largest counterparties", [k for k, _ in cps], [[v for _, v in cps]], PALETTES["avati"])
    c3 = hbars(wide, ch_h, "What the factors kept, by rule", [k for k, _ in rules],
               [[v for _, v in rules], [kept[k] for k, _ in rules]], PALETTES["two tone"], legend=["Read", "Kept"])
    grid = ("<div style='position:relative;height:%.1fpx;margin-left:%.1fpx'>"
            "<div style='position:absolute;left:0;top:%.1fpx;width:%.1fpx;height:%.1fpx'>%s</div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx'>%s</div>"
            "<div style='position:absolute;left:0;top:%.1fpx;width:%.1fpx;height:%.1fpx'>%s</div></div>") % (
        2 * (ch_h + gap), px(14) - 9, gap / 2, third, ch_h, c1, third + gap, gap / 2, two, ch_h, c2,
        ch_h + gap * 1.5, wide, ch_h, c3)

    gross = sum(abs(r["pre"]) for r in rows)
    post = sum(abs(r["post"]) for r in rows)
    local = collections.Counter(r["ccy"] for r in rows).most_common(1)[0][0]
    rule_names = [k for k, _ in top(ruled, "rule", lambda r: abs(r["pre"]), 3)]
    made = [("LCR Output", "Every rule, in the order the engine evaluates them, against what it read and what it kept."),
            ("Balance sheet", "The same balances as the balance sheet reads them, down to the COA. Pre-factor only."),
            ("LCR top counterparties", "The 25 largest counterparties by gross exposure - each one's share of the "
                                       "25 and the running share down the list.")]
    made += [(n[:27] + " LCY", n + "  -  LCY") for n in rule_names]
    items = [("ROWS STAGED", format(len(rows), ",")), ("PRE-FACTOR", P.compact(gross)),
             ("POST-FACTOR", P.compact(post)), ("SHEETS", str(len(made))), ("LOCAL CURRENCY", local),
             ("DATA AS OF", "30 Nov 2025")]
    tw = (wide / P.PT - 5 * 10) / 6

    def section(text):
        return ("<div style='height:%.1fpx;width:%dpx;box-sizing:border-box;border-bottom:1px solid #%s;display:flex;"
                "align-items:flex-end;padding:0 0 5px 9px;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
                "letter-spacing:.6px;color:#%s'>%s</div>") % (px(30), P.colw(38) + P.colw(110), C["DEEP"], px(8),
                                                             C["M300"], text)
    status = ("<div style='width:1060px;height:%.1fpx;background:#%s;border-left:4px solid #%s;display:flex;"
              "align-items:center;padding-left:12px;box-sizing:border-box;font-family:\"Selawik Semibold\";"
              "font-size:%.1fpx;color:#%s'>&#9679;&nbsp;&nbsp;%s rows staged into one pivot cache.&nbsp; Every sheet "
              "below is a live PivotTable over it - drag a field, drop a slicer, drill a total.</div>") % (
        px(26), C["OK_BG"], C["OK"], px(9), C["OK"], format(len(rows), ","))
    tbl = P.table(["Sheet", "What is on it"], [38, 110], [[a, b] for a, b in made], muted_cols=(1,))
    tbl = re.sub(r'(<tr style="height:[0-9.]+px;background:#[0-9A-F]+"><td style=")',
                 r"\1color:#%s;font-family:'Selawik Semibold';" % C["M300"], tbl)
    built = "Built by Avati ALM Desk 3.0 on 26 Sep 2026, 14:05   ·   data as of 30 Nov 2025"
    body = (P.book_bar("LCR  ·  MIDBANK CAIRO", back=False, prev_next=False) +
            P.title_block("START HERE &nbsp;&#183;&nbsp; MIDBANK CAIRO", "LCR", built) +
            P.tiles(items, fixed_w=tw, h_row=68) + status + "<div style='height:%.1fpx'></div>" % px(12) +
            section("AT A GLANCE &nbsp;&#183;&nbsp; 3 CHARTS") + grid + "<div style='height:%.1fpx'></div>" % px(14) +
            section("IN THIS BOOK &nbsp;&#183;&nbsp; %d SHEETS" % len(made)) + tbl)
    return P.page("Start here", body, 1500)


def top_counterparties(rows):
    n = 25
    cps = top([r for r in rows if r["cp"] != "(no counterparty)"], "cp", lambda r: abs(r["pre"]), n)
    tot = sum(v for _, v in cps)
    # FitPivot: the label column as wide as its longest label, capped at 62.
    widths = [min(62, max(len(k) for k, _ in cps)) + 3, 16, 12, 14]
    cols = "".join('<col style="width:%dpx">' % P.colw(w) for w in widths)
    hd = "font-family:\"Selawik Semibold\";color:#%s;background:#000;border-bottom:3px solid #%s" % (C["SOFT"],
                                                                                                   C["BRAND"])
    head = "<tr style='height:%.1fpx'>%s</tr>" % (px(22), "".join(
        "<td style='%s;%s'>%s</td>" % (hd, "text-align:right;padding-right:9px" if i else "", h)
        for i, h in enumerate(["Counterparty&nbsp;&#9662;", "Exposure", "Share", "Cumulative"])))
    body_rows = []
    run = 0.0
    bar_max = cps[0][1]
    for i, (name, v) in enumerate(cps):
        run += v
        bg = C["ROW_ALT"] if i % 2 else C["ROW"]
        wbar = 100 * v / bar_max
        body_rows.append(
            ("<tr style='height:%.1fpx;background:#%s'><td style='border-bottom:1px solid #%s'>%s</td>"
             "<td style='border-bottom:1px solid #%s;text-align:right;padding-right:9px;"
             "background:linear-gradient(90deg,#%s %.1f%%,transparent %.1f%%)'>%s</td>"
             "<td style='border-bottom:1px solid #%s;text-align:right;padding-right:9px'>%.1f%%</td>"
             "<td style='border-bottom:1px solid #%s;text-align:right;padding-right:9px'>%.1f%%</td></tr>") % (
                px(20), bg, C["LINE"], html.escape(name), C["LINE"], "00794F", wbar, wbar,
                "{:,.1f}".format(v / 1e6), C["LINE"], 100 * v / tot, C["LINE"], 100 * run / tot))
    total_row = ("<tr style='height:%.1fpx;background:#%s'><td style='font-family:\"Selawik Semibold\";"
                 "border-top:3px solid #%s'>Total</td>%s</tr>") % (
        px(20), C["E950"], C["BRAND"], "".join(
            "<td style='text-align:right;padding-right:9px;font-family:\"Selawik Semibold\";border-top:3px solid #%s'>"
            "%s</td>" % (C["BRAND"], t) for t in ("{:,.1f}".format(tot / 1e6), "100.0%", "")))
    slicer = ("<div style='padding:%.1fpx 0 0 %.1fpx;height:%.1fpx;box-sizing:border-box'>"
              "<div style='width:%.1fpx;height:%.1fpx;background:#%s;box-shadow:inset 0 0 0 1px #%s;border-radius:4px;"
              "padding:6px 8px;box-sizing:border-box'><div style='color:#%s;font-family:\"Selawik Semibold\";"
              "font-size:11.5px;margin-bottom:2px'>LCY / FCY</div>%s</div></div>") % (
        px(6), px(14), px(74), px(156), px(62), C["ROW"], C["LINE2"], C["TX1"], "".join(
            "<div style='height:%.1fpx;border-radius:3px;margin:3px 0;padding-left:8px;font-size:11px;"
            "line-height:%.1fpx;background:#%s;color:#%s'>%s</div>" % (px(15), px(15), C["DEEP"], C["TX1"], t)
            for t in ("LCY", "FCY")))
    body = (P.book_bar("LCR  ·  MIDBANK CAIRO") +
            P.title_block("LCR &nbsp;&#183;&nbsp; PIVOT &nbsp;&#183;&nbsp; IN MILLIONS", "LCR top counterparties",
                          "The 25 largest counterparties by gross exposure - each one's share of the 25 and the "
                          "running share down the list. Slice by LCY / FCY.") +
            P.tiles([("EXPOSURE &nbsp;&#183;&nbsp; TOTAL", P.compact(tot))]) +
            "<div style='height:%.1fpx'></div>" % px(6) + slicer + "<div style='height:%.1fpx'></div>" % px(10) +
            "<table>%s%s%s%s</table>" % (cols, head, "".join(body_rows), total_row))
    return P.page("LCR top counterparties", body, 1100)


def main():
    out = P.OUT
    os.makedirs(out, exist_ok=True)
    jobs = [("sheet-chart-config", chart_config_sheet(), 3400, 470),
            ("sheet-workbooks", workbooks_sheet(), 1700, 380)]
    g, gh = gallery_sheet()
    jobs.append(("sheet-gallery", g, 1320, int(gh)))
    if SAMPLE and os.path.exists(SAMPLE):
        rows = staged()
        jobs.append(("built-start-here", start_here(rows), 1500, 1330))
        jobs.append(("built-top-counterparties", top_counterparties(rows), 1100, 900))
    for name, page_html, w, h in jobs:
        path = os.path.join(out, name + ".png")
        assets.render(page_html, path, w, h, scale=1.5 if w < 2000 else 1, transparent=False)
        print(path)


if __name__ == "__main__":
    main()
