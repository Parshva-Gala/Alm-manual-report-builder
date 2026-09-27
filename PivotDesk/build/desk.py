"""
The Desk, described once.

desk(state) returns every shape on the Desk for a given state of the desk.
The build calls it with the state the workbook ships in; the preview calls
it with the same state and with a busy, populated one. The VBA module
modPD_Desk repaints the same shapes by name at runtime, so the names here
are a contract: the build checks that every name the code paints exists.

Layout, in points, on a 1120 x 660 canvas:

    0    app bar
    68   hero: greeting, the one-sentence state, the next action, 4 tiles
    236  three cards - Add files, Build pivots, Reconcile
    524  recent activity, and the workbooks the last builds wrote
    642  footer

Over everything, hidden until asked for: the busy veil, the toast, the
macros-off banner, and the guided tour (a dimmed desk with one part lit, and
a card that explains it).
"""

from __future__ import annotations

import os

import assets
from design_tokens import *  # noqa: F401,F403
from shapes import Body, Glow, Line, Linear, Para, Run, Shadow, Shape, Solid, T

ASSETS = os.path.join(assets.OUT)
# The Avati logo (build/brand.py makes it from the supplied artwork).
LOGO = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "design", "brand", "avati-logo")
LOGO_RATIO = 785 / 205          # width / height of the trimmed mark
LOGO_H = 20

# ---------------------------------------------------------------------------
#  Icons the Desk uses, rendered on demand.
# ---------------------------------------------------------------------------
_ICON_COLORS = {"em": EM[300], "white": "FFFFFF", "ink": "04110C", "muted": TX_3, "soft": EM[200],
                "t2": TX_2}
_needed = set()


def icon(name, variant):
    _needed.add((name, variant))
    return os.path.join(ASSETS, "icon-%s-%s" % (name, variant))


def render_icons():
    os.makedirs(ASSETS, exist_ok=True)
    for name, variant in sorted(_needed):
        base = os.path.join(ASSETS, "icon-%s-%s" % (name, variant))
        svg = assets.icon_svg(name, _ICON_COLORS[variant])
        stale = not os.path.exists(base + ".svg") or open(base + ".svg").read() != svg
        if stale or not os.path.exists(base + ".png"):
            with open(base + ".svg", "w", encoding="utf-8") as f:
                f.write(svg)
            page = ('<!doctype html><html><body style="margin:0;background:transparent">'
                    '<div style="width:24px;height:24px">%s</div></body></html>') % svg
            assets.render(page, base + ".png", 24, 24, scale=4)


# ---------------------------------------------------------------------------
#  Building blocks
# ---------------------------------------------------------------------------
CARD_FILL = Linear(90, [(0, SURFACE_HI, 1), (1, SURFACE, 1)])
CARD_LINE = Line(0.75, HAIR_2, 1)
CARD_SHADOW = Shadow(14, 4, 90, "000000", 0.30)


def text(name, x, y, w, h, paras, anchor="t", wrap=True, macro="", descr="", clip=False, hidden=False):
    if isinstance(paras, Para):
        paras = [paras]
    return Shape(name, x, y, w, h, body=Body(paras, anchor, wrap, (0, 0, 0, 0), clip),
                 macro=macro, descr=descr, hidden=hidden)


def pic(name, x, y, w, h, base, macro="", descr="", svg=True, alpha=1.0, hidden=False):
    return Shape(name, x, y, w, h, image=base, image_svg=svg, macro=macro, descr=descr,
                 image_alpha=alpha, hidden=hidden)


def hit(name, x, y, w, h, macro, descr):
    """An invisible click target laid over a row of text and dots."""
    return Shape(name, x, y, w, h, fill=Solid("FFFFFF", 0.0), macro=macro, descr=descr)


BUTTON = {
    #          fill                                         line                  text     glow
    "primary": (Linear(0, [(0, EM[400], 1), (1, EM[500], 1)]), None, "04110C", Glow(7, EM[500], 0.35)),
    "soft":    (Solid(EM[950], 1), Line(0.75, EM[700], 1), EM[100], None),
    "ghost":   (Solid("FFFFFF", 0.03), Line(0.75, HAIR_2, 1), TX_1, None),
    "off":     (Solid("FFFFFF", 0.02), Line(0.75, HAIR, 1), TX_4, None),
}
# Only the hero carries a primary: one bright action per screen. Card actions
# are emerald-tinted, and dim when they cannot act.


def button(name, x, y, w, label, kind="primary", macro="", icon_base=None, h=28, size=9.0, descr=""):
    fill, line, color, glow = BUTTON[kind]
    left = 12
    out = []
    pad_l = 0
    if icon_base:
        pad_l = 18
    b = Shape(name, x, y, w, h, "roundRect", h / 2, fill=fill, line=line, glow=glow,
              shadow=Shadow(8, 2, 90, "000000", 0.35) if kind == "primary" else None,
              body=Body([T(label, size, color, F_SEMI, align="ctr")], "ctr", False, (pad_l, 0, 0, 0)),
              macro=macro, descr=descr or label)
    out.append(b)
    if icon_base:
        # icon sits to the left of centred text; both carry the same action
        est = len(label) * size * 0.52
        ix = x + (w + pad_l) / 2 - est / 2 - 18
        out.append(pic(name + "_ic", ix, y + (h - 13) / 2, 13, 13, icon_base, macro=macro, descr=label))
    return out


def pill(name, x, y, w, h, label, fg, bg, size=6.5, border=None, descr=""):
    return Shape(name, x, y, w, h, "roundRect", h / 2, fill=Solid(bg), line=border,
                 body=Body([T(label, size, fg, F_SEMI, spc=0.6, align="ctr")], "ctr", False),
                 descr=descr or label)


def dot(name, x, y, color, d=7, descr=""):
    return Shape(name, x, y, d, d, "ellipse", fill=Solid(color), descr=descr or "status")


LEVEL = {  # dot / pill colours on dark, keyed by the tool's own verdict words
    "OK": (OK, OK_BG), "LOADED": (OK, OK_BG),
    "CHECK": (WARN, WARN_BG), "WARN": (WARN, WARN_BG), "FORCED": (WARN, WARN_BG),
    "BREAK": (BAD, BAD_BG), "MISSING": (BAD, BAD_BG), "FAILED": (BAD, BAD_BG),
    "IDLE": (IDLE, IDLE_BG), "EMPTY": ("3A4B43", IDLE_BG), "NONE": ("3A4B43", IDLE_BG),
}


VERDICT_WORD = {"OK": "OK", "BREAK": "Break", "CHECK": "Check", "NOT SUPPLIED": "—", "—": "—", "": "—"}


def lv(word):
    return LEVEL.get((word or "IDLE").upper(), LEVEL["IDLE"])


# ---------------------------------------------------------------------------
#  The Desk
# ---------------------------------------------------------------------------
CARD_Y, CARD_H, CARD_W = 236, 272, 344
CARD_X = [MARGIN, MARGIN + CARD_W + GUTTER, MARGIN + 2 * (CARD_W + GUTTER)]
BOT_Y, BOT_H = 524, 108
ACT_W, ACT_ROWS = 612, 4
RB_X = MARGIN + ACT_W + GUTTER
RB_W = W - MARGIN - RB_X
RB_ROWS = 3


def desk(st):
    S = []
    add = S.extend

    # --- app bar --------------------------------------------------------------
    S.append(Shape("pdx_bar", 0, 0, W, 52, fill=Solid(INK), descr="Avati app bar"))
    S.append(Shape("pdx_bar_rule", 0, 51.25, W, 0.75,
                   fill=Linear(0, [(0, EM[900], 1), (0.5, EM[600], 1), (1, EM[900], 1)]), descr="rule"))
    # The Avati mark, then what this is and whose it is.
    lw = round(LOGO_H * LOGO_RATIO, 2)
    S.append(pic("pdx_logo", MARGIN, 26 - LOGO_H / 2, lw, LOGO_H, LOGO, svg=False, macro="PD_GoHome",
                 descr="Avati"))
    S.append(Shape("pdx_logo_rule", MARGIN + lw + 12, 15, 0.75, 22, fill=Solid(HAIR_2), descr="rule"))
    S.append(text("pdx_wordmark", MARGIN + lw + 24, 12.5, 120, 14, T("ALM DESK", 8, TX_1, F_SEMI, spc=2.0),
                  wrap=False, descr="ALM Desk"))
    S.append(text("pdx_bank", MARGIN + lw + 24, 27, 120, 12, T("MIDBANK  ·  CAIRO", 6.5, TX_3, F_SEMI, spc=1.6),
                  wrap=False, descr="MIDBANK Cairo"))

    # navigation
    nav_x, nav_y = 262, 13
    items = [("desk", "Desk", "spark", 66, "PD_GoHome"), ("files", "Files", "files", 64, "PD_GoFiles"),
             ("config", "Reports", "layers", 84, "PD_GoConfig"),
             ("recon", "Reconciliation", "balance", 110, "PD_GoRecon"), ("log", "Activity", "pulse", 80, "PD_GoLog")]
    total = 3 + sum(i[3] for i in items) + 2 * (len(items) - 1) + 3
    S.append(Shape("pdx_nav", nav_x, nav_y, total, 26, "roundRect", 13, fill=Solid("0B1310"),
                   line=Line(0.75, HAIR, 1), descr="navigation"))
    x = nav_x + 3
    for key, label, ic, w, mac in items:
        active = key == "desk"
        S.append(Shape("pdx_nav_%s" % key, x, nav_y + 2, w, 22, "roundRect", 11,
                       fill=Solid(EM[900]) if active else Solid("FFFFFF", 0.0),
                       line=Line(0.75, EM[700], 1) if active else None,
                       body=Body([T(label, 8.5, TX_1 if active else TX_2, F_SEMI, align="ctr")], "ctr", False, (14, 0, 0, 0)),
                       macro=mac, descr="Go to " + label))
        S.append(pic("pdx_nav_%s_ic" % key, x + (w + 14) / 2 - len(label) * 2.25 - 15, nav_y + 8.5, 10, 10,
                     icon(ic, "em" if active else "t2"), macro=mac, descr=label))
        x += w + 2

    # right side: as-of chip, Excel view, Reset, the tour
    S.append(Shape("pdx_asof_chip", 714, 13, 172, 26, "roundRect", 13, fill=Solid("0B1310"),
                   line=Line(0.75, HAIR, 1), descr="data as of"))
    S.append(pic("pdx_asof_ic", 725, 20, 12, 12, icon("calendar", "em"), descr="as of"))
    S.append(text("pdx_asof", 742, 13, 140, 26, T(st["asof"], 7, TX_2, F_SEMI, spc=0.8), anchor="ctr",
                  wrap=False, descr="data as of"))
    add(button("pdx_btn_view", 894, 13, 86, st["view_label"], "ghost", "PD_ToggleAppView",
               icon("expand", "t2"), h=26, size=8, descr="Show or hide Excel's ribbon"))
    add(button("pdx_btn_reset", 986, 13, 72, "Reset", "ghost", "PD_Reset", icon("reset", "t2"), h=26, size=8,
               descr="Clear loaded files and results"))
    S.append(Shape("pdx_btn_help", W - MARGIN - 26, 13, 26, 26, "ellipse", fill=Solid("FFFFFF", 0.03),
                   line=Line(0.75, HAIR_2, 1), macro="PD_TourStart", descr="A one-minute tour of the desk (F1)"))
    S.append(pic("pdx_btn_help_ic", W - MARGIN - 26 + 6, 19, 14, 14, icon("help", "t2"), macro="PD_TourStart",
                 descr="Tour"))

    # --- hero -----------------------------------------------------------------
    HX, HY, HW, HH = MARGIN, 68, W - 2 * MARGIN, 152
    S.append(Shape("pdx_hero_art", HX, HY, HW, HH, "roundRect", 14,
                   fill=Linear(0, [(0, SURFACE_HI, 1), (0.55, EM[950], 1), (1, SURFACE, 1)]),
                   line=Line(0.75, EM[800], 1), descr="Midbank ALM operations"))
    S.append(Shape("pdx_hero_accent", HX + 1, HY + 24, 3, HH - 48, "roundRect", 1.5,
                   fill=Solid(EM[400]), descr="Midbank emerald"))
    S.append(text("pdx_hero_context", HX + 24, HY + 17, 258, 12,
                  T("MIDBANK / ALM OPERATIONS", 7, EM[300], F_SEMI, spc=1.2), wrap=False,
                  descr="Midbank asset and liability management"))
    S.append(text("pdx_hero_date", HX + 286, HY + 17, 266, 12, T(st["date"], 6.5, TX_3, F_SEMI, align="r"),
                  wrap=False, descr="today"))
    S.append(text("pdx_hero_greet", HX + 23, HY + 32, 544, 39, T(st["greeting"], 27, TX_1, F_SEMI),
                  wrap=False, descr="current desk priority"))
    S.append(text("pdx_hero_lede", HX + 24, HY + 75, 532, 30, T(st["lede"], 9, TX_2, F_BODY), descr="where the desk stands"))
    add(button("pdx_cta", HX + 24, HY + 112, 188, st["cta"], "primary", "PD_NextAction", icon("arrow", "ink"),
               descr="The next thing to do"))
    add(button("pdx_cta2", HX + 220, HY + 112, 104, st["cta2"], "ghost", "PD_LoadFiles", descr="Pick files"))
    state_fg, state_bg = lv(st.get("hero_level", "IDLE"))
    S.append(pill("pdx_hero_state", HX + 364, HY + 116, 188, 20, st.get("hero_status", "START WITH SOURCE FILES"),
                  state_fg, state_bg, size=6.5, border=Line(0.75, state_fg, 0.35), descr="desk readiness"))

    # KPI tiles
    tw, tg = 102, 10
    tx0 = HX + HW - 24 - (4 * tw + 3 * tg)
    kpis = st["kpis"]
    for i, k in enumerate(kpis):
        x = tx0 + i * (tw + tg)
        y = HY + 18
        nm = "pdx_kpi_%s" % k["key"]
        S.append(Shape(nm, x, y, tw, 116, "roundRect", 10, fill=Solid(CANVAS, 0.85),
                       line=Line(0.75, HAIR_2, 1), descr=k["label"]))
        S.append(text(nm + "_label", x + 12, y + 12, tw - 20, 10, T(k["label"], 6.5, TX_2, F_SEMI, spc=0.75),
                      wrap=False, descr=k["label"]))
        S.append(text(nm + "_value", x + 11, y + 25, tw - 20, 36, T(k["value"], 27, k.get("color", TX_1), F_SEMI),
                      wrap=False, descr=k["label"] + " value"))
        S.append(text(nm + "_sub", x + 12, y + 66, tw - 22, 26, T(k["sub"], 7, TX_3, F_BODY), descr=k["label"]))
        n = k["segments"]
        gap = 3
        sw = (tw - 24 - (n - 1) * gap) / n
        for j in range(n):
            on = j < k["filled"]
            S.append(Shape("%s_seg%d" % (nm, j + 1), x + 12 + j * (sw + gap), y + 100, sw, 4, "roundRect", 2,
                           fill=Solid(k.get("seg_color", EM[400])) if on else Solid("FFFFFF", 0.10),
                           descr="meter"))

    # --- the three cards ------------------------------------------------------------
    cards = [
        ("1", "01", "Connect source files", "Load framework outputs and independent controls. Files are recognised "
         "by their columns.", "upload"),
        ("2", "02", "Build your workbooks", "Choose LCR, NSFR and maturity ladder. Reports defines the live pivots "
         "and charts in every workbook.", "pivot"),
        ("3", "03", "Review & reconcile", "Compare outputs with controls 3 and 6. Review checks and coverage "
         "before interpreting a difference.", "balance"),
    ]
    for (k, num, title, desc, ic), x in zip(cards, CARD_X):
        y = CARD_Y
        S.append(Shape("pdx_card%s" % k, x, y, CARD_W, CARD_H, "roundRect", 14, fill=CARD_FILL,
                       line=CARD_LINE, shadow=CARD_SHADOW, descr=title))
        S.append(Shape("pdx_card%s_sheen" % k, x + 14, y + 0.4, CARD_W - 28, 0.75,
                       fill=Linear(0, [(0, "FFFFFF", 0), (0.5, "FFFFFF", 0.10), (1, "FFFFFF", 0)]), descr="edge"))
        S.append(pill("pdx_card%s_step" % k, x + 20, y + 20, 28, 15, num, EM[300], EM[950],
                      size=7, border=Line(0.75, EM[800], 1), descr="step " + num))
        S.append(text("pdx_card%s_count" % k, x + 56, y + 20, 200, 15,
                      T(st["counts"][int(k) - 1], 6.5, TX_2, F_SEMI, spc=0.55), anchor="ctr", wrap=False,
                      descr="progress"))
        S.append(Shape("pdx_card%s_icon_tile" % k, x + CARD_W - 52, y + 18, 32, 32, "roundRect", 9,
                       fill=Linear(90, [(0, EM[900], 1), (1, EM[950], 1)]), line=Line(0.75, EM[800], 1),
                       descr=title))
        S.append(pic("pdx_card%s_icon" % k, x + CARD_W - 44, y + 26, 16, 16, icon(ic, "em"), descr=title))
        S.append(text("pdx_card%s_title" % k, x + 20, y + 41, 282, 24, T(title, 15, TX_1, F_SEMI), wrap=False,
                      descr=title))
        S.append(text("pdx_card%s_desc" % k, x + 20, y + 64, 300, 28, T(desc, 8.5, TX_2, F_BODY), descr=desc))
        S.append(Shape("pdx_card%s_rule" % k, x + 20, y + 98, CARD_W - 40, 0.75, fill=Solid(HAIR), descr="rule"))

    # connectors between the steps
    for i in range(2):
        cx = CARD_X[i] + CARD_W + GUTTER / 2 - 10
        cy = CARD_Y + 49 - 10
        S.append(Shape("pdx_join%d" % (i + 1), cx, cy, 20, 20, "ellipse", fill=Solid(CANVAS),
                       line=Line(0.75, HAIR_2, 1), descr="then"))
        S.append(pic("pdx_join%d_ic" % (i + 1), cx + 5, cy + 5, 10, 10, icon("chevron", "em"), descr="then"))

    # card 1: the five file slots
    x, y = CARD_X[0], CARD_Y
    for i, slot in enumerate(st["slots"], start=1):
        ry = y + 104 + (i - 1) * 23.5
        fg, _ = lv(slot["level"])
        S.append(dot("pdx_slot%d_dot" % i, x + 22, ry + 8.25, fg, descr=slot["label"] + " status"))
        S.append(text("pdx_slot%d_label" % i, x + 36, ry + 2.5, 130, 18, T(slot["label"], 9, TX_1, F_BODY),
                      anchor="ctr", wrap=False, descr=slot["label"]))
        S.append(text("pdx_slot%d_meta" % i, x + 134, ry + 2.5, 190, 18,
                      T(slot["meta"], 7.5, slot.get("meta_color", TX_3), F_BODY, align="r"),
                      anchor="ctr", wrap=False, clip=True, descr=slot["label"] + " file"))
        if i < 5:
            S.append(Shape("pdx_slot%d_rule" % i, x + 20, ry + 23.25, CARD_W - 40, 0.5, fill=Solid(HAIR, 0.8),
                           descr="rule"))
        S.append(hit("pdx_slot%d_hit" % i, x + 16, ry + 1, CARD_W - 32, 21.5, "PD_SlotClick",
                     "Choose the file for " + slot["label"]))
    add(button("pdx_c1_scan", x + 20, y + 228, 130, "Scan a folder", "soft", "PD_LoadFolder", icon("folder", "soft")))
    add(button("pdx_c1_pick", x + 156, y + 228, 94, "Pick files", "ghost", "PD_LoadFiles"))
    S.append(text("pdx_c1_link", x + 250, y + 228, 74, 28, T("Files  →", 8, EM[300], F_SEMI, align="r"),
                  anchor="ctr", wrap=False, macro="PD_GoFiles", descr="Open the Files sheet"))

    # card 2: framework switches
    x, y = CARD_X[1], CARD_Y
    for j, fw in enumerate(st["frameworks"], start=1):
        ry = y + 104 + (j - 1) * 35
        state = fw["state"]            # on | off | na
        row_fill = {"on": Solid(EM[950]), "off": Solid(SURFACE_2), "na": Solid("FFFFFF", 0.0)}[state]
        row_line = {"on": Line(0.75, EM[800], 1), "off": Line(0.75, HAIR, 1), "na": Line(0.75, HAIR, 0.7)}[state]
        S.append(Shape("pdx_fw%d_row" % j, x + 20, ry, CARD_W - 40, 31, "roundRect", 9, fill=row_fill,
                       line=row_line, descr=fw["label"]))
        S.append(text("pdx_fw%d_label" % j, x + 34, ry + 2.5, 200, 15,
                      T(fw["label"], 9.5, TX_1 if state != "na" else TX_4, F_SEMI), wrap=False, descr=fw["label"]))
        S.append(text("pdx_fw%d_meta" % j, x + 34, ry + 17, 230, 12,
                      T(fw["meta"], 7, TX_3 if state != "na" else TX_4, F_BODY), wrap=False, clip=True,
                      descr=fw["label"] + " detail"))
        track = {"on": EM[400], "off": "2A3B33", "na": "17221D"}[state]
        S.append(Shape("pdx_fw%d_track" % j, x + CARD_W - 60, ry + 8.5, 28, 14, "roundRect", 7,
                       fill=Solid(track), descr=fw["label"] + " switch"))
        kx = x + CARD_W - 60 + (16 if state == "on" else 2)
        S.append(Shape("pdx_fw%d_knob" % j, kx, ry + 10.5, 10, 10, "ellipse",
                       fill=Solid("FFFFFF" if state != "na" else "34443D"),
                       shadow=Shadow(2, 1, 90, "000000", 0.35), descr="switch"))
        S.append(hit("pdx_fw%d_hit" % j, x + 20, ry, CARD_W - 40, 31, "PD_FwToggle", "Include " + fw["label"]))
    S.append(text("pdx_c2_last", x + 20, y + 211, 214, 12, T(st["last_build"], 7, TX_3, F_BODY),
                  wrap=False, clip=True, descr="last build"))
    S.append(text("pdx_c2_config", x + 238, y + 208, 86, 18, T(st.get("config_link", "4 pivots  →"), 8, EM[300], F_SEMI, align="r"),
                  anchor="ctr", wrap=False, macro="PD_GoConfig", descr="Open Pivot config - what each workbook is built from"))
    add(button("pdx_c2_build", x + 20, y + 228, 164, st["build_label"], st["build_kind"], "PD_BuildSelected",
               descr="Build the selected frameworks"))
    add(button("pdx_c2_open", x + 190, y + 228, 106, "Open folder", st.get("open_kind", "ghost"), "PD_OpenOutputFolder",
               descr="Open the folder the last build wrote to"))

    # card 3: controls, the verdict matrix, the overall verdict
    x, y = CARD_X[2], CARD_Y
    for k, ctl in enumerate(st["controls"], start=1):
        cx = x + 20 + (k - 1) * 156
        fg, _ = lv(ctl["level"])
        S.append(Shape("pdx_ctl%d_tile" % k, cx, y + 104, 148, 40, "roundRect", 8, fill=Solid(SURFACE_2),
                       line=Line(0.75, HAIR, 1), descr=ctl["label"]))
        S.append(dot("pdx_ctl%d_dot" % k, cx + 11, y + 113, fg, descr=ctl["label"] + " status"))
        S.append(text("pdx_ctl%d_label" % k, cx + 24, y + 107, 118, 14, T(ctl["label"], 8.5, TX_1, F_SEMI),
                      wrap=False, descr=ctl["label"]))
        S.append(text("pdx_ctl%d_meta" % k, cx + 24, y + 123, 118, 12, T(ctl["meta"], 7, TX_3, F_BODY),
                      wrap=False, clip=True, descr=ctl["label"] + " detail"))
        S.append(hit("pdx_ctl%d_hit" % k, cx, y + 104, 148, 40, "PD_SlotClick", "Choose " + ctl["label"]))
    S.append(text("pdx_c3_when", x + 20, y + 152, 220, 11, T(st["recon_when"], 6.5, TX_3, F_SEMI, spc=1.2),
                  wrap=False, descr="last reconciliation"))
    cols = ["LCR", "NSFR", "LADDER"]
    cell_w, cell_h, gx0 = 50, 18, x + 52
    for f, cname in enumerate(cols):
        S.append(text("pdx_m_head%d" % (f + 1), gx0 + f * (cell_w + 5), y + 165, cell_w, 10,
                      T(cname, 6, TX_3, F_SEMI, spc=1.0, align="ctr"), wrap=False, descr=cname))
    for r_, ctl in enumerate(["CR3", "CR6"]):
        ry = y + 177 + r_ * (cell_h + 5)
        S.append(text("pdx_m_row%d" % (r_ + 1), x + 20, ry, 30, cell_h, T(ctl, 7, TX_2, F_SEMI), anchor="ctr",
                      wrap=False, descr=ctl))
        for f in range(3):
            v = st["matrix"][r_][f]
            fg, bg = lv(v)
            ln = Line(0.75, fg, 0.4)
            if VERDICT_WORD.get(v.upper(), v) == "—":
                # not compared: a placeholder, painted as one (as modPD_Desk does)
                fg, bg, ln = TX_4, IDLE_BG, Line(0.75, HAIR_2, 1)
            S.append(Shape("pdx_m_%d%d" % (r_ + 1, f + 1), gx0 + f * (cell_w + 5), ry, cell_w, cell_h, "roundRect", 5,
                           fill=Solid(bg), line=ln,
                           body=Body([T(VERDICT_WORD.get(v.upper(), v), 6.5, fg, F_SEMI, align="ctr")], "ctr", False),
                           descr="%s against %s" % (ctl, cols[f])))
    vfg, vbg = lv(st["verdict_level"])
    S.append(Shape("pdx_c3_vbox", x + 226, y + 165, 98, 53, "roundRect", 10, fill=Solid(vbg),
                   line=Line(0.75, vfg, 0.45), descr="overall verdict"))
    S.append(text("pdx_c3_vtitle", x + 226, y + 172, 98, 11, T(st["verdict_title"], 6.5, vfg, F_SEMI, spc=1.2, align="ctr"),
                  wrap=False, descr="verdict"))
    S.append(text("pdx_c3_vvalue", x + 226, y + 183, 98, 30, T(st["verdict_value"], 17, TX_1, F_SEMILIGHT, align="ctr"),
                  anchor="ctr", wrap=False, descr="verdict detail"))
    add(button("pdx_c3_run", x + 20, y + 228, 140, "Reconcile now", st["recon_kind"], "PD_Reconcile",
               descr="Reconcile the loaded outputs"))
    add(button("pdx_c3_open", x + 166, y + 228, 112, "Open results", "ghost", "PD_GoRecon",
               descr="Open the Reconciliation sheet"))

    # --- recent activity ----------------------------------------------------------
    ax, ay, aw, ah = MARGIN, BOT_Y, ACT_W, BOT_H
    S.append(Shape("pdx_act", ax, ay, aw, ah, "roundRect", 14, fill=CARD_FILL, line=CARD_LINE,
                   shadow=CARD_SHADOW, descr="recent activity"))
    S.append(pic("pdx_act_ic", ax + 20, ay + 12, 12, 12, icon("pulse", "em"), descr="activity"))
    S.append(text("pdx_act_title", ax + 38, ay + 9, 200, 18, T("Recent activity", 9.5, TX_1, F_SEMI),
                  anchor="ctr", wrap=False, descr="Recent activity"))
    S.append(text("pdx_act_all", ax + aw - 130, ay + 9, 110, 18, T("All activity  →", 8, EM[300], F_SEMI, align="r"),
                  anchor="ctr", wrap=False, macro="PD_GoLog", descr="Open the Activity sheet"))
    acts = st["activity"]
    for a in range(ACT_ROWS):
        ry = ay + 33 + a * 17
        e = acts[a] if a < len(acts) else None
        fg, bg = lv(e["level"]) if e else lv("IDLE")
        S.append(text("pdx_act%d_when" % (a + 1), ax + 20, ry, 74, 14, T(e["when"] if e else "", 7.5, TX_3, F_MONO),
                      anchor="ctr", wrap=False, hidden=e is None, descr="when"))
        S.append(pill("pdx_act%d_lvl" % (a + 1), ax + 96, ry + 1.5, 44, 11, (e["level"] if e else "").upper(),
                      fg, bg, size=6, border=Line(0.75, fg, 0.35), descr="level"))
        S[-1].hidden = e is None
        S.append(text("pdx_act%d_stage" % (a + 1), ax + 150, ry, 54, 14, T(e["stage"] if e else "", 8, TX_2, F_SEMI),
                      anchor="ctr", wrap=False, hidden=e is None, descr="stage"))
        S.append(text("pdx_act%d_msg" % (a + 1), ax + 206, ry, aw - 226, 14,
                      T(e["msg"] if e else "", 8, TX_1, F_BODY), anchor="ctr", wrap=False, clip=True,
                      hidden=e is None, descr="what happened"))
    S.append(pic("pdx_act_art", ax + aw - 256, ay + 18, 240, 78, os.path.join(ASSETS, "lattice-medallion"),
                 svg=False, hidden=bool(acts), descr="pattern"))
    S.append(text("pdx_act_empty", ax + 20, ay + 36, 320, 52,
                  T("Your working history starts here. Loads, builds and reconciliation results appear as you work.",
                    8.5, TX_3, F_BODY), anchor="ctr", hidden=bool(acts), descr="no activity yet"))

    # --- recent builds: what was made, ready to open --------------------------
    gx, gy, gw, gh = RB_X, BOT_Y, RB_W, BOT_H
    builds = st.get("builds") or []
    S.append(Shape("pdx_rb", gx, gy, gw, gh, "roundRect", 14, fill=CARD_FILL, line=CARD_LINE,
                   shadow=CARD_SHADOW, descr="recent builds"))
    S.append(pic("pdx_rb_ic", gx + 20, gy + 12, 12, 12, icon("pivot", "em"), descr="recent builds"))
    S.append(text("pdx_rb_title", gx + 38, gy + 9, 160, 18, T("Recent builds", 9.5, TX_1, F_SEMI),
                  anchor="ctr", wrap=False, descr="Recent builds"))
    S.append(text("pdx_rb_all", gx + gw - 130, gy + 9, 110, 18, T("Open folder  →", 8, EM[300], F_SEMI, align="r"),
                  anchor="ctr", wrap=False, macro="PD_OpenOutputFolder", descr="Open the folder the last build wrote to"))
    for i in range(RB_ROWS):
        ry = gy + 34 + i * 23
        b = builds[i] if i < len(builds) else None
        hidden = b is None
        S.append(pill("pdx_rb%d_fw" % (i + 1), gx + 20, ry + 3.5, 50, 13, b["fw"] if b else "LCR", EM[200], EM[950],
                      size=6, border=Line(0.75, EM[800], 1), descr="framework"))
        S[-1].hidden = hidden
        S.append(text("pdx_rb%d_name" % (i + 1), gx + 78, ry, 236, 12, T(b["name"] if b else " ", 8.5, TX_1, F_BODY),
                      wrap=False, clip=True, hidden=hidden, descr="workbook"))
        S.append(text("pdx_rb%d_meta" % (i + 1), gx + 78, ry + 12, 236, 11, T(b["meta"] if b else " ", 6.5, TX_3, F_BODY),
                      wrap=False, clip=True, hidden=hidden, descr="when and how many sheets"))
        S.append(Shape("pdx_rb%d_open" % (i + 1), gx + gw - 20 - 52, ry + 2, 52, 17, "roundRect", 8.5,
                       fill=Solid("FFFFFF", 0.03), line=Line(0.75, HAIR_2, 1),
                       body=Body([T("Open", 7.5, TX_1, F_SEMI, align="ctr")], "ctr", False),
                       macro="PD_OpenBuild", hidden=hidden, descr="Open this workbook"))
        S.append(hit("pdx_rb%d_hit" % (i + 1), gx + 16, ry - 1, gw - 20 - 52 - 24, 22, "PD_OpenBuild",
                     "Open this workbook"))
        S[-1].hidden = hidden
        if i < RB_ROWS - 1:
            S.append(Shape("pdx_rb%d_rule" % (i + 1), gx + 20, ry + 22.25, gw - 40, 0.5, fill=Solid(HAIR, 0.8),
                           hidden=i + 1 >= len(builds), descr="rule"))
    S.append(text("pdx_rb_empty", gx + 20, gy + 34, gw - 40, 60,
                  [T("Your next deliverable starts above.", 9, TX_2, F_SEMI),
                   T("Build a framework, then open its workbook here. The latest three builds stay within reach.", 8, TX_3, F_BODY)],
                  anchor="ctr", hidden=bool(builds), descr="nothing built yet"))

    # --- footer -------------------------------------------------------------
    S.append(text("pdx_foot", MARGIN, 642, 560, 12,
                  T("Avati ALM Desk %s  ·  MIDBANK  ·  Live pivots, shared caches and traceable builds." % st["version"],
                    7, TX_3, F_BODY), wrap=False,
                  descr="about"))
    S.append(text("pdx_keys", W - MARGIN - 440, 642, 440, 12,
                  T("F1 Tour  ·  Ctrl+Shift+  D Desk  ·  F Files  ·  P Reports  ·  R Reconciliation  ·  A Activity",
                    7, TX_3, F_BODY, align="r"),
                  wrap=False, descr="keyboard shortcuts"))

    # --- overlays: busy, toast, macros-off ---------------------------------------
    S.append(Shape("pdx_busy", MARGIN, CARD_Y, W - 2 * MARGIN, BOT_Y + BOT_H - CARD_Y, "roundRect", 14,
                   fill=Solid(CANVAS, 0.82),
                   body=Body([T("Working…", 16, TX_1, F_LIGHT, align="ctr"),
                              T("Progress is in the status bar at the bottom of the window.", 8.5, TX_2, align="ctr")],
                             "ctr", True), hidden=not st.get("busy"), descr="working"))
    toast = st.get("toast")
    tfg = lv(toast["level"])[0] if toast else OK
    S.append(Shape("pdx_toast", 700, 8, 392, 36, "roundRect", 10, fill=Solid("0E1B16"),
                   line=Line(0.75, tfg, 0.55), shadow=Shadow(20, 6, 90, "000000", 0.6),
                   body=Body([T(toast["text"] if toast else "", 8, TX_1, F_BODY)], "ctr", True, (30, 3, 12, 3)),
                   macro="PD_ToastHide", hidden=toast is None, descr="message - click to dismiss"))
    S.append(Shape("pdx_toast_dot", 713, 22, 8, 8, "ellipse", fill=Solid(tfg), glow=Glow(4, tfg, 0.5),
                   macro="PD_ToastHide", hidden=toast is None, descr="message level"))
    S.append(Shape("pdx_macros", HX + 28, HY + 112, 500, 28, "roundRect", 14, fill=Solid(WARN_BG),
                   line=Line(0.75, WARN, 0.6),
                   body=Body([T("Macros are off  ·  choose Enable Content above the sheet to use the desk.",
                                8.5, WARN, F_SEMI, align="ctr")], "ctr", False),
                   hidden=not st.get("macros_banner", True), descr="macros are off"))

    # --- the tour -----------------------------------------------------------------
    # Four veils around the part being explained leave it lit; a ring marks it
    # and a card beside it says what it is for. modPD_Desk moves the same
    # shapes step by step with the same arithmetic (tour_place).
    step = st.get("tour")
    n = step or 1
    names, title, body = TOUR[n - 1]
    tx, ty, tw, th = union_of(S, names)
    ring, card, veils = tour_place(tx, ty, tw, th)
    for (nm, r) in zip(("t", "b", "l", "r"), veils):
        S.append(Shape("pdx_tour_dim_" + nm, *r, fill=Solid("000000", 0.66), hidden=step is None,
                       descr="tour veil"))
    S.append(Shape("pdx_tour_ring", *ring, "roundRect", 16, line=Line(1.5, EM[300], 1),
                   glow=Glow(10, EM[400], 0.45), hidden=step is None, descr="this part"))
    kx, ky = card
    S.append(Shape("pdx_tour_card", kx, ky, TOUR_W, TOUR_H, "roundRect", 14, fill=Solid("0E1B16"),
                   line=Line(0.75, EM[600], 1), shadow=Shadow(28, 8, 90, "000000", 0.7),
                   hidden=step is None, descr="tour"))
    S.append(text("pdx_tour_step", kx + 20, ky + 18, 150, 10,
                  T("STEP %d OF %d" % (n, len(TOUR)), 6.5, EM[300], F_SEMI, spc=1.6), wrap=False,
                  hidden=step is None, descr="tour step"))
    for p in range(len(TOUR)):
        colour = EM[400] if p + 1 == n else (EM[700] if p + 1 < n else HAIR_2)
        S.append(dot("pdx_tour_pip%d" % (p + 1), kx + TOUR_W - 20 - (len(TOUR) - p) * 11 + 5, ky + 20, colour,
                     d=6, descr="step %d" % (p + 1)))
        S[-1].hidden = step is None
    S.append(text("pdx_tour_title", kx + 20, ky + 32, TOUR_W - 40, 22, T(title, 13, TX_1, F_SEMI), wrap=False,
                  hidden=step is None, descr="tour title"))
    S.append(text("pdx_tour_body", kx + 20, ky + 58, TOUR_W - 40, 40, T(body, 8.5, TX_2, F_BODY),
                  hidden=step is None, descr="tour text"))
    S.append(text("pdx_tour_skip", kx + 20, ky + TOUR_H - 40, 70, 26, T("Skip tour", 8, TX_3, F_SEMI),
                  anchor="ctr", wrap=False, macro="PD_TourEnd", hidden=step is None, descr="End the tour"))
    add(button("pdx_tour_back", kx + TOUR_W - 20 - 76 - 8 - 62, ky + TOUR_H - 40, 62, "Back",
               "ghost" if n > 1 else "off", "PD_TourBack", h=26, size=8.5, descr="Previous step"))
    S[-1].hidden = step is None
    add(button("pdx_tour_next", kx + TOUR_W - 20 - 76, ky + TOUR_H - 40, 76, "Next" if n < len(TOUR) else "Done",
               "primary", "PD_TourNext", h=26, size=8.5, descr="Next step"))
    S[-1].hidden = step is None
    return S


# ---------------------------------------------------------------------------
#  The tour, as arithmetic both sides share
# ---------------------------------------------------------------------------
TOUR_W, TOUR_H, TOUR_PAD, TOUR_GAP = 300, 144, 6, 14
TOUR = [
    (["pdx_hero_greet", "pdx_hero_lede", "pdx_cta", "pdx_cta2", "pdx_hero_state"], "Your next action, in context",
     "The headline and readiness badge follow your files, builds and checks. The bright button takes the next step; the four metrics show the current position."),
    (["pdx_card1"], "Put the files on the desk",
     "Scan a folder or pick files. Each is recognised by its columns, not its name. Click a row to choose "
     "the file for that slot."),
    (["pdx_card2"], "Build what you need",
     "Switch frameworks on or off, then build. Every sheet comes from a row on Pivot config, and every row "
     "can be changed."),
    (["pdx_card3"], "Review every flagged check",
     "Compare outputs with controls 3 and 6. The matrix separates breaks from checks needing review. "
     "Open Reconciliation and Activity for the differences and scope."),
    (["pdx_rb"], "Open what you built",
     "Every workbook a build writes is listed here, newest first. Open one straight from the Desk, or the "
     "whole folder."),
    (["pdx_nav"], "Everything is one click away",
     "The bar goes to every sheet, as do Ctrl+Shift+D, F, P, R and A. Excel view brings the ribbon back. "
     "F1 plays this tour again."),
]


def union_of(shapes, names):
    got = [s for s in shapes if s.name in names]
    x0 = min(s.x for s in got)
    y0 = min(s.y for s in got)
    x1 = max(s.x + s.w for s in got)
    y1 = max(s.y + s.h for s in got)
    return x0, y0, x1 - x0, y1 - y0


def tour_place(tx, ty, tw, th):
    """The ring around a target, the card beside it, and the four veils.
    Beside it on the right if there is room, else on the left, else below,
    else above; anything in the app bar gets its card below."""
    p = TOUR_PAD
    rx, ry, rw, rh = tx - p, ty - p, tw + 2 * p, th + 2 * p

    def clamp(v, lo, hi):
        return max(lo, min(v, hi))

    if ry + rh < 60:
        kx, ky = clamp(rx, 12, W - TOUR_W - 12), ry + rh + TOUR_GAP
    elif rx + rw + TOUR_GAP + TOUR_W <= W - 12:
        kx, ky = rx + rw + TOUR_GAP, clamp(ry, 60, H - TOUR_H - 12)
    elif rx - TOUR_GAP - TOUR_W >= 12:
        kx, ky = rx - TOUR_GAP - TOUR_W, clamp(ry, 60, H - TOUR_H - 12)
    elif ry + rh + TOUR_GAP + TOUR_H <= H - 12:
        kx, ky = clamp(rx, 12, W - TOUR_W - 12), ry + rh + TOUR_GAP
    else:
        kx, ky = clamp(rx, 12, W - TOUR_W - 12), ry - TOUR_GAP - TOUR_H
    far = 800      # past the canvas, for a window wider than the design
    veils = [(0, 0, W + far, max(0.75, ry)),                       # above
             (0, ry + rh, W + far, H + far - ry - rh),            # below
             (0, ry, max(0.75, rx), rh),                            # left
             (rx + rw, ry, W + far - rx - rw, rh)]                  # right
    return (rx, ry, rw, rh), (kx, ky), veils


# ---------------------------------------------------------------------------
#  States
# ---------------------------------------------------------------------------
def empty_state(today="SUNDAY  ·  27 SEPTEMBER 2026", greeting="Your ALM workbench."):
    return {
        "version": "3.2",
        "asof": "NO DATA LOADED",
        "view_label": "Excel view",
        "date": today,
        "greeting": greeting,
        "lede": "Start with LCR, NSFR and maturity ladder outputs, plus controls 3 and 6. "
                "Scan their folder or choose individual files.",
        "hero_status": "START WITH SOURCE FILES",
        "hero_level": "IDLE",
        "cta": "Scan a folder",
        "cta2": "Pick files",
        "kpis": [
            {"key": "files", "label": "FILES", "value": "0", "sub": "of 5 on the desk", "segments": 5, "filled": 0},
            {"key": "fw", "label": "FRAMEWORKS", "value": "0", "sub": "of 3 ready to pivot", "segments": 3, "filled": 0},
            {"key": "ctl", "label": "CONTROLS", "value": "0", "sub": "of 2 control reports", "segments": 2, "filled": 0},
            {"key": "recon", "label": "RECONCILIATION", "value": "—", "sub": "not run yet", "segments": 1, "filled": 0},
        ],
        "counts": ["0 OF 5 CONNECTED", "AWAITING FRAMEWORK OUTPUTS", "ADD OUTPUTS + CONTROLS"],
        "slots": [
            {"label": "LCR output", "level": "EMPTY", "meta": "Not added  ·  click to choose", "meta_color": TX_4},
            {"label": "NSFR output", "level": "EMPTY", "meta": "Not added  ·  click to choose", "meta_color": TX_4},
            {"label": "Maturity ladder output", "level": "EMPTY", "meta": "Not added  ·  click to choose", "meta_color": TX_4},
            {"label": "Control report 3", "level": "EMPTY", "meta": "Not added  ·  click to choose", "meta_color": TX_4},
            {"label": "Control report 6", "level": "EMPTY", "meta": "Not added  ·  click to choose", "meta_color": TX_4},
        ],
        "frameworks": [
            {"label": "LCR", "state": "na", "meta": "Add an LCR output first"},
            {"label": "NSFR", "state": "na", "meta": "Add an NSFR output first"},
            {"label": "Maturity ladder", "state": "na", "meta": "Add a maturity ladder output first"},
        ],
        "last_build": "Nothing built yet.",
        "config_link": "4 pivots  →",
        "build_label": "Build pivots",
        "build_kind": "off",
        "open_kind": "off",
        "controls": [
            {"label": "Control report 3", "level": "EMPTY", "meta": "by COA  ·  not added"},
            {"label": "Control report 6", "level": "EMPTY", "meta": "by account  ·  not added"},
        ],
        "recon_when": "NOT RUN YET",
        "matrix": [["—", "—", "—"], ["—", "—", "—"]],
        "verdict_level": "IDLE",
        "verdict_title": "NOT RUN",
        "verdict_value": "—",
        "recon_kind": "off",
        "activity": [],
        "toast": None,
        "busy": False,
        "macros_banner": True,
        "builds": [],
        "tour": None,
    }


def showcase_state():
    s = empty_state()
    s.update({
        "asof": "DATA AS OF 30 NOV 2025",
        "greeting": "Bring the desk up to date.",
        "lede": "Control report 6 is no longer at its saved location. Choose it again, then rerun "
                "reconciliation against the current inputs.",
        "hero_status": "SOURCE FILE NEEDS ATTENTION",
        "hero_level": "WARN",
        "cta": "Find Control report 6",
        "cta2": "Pick files",
        "kpis": [
            {"key": "files", "label": "FILES", "value": "4", "sub": "of 5 on the desk", "segments": 5, "filled": 4},
            {"key": "fw", "label": "FRAMEWORKS", "value": "3", "sub": "of 3 ready to pivot", "segments": 3, "filled": 3},
            {"key": "ctl", "label": "CONTROLS", "value": "1", "sub": "of 2 control reports", "segments": 2, "filled": 1},
            {"key": "recon", "label": "RECONCILIATION", "value": "—", "sub": "inputs changed · rerun", "segments": 1,
             "filled": 1, "color": WARN, "seg_color": WARN},
        ],
        "counts": ["4 OF 5 CONNECTED", "3 OF 3 READY", "1 OF 2 CONTROLS CONNECTED"],
        "slots": [
            {"label": "LCR output", "level": "LOADED", "meta": "LCR_Output_30Nov.xlsx  ·  412,806 rows"},
            {"label": "NSFR output", "level": "LOADED", "meta": "NSFR_Output_30Nov.xlsx  ·  388,120 rows"},
            {"label": "Maturity ladder output", "level": "LOADED", "meta": "ML_Output_30Nov.xlsx  ·  506,332 rows"},
            {"label": "Control report 3", "level": "LOADED", "meta": "Control_Report_3.xlsx  ·  664 COAs"},
            {"label": "Control report 6", "level": "MISSING", "meta": "Moved or renamed  ·  click to find it",
             "meta_color": BAD},
        ],
        "frameworks": [
            {"label": "LCR", "state": "on", "meta": "412,806 rows  ·  as of 30 Nov 2025"},
            {"label": "NSFR", "state": "on", "meta": "388,120 rows  ·  as of 30 Nov 2025"},
            {"label": "Maturity ladder", "state": "on", "meta": "506,332 rows  ·  as of 30 Nov 2025"},
        ],
        "last_build": "Last built 26 Sep 14:06  ·  9 workbooks",
        "build_label": "Build 3 frameworks",
        "build_kind": "soft",
        "open_kind": "ghost",
        "controls": [
            {"label": "Control report 3", "level": "LOADED", "meta": "by COA  ·  664 keys"},
            {"label": "Control report 6", "level": "MISSING", "meta": "by account  ·  file moved"},
        ],
        "recon_when": "INPUTS CHANGED  ·  RERUN",
        "matrix": [["—", "—", "—"], ["—", "—", "—"]],
        "verdict_level": "WARN",
        "verdict_title": "RERUN",
        "verdict_value": "Required",
        "recon_kind": "soft",
        "activity": [
            {"when": "26 Sep 14:22", "level": "Break", "stage": "Recon",
             "msg": "Control 3 vs NSFR - 66 shared key(s) differ, gross 1,204,331,902, net 3,118 - the net is small because the differences offset."},
            {"when": "26 Sep 14:22", "level": "OK", "stage": "Recon",
             "msg": "Control 3 vs LCR - 664 key(s) on both sides, covering 99.82% of the output's balances. Every shared key agrees."},
            {"when": "26 Sep 14:21", "level": "OK", "stage": "Recon",
             "msg": "Control 3 vs Maturity ladder - 664 key(s) on both sides. Every shared key agrees."},
            {"when": "26 Sep 14:06", "level": "OK", "stage": "Pivots",
             "msg": "Maturity ladder - 506,332 row(s) staged, 7 workbooks (one per currency), 214 sheets, 58.9s."},
        ],
        "builds": [
            {"fw": "LADDER", "name": "Maturity Ladder - Egyptian Pound.xlsx", "meta": "26 Sep 14:06  ·  38 sheets"},
            {"fw": "LADDER", "name": "Maturity Ladder - US Dollar.xlsx", "meta": "26 Sep 14:06  ·  31 sheets"},
            {"fw": "NSFR", "name": "NSFR.xlsx", "meta": "26 Sep 14:05  ·  24 sheets"},
        ],
        "toast": {"level": "OK", "text": "9 workbooks written to D:\\ALM\\Pivots\\30 Nov, one per currency for the ladder. Each opens on its Start here sheet."},
        "macros_banner": False,
    })
    return s
