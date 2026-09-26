"""
The Desk, described once.

desk(state) returns every shape on the Desk for a given state of the desk.
The build calls it with the state the workbook ships in; the preview calls
it with the same state and with a busy, populated one. The VBA module
modPD_Desk repaints the same shapes by name at runtime, so the names here
are a contract: the build checks that every name the code paints exists.

Layout, in points, on a 1120 x 630 canvas:

    0    app bar
    68   hero: greeting, the one-sentence state, the next action, 4 tiles
    236  three cards - Add files, Build pivots, Reconcile
    524  recent activity
    612  footer
"""

from __future__ import annotations

import os

import assets
from design_tokens import *  # noqa: F401,F403
from shapes import Body, Glow, Line, Linear, Para, Run, Shadow, Shape, Solid, T

ASSETS = os.path.join(assets.OUT)

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
CARD_SHADOW = Shadow(18, 6, 90, "000000", 0.45)


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


def desk(st):
    S = []
    add = S.extend

    # --- app bar --------------------------------------------------------------
    S.append(Shape("pdx_bar", 0, 0, W, 52, fill=Solid(INK), descr="PivotDesk app bar"))
    S.append(Shape("pdx_bar_rule", 0, 51.25, W, 0.75,
                   fill=Linear(0, [(0, EM[900], 1), (0.5, EM[600], 1), (1, EM[900], 1)]), descr="rule"))
    S.append(pic("pdx_logo", MARGIN, 12, 28, 28, os.path.join(ASSETS, "logo"), macro="PD_GoHome",
                 descr="PivotDesk", svg=True))
    S.append(text("pdx_wordmark", 64, 10.5, 150, 20,
                  Para([Run("Pivot", 13, TX_1, F_SEMI), Run("Desk", 13, EM[300], F_SEMI)]),
                  descr="PivotDesk"))
    S.append(text("pdx_bank", 64.5, 29, 170, 12, T("MIDBANK  ·  CAIRO", 6.5, TX_3, F_SEMI, spc=1.6),
                  descr="MIDBANK Cairo"))

    # navigation
    nav_x, nav_y = 236, 13
    items = [("desk", "Desk", "spark", 66, "PD_GoHome"), ("files", "Files", "files", 64, "PD_GoFiles"),
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

    # right side: as-of chip, Excel view, Reset
    S.append(Shape("pdx_asof_chip", 748, 13, 172, 26, "roundRect", 13, fill=Solid("0B1310"),
                   line=Line(0.75, HAIR, 1), descr="data as of"))
    S.append(pic("pdx_asof_ic", 759, 20, 12, 12, icon("calendar", "em"), descr="as of"))
    S.append(text("pdx_asof", 776, 13, 140, 26, T(st["asof"], 7, TX_2, F_SEMI, spc=0.8), anchor="ctr",
                  wrap=False, descr="data as of"))
    add(button("pdx_btn_view", 928, 13, 86, st["view_label"], "ghost", "PD_ToggleAppView",
               icon("expand", "t2"), h=26, size=8, descr="Show or hide Excel's ribbon"))
    add(button("pdx_btn_reset", 1020, 13, 72, "Reset", "ghost", "PD_Reset", icon("reset", "t2"), h=26, size=8,
               descr="Clear loaded files and results"))

    # --- hero -----------------------------------------------------------------
    HX, HY, HW, HH = MARGIN, 68, W - 2 * MARGIN, 152
    S.append(pic("pdx_hero_art", HX, HY, HW, HH, os.path.join(ASSETS, "hero"), svg=False, descr="hero"))
    S.append(text("pdx_hero_date", HX + 28, HY + 22, 420, 12, T(st["date"], 7, EM[300], F_SEMI, spc=1.8),
                  wrap=False, descr="today"))
    S.append(text("pdx_hero_greet", HX + 26, HY + 33, 540, 40, T(st["greeting"], 26, TX_1, F_LIGHT),
                  wrap=False, descr="greeting"))
    S.append(text("pdx_hero_lede", HX + 28, HY + 74, 520, 32, T(st["lede"], 9.5, TX_2, F_BODY), descr="where the desk stands"))
    add(button("pdx_cta", HX + 28, HY + 112, 188, st["cta"], "primary", "PD_NextAction", icon("arrow", "ink"),
               descr="The next thing to do"))
    add(button("pdx_cta2", HX + 224, HY + 112, 112, st["cta2"], "ghost", "PD_LoadFiles", descr="Pick files"))

    # KPI tiles
    tw, tg = 102, 10
    tx0 = HX + HW - 24 - (4 * tw + 3 * tg)
    kpis = st["kpis"]
    for i, k in enumerate(kpis):
        x = tx0 + i * (tw + tg)
        y = HY + 20
        nm = "pdx_kpi_%s" % k["key"]
        S.append(Shape(nm, x, y, tw, 112, "roundRect", 10, fill=Solid("04100C", 0.78),
                       line=Line(0.75, EM[300], 0.16), descr=k["label"]))
        S.append(text(nm + "_label", x + 12, y + 12, tw - 20, 10, T(k["label"], 6.5, TX_2, F_SEMI, spc=1.2),
                      wrap=False, descr=k["label"]))
        S.append(text(nm + "_value", x + 11, y + 22, tw - 20, 36, T(k["value"], 24, k.get("color", TX_1), F_LIGHT),
                      wrap=False, descr=k["label"] + " value"))
        S.append(text(nm + "_sub", x + 12, y + 60, tw - 22, 24, T(k["sub"], 7, TX_2, F_BODY), descr=k["label"]))
        n = k["segments"]
        gap = 3
        sw = (tw - 24 - (n - 1) * gap) / n
        for j in range(n):
            on = j < k["filled"]
            S.append(Shape("%s_seg%d" % (nm, j + 1), x + 12 + j * (sw + gap), y + 96, sw, 4, "roundRect", 2,
                           fill=Solid(k.get("seg_color", EM[400])) if on else Solid("FFFFFF", 0.10),
                           descr="meter"))

    # --- the three cards ------------------------------------------------------------
    cards = [
        ("1", "01", "Add files", "Point at a folder or pick files. Each is recognised by the columns it "
         "carries, never its name.", "upload"),
        ("2", "02", "Build pivots", "One workbook per framework: the Output, the Balance sheet and a sheet "
         "per rule, all live PivotTables.", "pivot"),
        ("3", "03", "Reconcile", "Outputs against control reports 3 and 6. Scope is settled before any "
         "difference is called a break.", "balance"),
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
                      T(st["counts"][int(k) - 1], 6.5, TX_3, F_SEMI, spc=1.2), anchor="ctr", wrap=False,
                      descr="progress"))
        S.append(Shape("pdx_card%s_icon_tile" % k, x + CARD_W - 52, y + 18, 32, 32, "roundRect", 9,
                       fill=Linear(90, [(0, EM[900], 1), (1, EM[950], 1)]), line=Line(0.75, EM[800], 1),
                       descr=title))
        S.append(pic("pdx_card%s_icon" % k, x + CARD_W - 44, y + 26, 16, 16, icon(ic, "em"), descr=title))
        S.append(text("pdx_card%s_title" % k, x + 20, y + 42, 250, 22, T(title, 14, TX_1, F_SEMI), wrap=False,
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
    S.append(text("pdx_c2_last", x + 20, y + 211, CARD_W - 40, 12, T(st["last_build"], 7, TX_3, F_BODY),
                  wrap=False, clip=True, descr="last build"))
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

    # --- recent activity --------------------------------------------------------
    ay = 524
    S.append(Shape("pdx_act", MARGIN, ay, W - 2 * MARGIN, 80, "roundRect", 14, fill=CARD_FILL, line=CARD_LINE,
                   shadow=CARD_SHADOW, descr="recent activity"))
    S.append(pic("pdx_act_ic", MARGIN + 20, ay + 12, 12, 12, icon("pulse", "em"), descr="activity"))
    S.append(text("pdx_act_title", MARGIN + 38, ay + 9, 200, 18, T("Recent activity", 9.5, TX_1, F_SEMI),
                  anchor="ctr", wrap=False, descr="Recent activity"))
    S.append(text("pdx_act_all", W - MARGIN - 130, ay + 9, 110, 18, T("All activity  →", 8, EM[300], F_SEMI, align="r"),
                  anchor="ctr", wrap=False, macro="PD_GoLog", descr="Open the Activity sheet"))
    acts = st["activity"]
    for a in range(3):
        ry = ay + 31 + a * 15.5
        e = acts[a] if a < len(acts) else None
        fg, bg = lv(e["level"]) if e else lv("IDLE")
        S.append(text("pdx_act%d_when" % (a + 1), MARGIN + 20, ry, 78, 14, T(e["when"] if e else "", 7.5, TX_3, F_MONO),
                      anchor="ctr", wrap=False, hidden=e is None, descr="when"))
        S.append(pill("pdx_act%d_lvl" % (a + 1), MARGIN + 100, ry + 1.5, 46, 11, (e["level"] if e else "").upper(),
                      fg, bg, size=6, border=Line(0.75, fg, 0.35), descr="level"))
        S[-1].hidden = e is None
        S.append(text("pdx_act%d_stage" % (a + 1), MARGIN + 156, ry, 70, 14, T(e["stage"] if e else "", 8, TX_2, F_SEMI),
                      anchor="ctr", wrap=False, hidden=e is None, descr="stage"))
        S.append(text("pdx_act%d_msg" % (a + 1), MARGIN + 228, ry, W - 2 * MARGIN - 248, 14,
                      T(e["msg"] if e else "", 8, TX_1, F_BODY), anchor="ctr", wrap=False, clip=True,
                      hidden=e is None, descr="what happened"))
    S.append(pic("pdx_act_art", W - MARGIN - 250, ay + 1, 240, 78, os.path.join(ASSETS, "lattice-medallion"),
                 svg=False, hidden=bool(acts), descr="pattern"))
    S.append(text("pdx_act_empty", MARGIN + 20, ay + 34, 700, 36,
                  T("Nothing has happened on this desk yet. Add files and every step will be written here, newest first.",
                    8.5, TX_3, F_BODY), anchor="ctr", hidden=bool(acts), descr="no activity yet"))

    # --- footer -------------------------------------------------------------
    S.append(text("pdx_foot", MARGIN, 612, 700, 12,
                  T("PivotDesk %s  ·  Every workbook it writes is live PivotTables on one cache — drag a field, "
                    "add a slicer, double-click a total." % st["version"], 7, TX_3, F_BODY), wrap=False,
                  descr="about"))
    S.append(text("pdx_keys", W - MARGIN - 380, 612, 380, 12,
                  T("Ctrl+Shift+  D Desk  ·  F Files  ·  R Reconciliation  ·  A Activity", 7, TX_3, F_BODY, align="r"),
                  wrap=False, descr="keyboard shortcuts"))

    # --- overlays: busy, toast, macros-off ---------------------------------------
    S.append(Shape("pdx_busy", MARGIN, CARD_Y, W - 2 * MARGIN, 368, "roundRect", 14, fill=Solid(CANVAS, 0.82),
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
    return S


# ---------------------------------------------------------------------------
#  States
# ---------------------------------------------------------------------------
def empty_state(today="SATURDAY  ·  26 SEPTEMBER 2026", greeting="Good evening."):
    return {
        "version": "2.0",
        "asof": "NO DATA LOADED",
        "view_label": "Excel view",
        "date": today,
        "greeting": greeting,
        "lede": "Nothing is on the desk yet. Point PivotDesk at the folder holding your framework outputs "
                "and control reports 3 and 6 - it works out which file is which.",
        "cta": "Scan a folder",
        "cta2": "Pick files",
        "kpis": [
            {"key": "files", "label": "FILES", "value": "0", "sub": "of 5 on the desk", "segments": 5, "filled": 0},
            {"key": "fw", "label": "FRAMEWORKS", "value": "0", "sub": "of 3 ready to pivot", "segments": 3, "filled": 0},
            {"key": "ctl", "label": "CONTROLS", "value": "0", "sub": "of 2 control reports", "segments": 2, "filled": 0},
            {"key": "recon", "label": "RECONCILIATION", "value": "—", "sub": "not run yet", "segments": 1, "filled": 0},
        ],
        "counts": ["0 OF 5 IN PLACE", "NOTHING TO BUILD YET", "NEEDS AN OUTPUT AND A CONTROL"],
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
    }


def showcase_state():
    s = empty_state()
    s.update({
        "asof": "DATA AS OF 30 NOV 2025",
        "lede": "Four of five files are on the desk. LCR and NSFR are built and reconciled against control "
                "report 3 - one comparison broke. Control report 6 has moved; add it again to finish.",
        "cta": "Open the breaks",
        "cta2": "Pick files",
        "kpis": [
            {"key": "files", "label": "FILES", "value": "4", "sub": "of 5 on the desk", "segments": 5, "filled": 4},
            {"key": "fw", "label": "FRAMEWORKS", "value": "2", "sub": "of 3 ready to pivot", "segments": 3, "filled": 2},
            {"key": "ctl", "label": "CONTROLS", "value": "1", "sub": "of 2 control reports", "segments": 2, "filled": 1},
            {"key": "recon", "label": "RECONCILIATION", "value": "1", "sub": "break in 2 comparisons", "segments": 1,
             "filled": 1, "color": BAD, "seg_color": BAD},
        ],
        "counts": ["4 OF 5 IN PLACE", "2 OF 3 READY", "1 OF 2 CONTROLS IN PLACE"],
        "slots": [
            {"label": "LCR output", "level": "LOADED", "meta": "LCR_Output_30Nov.xlsx  ·  412,806 rows"},
            {"label": "NSFR output", "level": "LOADED", "meta": "NSFR_Output_30Nov.xlsx  ·  388,120 rows"},
            {"label": "Maturity ladder output", "level": "EMPTY", "meta": "Not added  ·  click to choose", "meta_color": TX_4},
            {"label": "Control report 3", "level": "LOADED", "meta": "Control_Report_3.xlsx  ·  664 COAs"},
            {"label": "Control report 6", "level": "MISSING", "meta": "Moved or renamed  ·  click to find it",
             "meta_color": BAD},
        ],
        "frameworks": [
            {"label": "LCR", "state": "on", "meta": "412,806 rows  ·  as of 30 Nov 2025"},
            {"label": "NSFR", "state": "on", "meta": "388,120 rows  ·  as of 30 Nov 2025"},
            {"label": "Maturity ladder", "state": "na", "meta": "Add a maturity ladder output first"},
        ],
        "last_build": "Last built 26 Sep 14:05  ·  2 workbooks  ·  D:\\ALM\\Pivots\\30 Nov",
        "build_label": "Build 2 workbooks",
        "build_kind": "soft",
        "open_kind": "ghost",
        "controls": [
            {"label": "Control report 3", "level": "LOADED", "meta": "by COA  ·  664 keys"},
            {"label": "Control report 6", "level": "MISSING", "meta": "by account  ·  file moved"},
        ],
        "recon_when": "LAST RUN  ·  26 SEP 14:22",
        "matrix": [["OK", "BREAK", "—"], ["—", "—", "—"]],
        "verdict_level": "BREAK",
        "verdict_title": "BREAKS",
        "verdict_value": "1 of 2",
        "recon_kind": "soft",
        "activity": [
            {"when": "26 Sep 14:22", "level": "Break", "stage": "Recon",
             "msg": "Control 3 vs NSFR - 66 shared key(s) differ, gross 1,204,331,902, net 3,118 - the net is small because the differences offset."},
            {"when": "26 Sep 14:22", "level": "OK", "stage": "Recon",
             "msg": "Control 3 vs LCR - 664 key(s) on both sides, covering 99.82% of the output's balances. Every shared key agrees."},
            {"when": "26 Sep 14:05", "level": "OK", "stage": "Pivots",
             "msg": "412,806 row(s) staged, 46 pivot sheet(s), 38.4s."},
        ],
        "toast": {"level": "OK", "text": "2 workbooks written to D:\\ALM\\Pivots\\30 Nov. Each opens on its Start here sheet."},
        "macros_banner": False,
    })
    return s
