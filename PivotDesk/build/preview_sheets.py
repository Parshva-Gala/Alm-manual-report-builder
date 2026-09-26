"""
Previews of the sheets the VBA paints - Files, Reconciliation, Activity -
and of a workbook it builds, drawn as HTML with the same rules modPD_Theme
and modPD_Pivot apply (fonts, fills, borders, row heights, the app bar).

These are pictures of what the code does, for review off Windows. The
numbers in the built-workbook preview are computed from the LCR sample the
bank supplied, the same way the Balance sheet pivot groups them.
"""

from __future__ import annotations

import collections
import html
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import assets  # noqa: E402

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "preview")
SAMPLE = os.environ.get("LCR_SAMPLE", "")

C = dict(INK="000000", CANVAS="060B09", ROW="0C1512", ROW_ALT="111F19", LINE="1A2922", LINE2="243A31",
         WELL="0B1310", TX1="F2F7F4", TX2="A9BDB3", TX3="7E9388", TX4="4A5E55", SOFT="C2EBDA", BRAND="009060",
         DEEP="006141", E900="003323", E950="001D14", M300="4FC79C", M200="8FDBBE",
         OK="2FC48D", OK_BG="0D2A20", WARN="F2B544", WARN_BG="2A2310", BAD="FF6B5E", BAD_BG="2E1614",
         IDLE="7E9388", IDLE_BG="121D19",
         # the light canvas, still used by the built-workbook previews until they move over
         PAPER="FFFFFF", MIST="F4F7F5", HAIR="E2E9E5", MUTED="5F7068", BODY="18241F", OK_TX="0B6B47",
         TINT="E6F6EF")
LEVEL = {"OK": ("OK", "OK_BG"), "LOADED": ("OK", "OK_BG"), "BREAK": ("BAD", "BAD_BG"),
         "MISSING": ("BAD", "BAD_BG"), "CHECK": ("WARN", "WARN_BG"), "EMPTY": ("IDLE", "IDLE_BG"),
         "IDLE": ("IDLE", "IDLE_BG"), "OFF": ("IDLE", "IDLE_BG")}

PX_PER_CHAR = 8          # Segoe UI 9.5 on a 7 px Calibri grid, near enough
PT = 4 / 3
LOGO = os.path.join(ROOT, "design", "brand", "avati-logo.png")
LOGO_RATIO = 785 / 205


def colw(chars):
    return round(chars * 7 + 5)


def lv(word):
    fg, bg = LEVEL.get(word.upper(), LEVEL["IDLE"])
    return C[fg], C[bg]


def px(pt):
    return round(pt * PT, 1)


def page(title, body, width):
    return ('<!doctype html><html><head><meta charset="utf-8"><title>%s</title><style>'
            'html,body{margin:0;background:#%s;font-family:Selawik;color:#%s}'
            '.wrap{width:%dpx;overflow:hidden;background:#%s}'
            'table{border-collapse:collapse;table-layout:fixed}'
            'td{padding:0 0 0 9px;white-space:nowrap;overflow:hidden;text-overflow:clip;font-size:%.1fpx}'
            '.pill{display:inline-flex;align-items:center;justify-content:center;border-radius:999px;'
            'font-family:"Selawik Semibold";font-size:11.3px;box-sizing:border-box;white-space:nowrap}'
            '</style></head><body><div class="wrap">%s</div></body></html>' % (
                title, C["CANVAS"], C["TX1"], width, C["CANVAS"], 9.5 * PT, body))


NAV = [("Desk", "desk", 52), ("Files", "files", 52), ("Reports", "reports", 70),
       ("Reconciliation", "recon", 102), ("Activity", "log", 66)]
TABS = [("Pivots", "Pivot config"), ("Charts", "Chart config"), ("Workbooks", "Workbooks"),
        ("Fields", "Pivot fields"), ("Gallery", "Report gallery")]


def app_bar(section):
    """modPD_Theme.Rail: the mark, ALM DESK over the bank, and the nav pills."""
    lw = 16 * LOGO_RATIO
    nav = "".join(
        "<span class='pill' style='width:%.1fpx;height:%.1fpx;margin-right:%.1fpx;%s'>%s</span>" % (
            px(w), px(22), px(2),
            ("background:#%s;box-shadow:inset 0 0 0 1px #%s;color:#%s" % (C["E900"], C["DEEP"], C["TX1"])
             if key == section else "color:#%s" % C["TX2"]), label)
        for label, key, w in NAV)
    return ("<div style='height:%.1fpx;background:#000;border-bottom:1px solid #%s;display:flex;align-items:center;"
            "position:relative'>"
            "<img src='file://%s' style='position:absolute;left:%.1fpx;height:%.1fpx'>"
            "<div style='position:absolute;left:%.1fpx;width:1px;height:%.1fpx;background:#%s'></div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "letter-spacing:1.6px;color:#%s'>ALM DESK</div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "letter-spacing:1.4px;color:#%s'>MIDBANK &nbsp;&#183;&nbsp; CAIRO</div>"
            "<div style='position:absolute;left:%.1fpx;height:%.1fpx;border-radius:999px;background:#%s;"
            "box-shadow:inset 0 0 0 1px #%s;display:flex;align-items:center;padding:0 %.1fpx'>%s</div></div>") % (
        px(44), C["DEEP"], LOGO, px(16), px(16), px(16 + lw + 10), px(16), C["LINE2"],
        px(16 + lw + 22), px(10), px(7.5), C["TX1"], px(16 + lw + 22), px(23), px(6), C["TX3"],
        px(16 + lw + 22 + 104), px(26), C["WELL"], C["LINE"], px(3), nav)


def toolbar(actions, tab=None):
    """Row 4: the Reports tabs, a hairline, then what this sheet can do."""
    h = []
    if tab:
        tabs = "".join(
            "<span class='pill' style='width:%.1fpx;height:%.1fpx;margin-right:%.1fpx;%s'>%s</span>" % (
                px(74), px(22), px(2),
                ("background:#%s;box-shadow:inset 0 0 0 1px #%s;color:#%s" % (C["E900"], C["DEEP"], C["TX1"])
                 if label == tab else "color:#%s" % C["TX2"]), label) for label, _ in TABS)
        h.append("<div style='height:%.1fpx;border-radius:999px;background:#%s;box-shadow:inset 0 0 0 1px #%s;"
                 "display:flex;align-items:center;padding:0 %.1fpx;margin-right:%.1fpx'>%s</div>"
                 "<div style='width:1px;height:%.1fpx;background:#%s;margin-right:%.1fpx'></div>" % (
                     px(26), C["WELL"], C["LINE"], px(3), px(10), tabs, px(16), C["LINE2"], px(12)))
    for a in actions:
        h.append("<span class='pill' style='height:%.1fpx;padding:0 %.1fpx;margin-right:%.1fpx;background:#%s;"
                 "box-shadow:inset 0 0 0 1px #%s;color:#%s'>%s</span>" % (
                     px(24), px(14), px(6), C["E950"], C["DEEP"], C["SOFT"], html.escape(a)))
    return "<div style='height:%.1fpx;display:flex;align-items:center;padding-left:%.1fpx'>%s</div>" % (
        px(34), px(14), "".join(h))


def chrome(section, title, about, status, level, actions, tab=None, overline=None, status_w=1040, band_row=None):
    """Rows 1 to 6 of every tool sheet: modPD_Theme.Dress, Rail, Toolbar and SetStatus."""
    overline = overline or "ALM DESK &nbsp;&#183;&nbsp; MIDBANK CAIRO"
    fg, bg = lv(level)
    title_html = ("<div style='height:%.1fpx;display:flex;flex-direction:column;justify-content:flex-end;"
                  "padding:0 0 %.1fpx %.1fpx;box-sizing:border-box'>"
                  "<div style='font-family:\"Selawik Semibold\";font-size:%.1fpx;letter-spacing:1.4px;color:#%s'>%s</div>"
                  "<div style='font-family:\"Selawik Light\";font-size:%.1fpx;color:#%s;line-height:1.1'>%s</div></div>"
                  "<div style='height:%.1fpx;padding-left:%.1fpx;font-size:%.1fpx;color:#%s'>%s</div>") % (
        px(44), px(1), px(14), px(7.5), C["M300"], overline, px(19), C["TX1"], html.escape(title),
        px(20), px(14), px(9), C["TX2"], html.escape(about))
    status_html = ("<div style='width:%dpx;height:%.1fpx;background:#%s;border-left:4px solid #%s;display:flex;"
                   "align-items:center;padding-left:12px;box-sizing:border-box;font-family:\"Selawik Semibold\";"
                   "font-size:%.1fpx;color:#%s'>&#9679;&nbsp;&nbsp;%s</div>") % (
        status_w, px(26), bg, fg, px(9), fg, html.escape(status))
    spacer = band_row if band_row is not None else "<div style='height:%.1fpx'></div>" % px(8)
    return app_bar(section) + title_html + toolbar(actions, tab) + status_html + spacer


def table(headers, widths, rows, verdict_col=None, mono_cols=(), muted_cols=(), right_cols=(), semi_cols=(),
          bars_col=None, band=True, pre_rows=""):
    """modPD_Theme.Head and DressTable, dark."""
    cols = "".join('<col style="width:%dpx">' % colw(w) for w in widths)
    h = ['<table>%s%s<tr style="height:%.1fpx;background:#000">' % (cols, pre_rows, px(28))]
    for t in headers:
        h.append('<td style="font-family:\'Selawik Semibold\';font-size:%.1fpx;color:#%s;border-bottom:3px solid #%s;'
                 'border-right:1px solid #%s">%s</td>' % (px(8.5), C["SOFT"], C["BRAND"], C["LINE"], html.escape(t)))
    h.append("</tr>")
    maxbar = max((abs(r[bars_col]) for r in rows if bars_col is not None and isinstance(r[bars_col], (int, float))),
                 default=1) or 1
    for i, r in enumerate(rows):
        bgc = C["ROW_ALT"] if (band and i % 2 == 1) else C["ROW"]
        h.append('<tr style="height:%.1fpx;background:#%s">' % (px(22), bgc))
        for j, v in enumerate(r):
            st = ["border-bottom:1px solid #%s" % C["LINE"]]
            txt = v
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                txt = "-" if v == 0 else ("(%s)" % format(abs(v), ",.0f") if v < 0 else format(v, ",.0f"))
                st.append("text-align:right;padding-right:9px")
                if v < 0:
                    st.append("color:#FF0000")
            if j == verdict_col and str(v):
                fg, bg = lv(str(v))
                st = ["border-bottom:1px solid #%s" % C["LINE"], "text-align:center", "background:#%s" % bg,
                      "color:#%s" % fg, "font-family:'Selawik Semibold'", "font-size:%.1fpx" % px(8.5)]
                txt = "&#9679;&nbsp;&nbsp;" + html.escape(str(v))
            else:
                txt = html.escape(str(txt))
            if j in mono_cols:
                st.append("font-family:'DejaVu Sans Mono';font-size:11px;color:#%s" % C["TX3"])
            if j in muted_cols:
                st.append("color:#%s" % C["TX3"])
            if j in semi_cols:
                st.append("font-family:'Selawik Semibold'")
            if j == bars_col and isinstance(v, (int, float)) and v:
                w = abs(v) / maxbar * 100
                st.append("background:linear-gradient(90deg, rgba(244,162,154,.55) %.1f%%, transparent %.1f%%)" % (w, w))
            h.append('<td style="%s">%s</td>' % (";".join(st), txt))
        h.append("</tr>")
    h.append("</table>")
    return "".join(h)


def files_sheet():
    widths = [30, 12, 54, 20, 11, 12, 13, 34, 56]
    rows = [
        ["LCR output", "Loaded", r"D:\ALM\30 Nov\LCR_Output_30Nov.xlsx", "Sheet1", 1, 412806, "30 Nov 2025",
         "CASHFLOW_AMOUNT_LCY_PRE_FACTOR  /  ...", "an ALM output for LCR, from its own framework column."],
        ["NSFR output", "Loaded", r"D:\ALM\30 Nov\LCR_Portfolio Segmentation.xlsx", "Sheet1", 1, 388120,
         "30 Nov 2025", "CASHFLOW_AMOUNT_LCY_PRE_FACTOR  /  ...", "an ALM output for NSFR, from its own framework column."],
        ["Maturity Ladder output", "Empty", "", "", "", "", "", "", ""],
        ["Control report 3  (by COA)", "Loaded", r"D:\ALM\30 Nov\Control_Report_3.xlsx", "Report", 3, "", "", "",
         "a control report - it carries GL_BALANCE_LCY."],
        ["Control report 6  (by account)", "Missing", r"D:\ALM\30 Nov\CR6_old.xlsx", "Report", 2, "", "", "",
         "The file is no longer at that path."],
    ]
    w = sum(colw(x) for x in widths)
    body = (chrome("files", "Files", "Every file is identified by the columns it carries, not by its name. Add them "
                   "from the Desk, or click a row here and press \"Use a file for this row\" to place one by hand.",
                   "4 of 5 file(s) loaded.  Ready to build pivots.", "OK",
                   ["Scan a folder", "Pick files", "Use a file for this row", "Clear this row"]) +
            table(["What", "Status", "File", "Sheet", "Header row", "Rows", "As of", "Amount field", "Note"], widths,
                  rows, verdict_col=1, muted_cols=(2, 8), semi_cols=(0,)))
    return page("Files", body, 1500)


def recon_sheet():
    widths = [12, 18, 26, 18, 18, 18, 13, 74]
    rows = [
        ["Control 3", "NSFR", "MBGL.1360", 812_441_902.0, 104_120_000.0, 708_321_902.0, "Check",
         "Both sides carry this key and the amounts differ."],
        ["Control 3", "NSFR", "MBGL.4010", -95_220_110.0, -380_120_000.0, 284_899_890.0, "Check",
         "Both sides carry this key and the amounts differ."],
        ["Control 3", "NSFR", "MBGL.6510", 0.0, 176_004_220.0, -176_004_220.0, "Break",
         "The control carries this and the output classified nothing against it. The whole amount is unclassified."],
        ["Control 3", "NSFR", "MBGL.2900", 42_118_004.0, 11_900_000.0, 30_218_004.0, "Check",
         "Both sides carry this key and the amounts differ."],
        ["Control 3", "NSFR", "MBGL.3180", 5_100_220.0, 1_880_000.0, 3_220_220.0, "Check",
         "Both sides carry this key and the amounts differ."],
    ]
    body = (chrome("recon", "Reconciliation", "Each loaded output against control report 3 (the ledger, by COA) and "
                   "control report 6 (the reporting balance, by account).",
                   "1 of 2 comparison(s) broke, the largest involving 1,204,331,902 LCY.   Scope is on Activity for "
                   "each - read it before treating a difference as an error.   (12.4s)", "Break", ["Reconcile now"],
                   status_w=1300) +
            table(["Control", "Framework", "Key", "Output", "Control", "Difference", "Verdict", "What it means"],
                  widths, rows, verdict_col=6, mono_cols=(2,), muted_cols=(7,), bars_col=5))
    return page("Reconciliation", body, 1500)


def log_sheet():
    widths = [20, 12, 14, 110, 40]
    rows = [
        ["26 Sep  14:22:05", "Break", "Recon", "1 shared key(s) differ ... the net is small because the differences "
         "offset, which is not agreement.", "Control 3 vs NSFR"],
        ["26 Sep  14:22:05", "OK", "Recon", "664 key(s) on both sides, covering 99.82% of the output's balances. "
         "Every shared key agrees.", "Control 3 vs LCR"],
        ["26 Sep  14:21:52", "OK", "Recon", "Control 3: 664 COA(s).", ""],
        ["26 Sep  14:05:41", "OK", "Pivots", "412,806 row(s) staged, 46 pivot sheet(s), 38.4s.", "LCR"],
        ["26 Sep  14:03:10", "Check", "Load", "Two files both look like OUTPUT NSFR. Keeping LCR_Portfolio "
         "Segmentation.xlsx; this one was NOT used.", "NSFR_copy.xlsx"],
        ["26 Sep  14:02:58", "OK", "Load", "Placed as OUTPUT LCR. an ALM output for LCR, from its own framework "
         "column.", "LCR_Output_30Nov.xlsx"],
    ]
    body = (chrome("log", "Activity", "What the desk did and why, newest first. A file that was not recognised says "
                   "here what was missing.", "The newest entry is at the top. The Desk shows the latest four.", "Idle",
                   ["Clear activity"], status_w=1300) +
            table(["When", "Level", "Stage", "What happened", "Which file"], widths, rows, verdict_col=1,
                  mono_cols=(0,), muted_cols=(4,), semi_cols=(2,), band=False))
    return page("Activity", body, 1500)


# ---------------------------------------------------------------------------
#  A built workbook, from the LCR sample
# ---------------------------------------------------------------------------
def balance_pivot():
    """The Balance sheet pivot of an LCR build of the sample: labels tidied as
    Pivot fields says (Line, Subline and COA name: drop codes + title case)."""
    import openpyxl
    import labels
    wb = openpyxl.load_workbook(SAMPLE, read_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    ix = {}
    for i, h in enumerate(hdr):
        ix.setdefault(h, i)
    rows = list(it)
    local = collections.Counter(r[ix["CURRENCY_NAME"]] for r in rows).most_common(1)[0][0]
    agg = collections.defaultdict(lambda: [0.0, 0.0])
    for r in rows:
        if not r[ix["BUCKET_DISPLAY_NAME"]]:
            continue                                    # the pivot's default: blanks out
        key = (r[ix["COA_BALANCESHEET_CATEGORY"]] or "(no type)",
               labels.tidy(r[ix["BALANCESHEET_LINE_NAME"]] or "(no line)", 2),
               labels.tidy(r[ix["BALANCESHEET_SUB_LINE_NAME"]] or "(no subline)", 2),
               labels.tidy(r[ix["COA_NAME"]] or "(no COA)", 2))
        side = 0 if r[ix["CURRENCY_NAME"]] == local else 1
        agg[key][side] += float(r[ix["CASHFLOW_AMOUNT_LCY_PRE_FACTOR"]] or 0)
    return local, sorted(agg.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2], kv[0][3]))


def fmt_num(v):
    if round(v) == 0:
        return "-"
    return "(%s)" % format(abs(v), ",.0f") if v < 0 else format(v, ",.0f")


def compact(v):
    """modPD_Pivot.CompactFormat / modPD_Util.Compact."""
    a = abs(v)
    if a >= 1e9:
        return "%.2f bn" % (v / 1e9)
    if a >= 1e6:
        return "%.1f m" % (v / 1e6)
    return format(v, ",.0f")


def book_bar(crumb, back=True, prev_next=True):
    """modPD_Theme.BookBar and modPD_Pivot.LinkSiblings."""
    lw = 15 * LOGO_RATIO
    x = 16 + lw + 22
    pills = ""
    if back:
        pills += ("<span class='pill' style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;"
                  "background:#%s;box-shadow:inset 0 0 0 1px #%s;color:#%s'>&#8249;&nbsp;&nbsp;Start here</span>") % (
            px(x + 190), px(8), px(96), px(24), C["E950"], C["DEEP"], C["SOFT"])
    if prev_next:
        for label, dx, w in (("&#8249;&nbsp;&nbsp;Previous", 104, 88), ("Next&nbsp;&nbsp;&#8250;", 198, 72)):
            pills += ("<span class='pill' style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;"
                      "box-shadow:inset 0 0 0 1px #%s;color:#%s'>%s</span>") % (
                px(x + 190 + dx), px(8), px(w), px(24), C["LINE2"], C["TX2"], label)
    return ("<div style='height:%.1fpx;background:#000;border-bottom:1px solid #%s;position:relative'>"
            "<img src='file://%s' style='position:absolute;left:%.1fpx;top:%.1fpx;height:%.1fpx'>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:1px;height:%.1fpx;background:#%s'></div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "letter-spacing:1.6px;color:#%s'>ALM DESK</div>"
            "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
            "letter-spacing:1.4px;color:#%s'>%s</div>%s</div>") % (
        px(40), C["DEEP"], LOGO, px(16), px(12.5), px(15), px(16 + lw + 10), px(12), px(16), C["LINE2"],
        px(x), px(8), px(7.5), C["TX1"], px(x), px(21), px(6), C["TX3"],
        html.escape(crumb.upper()).replace("  ", "&nbsp;&nbsp;"), pills)


def title_block(overline, title, about):
    return ("<div style='height:%.1fpx;display:flex;flex-direction:column;justify-content:flex-end;"
            "padding:0 0 %.1fpx %.1fpx;box-sizing:border-box'>"
            "<div style='font-family:\"Selawik Semibold\";font-size:%.1fpx;letter-spacing:1.4px;color:#%s'>%s</div>"
            "<div style='font-family:\"Selawik Light\";font-size:%.1fpx;color:#%s;line-height:1.1'>%s</div></div>"
            "<div style='height:%.1fpx;padding-left:%.1fpx;font-size:%.1fpx;color:#%s'>%s</div>") % (
        px(44), px(1), px(14), px(7.5), C["M300"], overline, px(19), C["TX1"], html.escape(title),
        px(20), px(14), px(9), C["TX2"], html.escape(about))


def tiles(items, w_pt=196, gap_pt=10, h_row=66, fixed_w=None):
    """modPD_Theme.Tile, in a row."""
    out = []
    for i, (label, value) in enumerate(items):
        w = fixed_w or w_pt
        size = 17
        if fixed_w and (w - 22) / (len(value) * 0.52) < size:
            size = max(10, int((w - 22) / (len(value) * 0.52)))
        out.append(("<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:#%s;"
                    "box-shadow:inset 0 0 0 1px #%s;border-radius:7px;box-sizing:border-box'>"
                    "<div style='position:absolute;left:0;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:#%s'></div>"
                    "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Semibold\";"
                    "font-size:%.1fpx;letter-spacing:1px;color:#%s'>%s</div>"
                    "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font-family:\"Selawik Light\";"
                    "font-size:%.1fpx;color:#%s;white-space:nowrap'>%s</div></div>") % (
            px(14 + i * (w + gap_pt)), px(8), px(w), px(50), C["ROW"], C["LINE2"], px(12), px(2.5), px(26),
            C["BRAND"], px(14), px(8), px(6.5), C["TX3"], label, px(13), px(19), px(size), C["TX1"],
            html.escape(value)))
    return "<div style='position:relative;height:%.1fpx'>%s</div>" % (px(h_row), "".join(out))


def slicers():
    """Dark, in the Avati Slicer style: chosen items emerald."""
    def one(title, items, sel):
        chips = "".join(
            "<div style='height:%.1fpx;border-radius:3px;margin:3px 0;padding-left:8px;font-size:11px;line-height:%.1fpx;"
            "background:#%s;color:#%s;white-space:nowrap;overflow:hidden'>%s</div>" % (
                px(15), px(15), C["DEEP"] if s_ else C["ROW_ALT"], C["TX1"] if s_ else C["TX2"], html.escape(t))
            for t, s_ in zip(items, sel))
        return ("<div style='width:%.1fpx;height:%.1fpx;background:#%s;box-shadow:inset 0 0 0 1px #%s;border-radius:4px;"
                "padding:6px 8px;box-sizing:border-box;margin-right:%.1fpx;display:inline-block;vertical-align:top;"
                "overflow:hidden'><div style='color:#%s;font-family:\"Selawik Semibold\";font-size:11.5px;"
                "margin-bottom:2px'>%s</div><div style='column-count:2;column-gap:6px'>%s</div></div>") % (
            px(156), px(62), C["ROW"], C["LINE2"], px(6), C["TX1"], title, chips)
    return ("<div style='padding:%.1fpx 0 0 %.1fpx;height:%.1fpx;box-sizing:border-box'>%s%s%s</div>" % (
        px(6), px(14), px(74),
        one("Type", ["Asset", "Liability", "Off Balance Sheet", "Shareholders Equity"], [1, 1, 1, 1]),
        one("LCY / FCY", ["LCY", "FCY"], [1, 1]),
        one("Rule category", ["OUTFLOW", "(no category)"], [1, 1])))


def pivot_sheet():
    local, items = balance_pivot()
    items = items[:18]
    widths = [22, 44, 48, 50, 16, 14, 16]
    tot = [sum(v[0] for _, v in items), sum(v[1] for _, v in items)]
    tot.append(tot[0] + tot[1])
    cols = "".join('<col style="width:%dpx">' % colw(w) for w in widths)
    hd = "font-family:\"Selawik Semibold\";color:#%s;background:#000" % C["SOFT"]
    head = ("<tr style='height:%.1fpx'><td colspan=4 style='background:#000;color:#%s;font-family:\"Selawik Semibold\"'>"
            "Pre-factor</td><td colspan=3 style='%s'>LCY / FCY&nbsp;&#9662;</td></tr>"
            "<tr style='height:%.1fpx'>%s</tr>") % (
        px(20), C["SOFT"], hd, px(22),
        "".join("<td style='%s;border-bottom:3px solid #%s;%s'>%s</td>" % (
            hd, C["BRAND"], "color:#%s" % C["TX1"] if h == "Total" else "", h + ("&nbsp;&#9662;" if i < 4 else ""))
            for i, h in enumerate(["Type", "Line", "Subline", "COA name", "LCY", "FCY", "Total"])))
    rows_html = []
    prev = [None, None, None]
    for i, (key, (lcy, fcy)) in enumerate(items):
        cells = []
        for j in range(4):
            show = key[j] if (j == 3 or key[j] != prev[j] or any(key[k] != prev[k] for k in range(j))) else ""
            cells.append(show)
        prev = list(key[:3])
        bg = C["ROW_ALT"] if i % 2 else C["ROW"]
        tds = "".join('<td style="border-bottom:1px solid #%s;%s">%s</td>' % (
            C["LINE"], "font-family:'Selawik Semibold'" if (j == 0 and c) else "", html.escape(str(c)))
            for j, c in enumerate(cells))
        for k, v in enumerate((lcy, fcy, lcy + fcy)):
            tds += '<td style="border-bottom:1px solid #%s;text-align:right;padding-right:9px;%s%s">%s</td>' % (
                C["LINE"], "color:#FF0000;" if v < -0.5 else "",
                "background:#16271F;font-family:'Selawik Semibold';" if k == 2 else "", fmt_num(v))
        rows_html.append('<tr style="height:%.1fpx;background:#%s">%s</tr>' % (px(20), bg, tds))
    total_row = ("<tr style='height:%.1fpx;background:#%s'><td colspan=4 style='font-family:\"Selawik Semibold\";"
                 "border-top:3px solid #%s'>Total</td>%s</tr>") % (
        px(20), C["E950"], C["BRAND"], "".join(
            "<td style='text-align:right;padding-right:9px;font-family:\"Selawik Semibold\";border-top:3px solid #%s;"
            "%s'>%s</td>" % (C["BRAND"], "color:#FF0000" if v < -0.5 else "", fmt_num(v)) for v in tot))
    filt = ("<table style='margin:0'>%s<tr style='height:%.1fpx'><td style='background:#%s;color:#%s;"
            "font-family:\"Selawik Semibold\"'>Bucket</td><td style='background:#%s'>(Multiple Items)&nbsp;&#9662;</td>"
            "</tr></table><div style='height:%.1fpx'></div>") % (
        "".join('<col style="width:%dpx">' % colw(w) for w in widths[:2]), px(20), C["ROW_ALT"], C["M300"], C["ROW"],
        px(10))
    crumb = "LCR  \u00b7  MIDBANK CAIRO"
    body = (book_bar(crumb) +
            title_block("LCR &nbsp;&#183;&nbsp; BALANCE SHEET", "Balance sheet",
                        "The same balances as the balance sheet reads them, down to the COA. Pre-factor only - what "
                        "the engine took in, before any weighting.") +
            tiles([("PRE-FACTOR &nbsp;&#183;&nbsp; TOTAL", compact(tot[2]))]) +
            "<div style='height:%.1fpx'></div>" % px(6) + slicers() + "<div style='height:%.1fpx'></div>" % px(8) +
            filt + "<table>%s%s%s%s</table>" % (cols, head, "".join(rows_html), total_row))
    return page("Balance sheet", body, 1500), local


def sample_facts():
    """What the Start here tiles and chart show for a build of the LCR sample:
    the staging pass's totals, and net pre-factor per bucket, LCY and FCY."""
    import openpyxl
    import tenor
    wb = openpyxl.load_workbook(SAMPLE, read_only=True)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    hdr = next(rows)
    ix = {}
    for i, h in enumerate(hdr):
        ix.setdefault(h, i)
    data = list(rows)
    local = collections.Counter(r[ix["CURRENCY_NAME"]] for r in data if r[ix["CURRENCY_NAME"]]).most_common(1)[0][0]
    gross = sum(abs(float(r[ix["CASHFLOW_AMOUNT_LCY_PRE_FACTOR"]] or 0)) for r in data)
    post = sum(abs(float(r[ix["CASHFLOW_AMOUNT_LCY_POST_FACTOR"]] or 0)) for r in data)
    net = sum(float(r[ix["CASHFLOW_AMOUNT_LCY_PRE_FACTOR"]] or 0) for r in data)
    by = collections.defaultdict(lambda: [0.0, 0.0])
    for r in data:
        b = r[ix["BUCKET_DISPLAY_NAME"]]
        if not b:
            continue                                    # "(no bucket)" is filtered out, as in the pivot
        by[b][0 if r[ix["CURRENCY_NAME"]] == local else 1] += float(r[ix["CASHFLOW_AMOUNT_LCY_PRE_FACTOR"]] or 0)
    buckets = [(b, by[b][0], by[b][1]) for b in tenor.bucket_order(list(by))]
    asof = next((r[ix["AS_OF_DATE"]] for r in data if r[ix["AS_OF_DATE"]]), None)
    return {"rows": len(data), "gross": gross, "post": post, "net": net, "factor": post / gross if gross else 0,
            "local": local, "asof": asof.strftime("%-d %b %Y") if asof else "", "buckets": buckets}


def guide_sheet(facts, made, title, built):
    wide_px = colw(38) + colw(110) - 18
    crumb = "%s  \u00b7  MIDBANK CAIRO" % title.upper()

    def section(text):
        return ("<div style='height:%.1fpx;width:%dpx;box-sizing:border-box;border-bottom:1px solid #%s;display:flex;"
                "align-items:flex-end;padding:0 0 5px 9px;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
                "letter-spacing:.6px;color:#%s'>%s</div>") % (px(30), colw(38) + colw(110), C["DEEP"], px(8),
                                                             C["M300"], text)
    items = [("ROWS STAGED", format(facts["rows"], ",")), ("PRE-FACTOR", compact(facts["gross"])),
             ("POST-FACTOR", compact(facts["post"])), ("SHEETS", str(len(made))),
             ("LOCAL CURRENCY", facts["local"]), ("DATA AS OF", facts["asof"])]
    tw = (wide_px / PT - 5 * 10) / 6
    status = ("<div style='width:1060px;height:%.1fpx;background:#%s;border-left:4px solid #%s;display:flex;"
              "align-items:center;padding-left:12px;box-sizing:border-box;font-family:\"Selawik Semibold\";"
              "font-size:%.1fpx;color:#%s'>&#9679;&nbsp;&nbsp;%s rows staged into one pivot cache.&nbsp; Every sheet "
              "below is a live PivotTable over it - drag a field, drop a slicer, drill a total.</div>") % (
        px(26), C["OK_BG"], C["OK"], px(9), C["OK"], format(facts["rows"], ","))
    tbl = table(["Sheet", "What is on it"], [38, 110], [[a_, b_] for a_, b_ in made], muted_cols=(1,))
    import re
    tbl = re.sub(r'(<tr style="height:[0-9.]+px;background:#[0-9A-F]+"><td style=")',
                 r"\1color:#%s;font-family:'Selawik Semibold';" % C["M300"], tbl)
    notes = [("Amounts", "Pre-factor and post-factor from CASHFLOW_AMOUNT_LCY_PRE_FACTOR  /  CASHFLOW_AMOUNT_LCY_POST_FACTOR."),
             ("Local currency", "\"%s\" - the currency on the most rows. Nothing in the extract says which is local, "
              "so check this before relying on the LCY / FCY split." % facts["local"]),
             ("Factor", "Post-factor divided by pre-factor, so it always agrees with the two figures beside it."),
             ("Labels", "Ledger names are shown without their codes and in title case where Pivot fields says so - "
              "1.07.00.MBGL.1360.LOANS TO CUSTOMERS reads Loans to Customers."),
             ("Blanks", "Rows with no bucket are excluded from the bucket filter by default. They are still in the "
              "data - clear the filter to see them.")]
    nrows = "".join("<div style='display:flex;padding:7px 0;font-size:12.7px'><div style='width:%dpx;padding-left:9px;"
                    "font-family:\"Selawik Semibold\";color:#%s'>%s</div><div style='width:%dpx;padding-left:9px;color:#%s;"
                    "white-space:normal'>%s</div></div>" % (colw(38) - 9, C["TX1"], a_, colw(110) - 9, C["TX2"],
                                                            html.escape(b_)) for a_, b_ in notes)
    body = (book_bar(crumb, back=False, prev_next=False) +
            title_block("START HERE &nbsp;&#183;&nbsp; MIDBANK CAIRO", title, built) +
            tiles(items, fixed_w=tw, h_row=68) + status + "<div style='height:%.1fpx'></div>" % px(12) +
            section("IN THIS BOOK &nbsp;&#183;&nbsp; %d SHEETS" % len(made)) + tbl +
            "<div style='height:%.1fpx'></div>" % px(8) + section("HOW TO READ THIS") + nrows)
    return page("Start here", body, 1500)


def main():
    os.makedirs(OUT, exist_ok=True)
    jobs = [("sheet-files", files_sheet(), 1500, 560), ("sheet-recon", recon_sheet(), 1500, 470),
            ("sheet-activity", log_sheet(), 1500, 520)]
    local = "Egyptian Pound"
    built = "Built by Avati ALM Desk 3.0 on 26 Sep 2026, 14:05   ·   data as of 30 Nov 2025"
    if SAMPLE and os.path.exists(SAMPLE):
        pv, local = pivot_sheet()
        jobs.append(("built-balance-sheet", pv, 1500, 1030))
        made = [("LCR Output", "Rule-level output"), ("Balance sheet", "Balance sheet"),
                ("Deposits from all instituti LCY", "Deposits from all institutions for operational purposes ... - LCY"),
                ("TOther cash outflows due wi LCY", "TOther cash outflows due within 30 days - LCY")]
        jobs.append(("built-start-here", guide_sheet(sample_facts(), made, "LCR", built), 1500, 760))
    for name, html_, w, h in jobs:
        path = os.path.join(OUT, name + ".png")
        assets.render(html_, path, w, h, scale=1.5, transparent=False)
        print(path)


if __name__ == "__main__" and not os.environ.get("CONFIG_ONLY"):
    main()


# ---------------------------------------------------------------------------
#  Pivot config, with the defaults read out of the VBA itself
# ---------------------------------------------------------------------------
def _vba(name):
    with open(os.path.join(ROOT, "src", "vba", name + ".bas"), encoding="cp1252") as f:
        return f.read()


def _consts():
    import re
    out = {}
    for m in re.finditer(r'Public Const (\w+) As String = "([^"]*)"', _vba("modPD_Const") + _vba("modPD_Config")):
        out[m.group(1)] = m.group(2)
    out["MAX_SHEETS_DEFAULT"] = "120"
    return out


def _args(text, consts):
    """Split a VBA argument list into Python strings, resolving constants."""
    import re
    vals, cur, in_str, i = [], "", False, 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if in_str and i + 1 < len(text) and text[i + 1] == '"':
                cur += '"'
                i += 2
                continue
            in_str = not in_str
            cur += ch
        elif ch == "," and not in_str:
            vals.append(cur.strip())
            cur = ""
        else:
            cur += ch
        i += 1
    vals.append(cur.strip())
    out = []
    for v in vals:
        v = re.sub(r"CStr\((\w+)\)", r"\1", v)
        parts = [p.strip() for p in re.split(r"&(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", v)]
        s = ""
        for p in parts:
            if p.startswith('"'):
                s += p[1:-1]
            elif p in consts:
                s += consts[p]
            elif p.startswith("Chr$(34)"):
                s += '"'
        out.append(s)
    return out


def default_recipes():
    import re
    src = _vba("modPD_Config")
    body = src[src.index("Private Sub WriteDefaultRecipes"):src.index("Private Sub Recipe(")]
    body = re.sub(r" _\n\s*", " ", body)
    consts = _consts()
    rows = []
    for m in re.finditer(r"^\s*Recipe ws, r(?: \+ \d)?, (.*)$", body, re.M):
        rows.append(_args(m.group(1), consts))
    return rows


def default_fields():
    import re
    src = _vba("modPD_Config")
    consts = _consts()
    consts["NUM_FMT"] = "#,##0;[Red](#,##0);-"
    out = []
    for m in re.finditer(r"^\s*c\.Add Array\((.*)\)\s*$", src, re.M):
        out.append(_args(m.group(1), consts))
    return out


def band_row(widths, groups):
    """Pivot config's row 6: which run of columns is for what."""
    cols = "".join('<col style="width:%dpx">' % colw(w) for w in widths)
    g = "<tr style='height:%.1fpx'>" % px(20)
    for c1, c2, label in groups:
        g += ("<td colspan=%d style='background:#%s;color:#%s;font-family:\"Selawik Semibold\";font-size:%.1fpx;"
              "letter-spacing:.8px;border-left:3px solid #%s'>%s</td>") % (c2 - c1 + 1, C["E950"], C["M300"], px(7.5),
                                                                          C["BRAND"], label)
    return g + "</tr>"


def config_sheet():
    heads = ["On", "Pivot", "Frameworks", "One sheet per", "Rows", "Columns", "Values", "Show only / hide",
             "Slicers", "Layout", "Subtotals", "Grand totals", "Repeat labels", "Sort", "Widths",
             "Number format", "Tab", "Max sheets", "Description", "Check", "What to fix"]
    widths = [7, 20, 16, 20, 36, 14, 38, 30, 26, 10, 11, 14, 9, 16, 24, 16, 9, 9, 40, 10, 60]
    groups = [(1, 4, "WHICH PIVOT"), (5, 9, "WHAT IT SHOWS"), (10, 16, "HOW IT LOOKS"), (17, 19, "THE SHEET"),
              (20, 21, "CHECK")]
    rows = []
    for r in default_recipes():
        on = r[0] == "Yes"
        rows.append(r + (["OK", ""] if on else ["Off", ""]))
    rows += [[""] * 21 for _ in range(4)]
    t = table(heads, widths, rows, verdict_col=19, muted_cols=(18, 20), semi_cols=(1,),
              pre_rows=band_row(widths, groups))
    body = (chrome("reports", "Pivot config", "What every framework workbook is built from. One row is one pivot, or "
                   "one sheet per value of a field. Edit a row or add one; Check says whether it will build. Select "
                   "any cell for how to fill it.", "4 pivot(s) switched on, every one ready to build.", "OK",
                   ["Add a pivot", "Check", "Restore defaults", "Use the 1.0 layout"], tab="Pivots",
                   overline="REPORTS &nbsp;&#183;&nbsp; PIVOTS", status_w=1400) + t)
    return page("Pivot config", body, 2700)


def fields_sheet():
    fh = ["Field", "Source column", "Kind", "Blank shows as", "Width", "Number format", "Note"]
    fw = [26, 42, 10, 20, 8, 22, 64]
    frows = [r[:7] for r in default_fields()]
    t = table(fh, fw, frows, mono_cols=(1,), muted_cols=(6,), semi_cols=(0,))
    body = (chrome("reports", "Pivot fields", "Every column a pivot or a chart may name. The first thirteen are "
                   "built in; add any column of an output under a name of your own, and use that name anywhere.",
                   "%d fields. Built-in fields cannot be renamed; any other row can be changed, and new ones added "
                   "at the bottom." % len(frows), "Idle", ["Check"], tab="Fields",
                   overline="REPORTS &nbsp;&#183;&nbsp; FIELDS", status_w=1400) + t)
    return page("Pivot fields", body, 1500)


if __name__ == "__main__" and os.environ.get("CONFIG_ONLY"):
    assets.render(config_sheet(), os.path.join(OUT, "sheet-pivot-config.png"), 2700, 520, scale=1, transparent=False)
    assets.render(fields_sheet(), os.path.join(OUT, "sheet-pivot-fields.png"), 1500, 1340, scale=1, transparent=False)
    print("config previews")
