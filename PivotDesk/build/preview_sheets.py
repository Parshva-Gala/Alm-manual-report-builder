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

C = dict(INK="000000", TX1="F2F7F4", TX2="A9BDB3", SOFT="C2EBDA", BRAND="009060", DEEP="006141",
         E900="003323", E950="001D14", PAPER="FFFFFF", MIST="F4F7F5", HAIR="E2E9E5", MUTED="5F7068",
         BODY="18241F", OK_TX="0B6B47", OK_BG="DDF5EA", WARN_TX="8A5A00", WARN_BG="FFF1CF",
         BAD_TX="B42318", BAD_BG="FDE5E2", IDLE_TX="5E6F67", IDLE_BG="EDF2EF", TINT="E6F6EF", M300="4FC79C")
LEVEL = {"OK": ("OK_TX", "OK_BG"), "LOADED": ("OK_TX", "OK_BG"), "BREAK": ("BAD_TX", "BAD_BG"),
         "MISSING": ("BAD_TX", "BAD_BG"), "CHECK": ("WARN_TX", "WARN_BG"), "EMPTY": ("IDLE_TX", "IDLE_BG"),
         "IDLE": ("IDLE_TX", "IDLE_BG")}

PX_PER_CHAR = 8          # Segoe UI 9.5 on a 7 px Calibri grid, near enough
PT = 4 / 3


def colw(chars):
    return round(chars * 7 + 5)


def lv(word):
    fg, bg = LEVEL.get(word.upper(), LEVEL["IDLE"])
    return C[fg], C[bg]


def page(title, body, width):
    return ('<!doctype html><html><head><meta charset="utf-8"><title>%s</title><style>'
            'html,body{margin:0;background:#fff;font-family:Selawik;color:#%s}'
            '.wrap{width:%dpx;overflow:hidden;background:#fff}'
            'table{border-collapse:collapse;table-layout:fixed}'
            'td{padding:0 0 0 9px;white-space:nowrap;overflow:hidden;text-overflow:clip;font-size:%.1fpx}'
            '.bar td{background:#000}'
            '.pill{display:inline-block;padding:0 12px;border-radius:12px;line-height:22px;font-family:"Selawik Semibold";font-size:11.3px}'
            '.nav{height:45px;background:#000;display:flex;align-items:center;gap:5px;padding-left:16px;position:relative}'
            '.nb{height:29px;border-radius:15px;display:flex;align-items:center;padding:0 14px;font-family:"Selawik Semibold";font-size:11.3px;color:#%s}'
            '.nb.cur{background:#%s;box-shadow:inset 0 0 0 1px #%s;color:#%s}'
            '.nb.act{background:#%s;box-shadow:inset 0 0 0 1px #%s;color:#%s}'
            '.div{width:1px;height:19px;background:#1A2922;margin:0 8px}'
            '.wm{font-family:"Selawik Semibold";font-size:14.7px;color:#%s;margin-right:6px}.wm b{font-weight:normal;color:#%s}'
            '</style></head><body><div class="wrap">%s</div></body></html>' % (
                title, C["BODY"], width, 9.5 * PT, C["TX2"], C["E900"], C["DEEP"], C["TX1"], C["E950"], C["DEEP"],
                C["SOFT"], C["TX1"], C["M300"], body))


def nav(current, actions):
    items = [("Desk", "desk"), ("Files", "files"), ("Reconciliation", "recon"), ("Activity", "log")]
    h = ['<div class="nav"><span class="wm">Pivot<b>Desk</b></span><span class="div"></span>']
    for label, key in items:
        h.append('<span class="nb%s">%s</span>' % (" cur" if key == current else "", label))
    h.append('<span class="div"></span>')
    for a in actions:
        h.append('<span class="nb act">%s</span>' % a)
    h.append("</div>")
    return "".join(h)


def masthead(title, about, status, level, ncols_px, statuswidth_px):
    fg, bg = lv(level)
    return (
        '<div style="background:#000;height:40px;border-left:4px solid #%s;display:flex;align-items:flex-end;'
        'padding-left:12px;font-family:\'Selawik Semibold\';font-size:%.1fpx;color:#%s">%s</div>'
        '<div style="background:#000;height:28px;border-left:4px solid #%s;border-bottom:3px solid #%s;padding-left:12px;'
        'font-size:%.1fpx;color:#%s">%s</div>'
        '<div style="height:13px"></div>'
        '<div style="width:%dpx;height:35px;background:#%s;border-left:4px solid #%s;display:flex;align-items:center;'
        'padding-left:12px;font-family:\'Selawik Semibold\';font-size:%.1fpx;color:#%s">&#9679;&nbsp;&nbsp;%s</div>'
        '<div style="height:13px"></div>' % (
            C["BRAND"], 17 * PT, C["TX1"], html.escape(title), C["BRAND"], C["BRAND"], 9 * PT, C["TX2"],
            html.escape(about), statuswidth_px, bg, fg, 9.5 * PT, fg, html.escape(status)))


def table(headers, widths, rows, verdict_col=None, mono_cols=(), muted_cols=(), right_cols=(), semi_cols=(),
          bars_col=None, band=True):
    cols = "".join('<col style="width:%dpx">' % colw(w) for w in widths)
    h = ['<table>%s<tr style="height:35px;background:#000">' % cols]
    for t in headers:
        h.append('<td style="font-family:\'Selawik Semibold\';font-size:%.1fpx;color:#%s;border-bottom:2px solid #%s">%s</td>'
                 % (8.5 * PT, C["SOFT"], C["BRAND"], html.escape(t)))
    h.append("</tr>")
    maxbar = max((abs(r[bars_col]) for r in rows if bars_col is not None and isinstance(r[bars_col], (int, float))),
                 default=1) or 1
    for i, r in enumerate(rows):
        bgc = C["MIST"] if (band and i % 2 == 1) else C["PAPER"]
        h.append('<tr style="height:28px;background:#%s">' % bgc)
        for j, v in enumerate(r):
            st = ["border-bottom:1px solid #%s" % C["HAIR"]]
            txt = v
            if isinstance(v, (int, float)):
                txt = "-" if v == 0 else ("(%s)" % format(abs(v), ",.0f") if v < 0 else format(v, ",.0f"))
                st.append("text-align:right;padding-right:9px")
                if v < 0:
                    st.append("color:#C00000")
            if j == verdict_col:
                fg, bg = lv(str(v))
                st = ["border-bottom:1px solid #%s" % C["HAIR"], "text-align:center", "background:#%s" % bg,
                      "color:#%s" % fg, "font-family:'Selawik Semibold'", "font-size:%.1fpx" % (8.5 * PT)]
                txt = "&#9679;&nbsp;&nbsp;" + html.escape(str(v))
            else:
                txt = html.escape(str(txt))
            if j in mono_cols:
                st.append("font-family:'DejaVu Sans Mono';font-size:11px;color:#%s" % C["MUTED"])
            if j in muted_cols:
                st.append("color:#%s" % C["MUTED"])
            if j in semi_cols:
                st.append("font-family:'Selawik Semibold'")
            if j == bars_col and isinstance(v, (int, float)) and v:
                w = abs(v) / maxbar * 100
                st.append("background:linear-gradient(90deg, rgba(244,162,154,.85) %.1f%%, transparent %.1f%%)" % (w, w))
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
    body = (nav("files", ["Scan a folder", "Pick files", "Use a file for this row", "Clear this row"]) +
            masthead("Files", "Every file is identified by the columns it carries, not by its name. Add them from "
                     "the Desk, or click a row here and press \"Use a file for this row\" to place one by hand.",
                     "4 of 5 file(s) loaded.  Ready to build pivots.", "OK", w, 1040) +
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
    body = (nav("recon", ["Reconcile now"]) +
            masthead("Reconciliation", "Each loaded output against control report 3 (the ledger, by COA) and control "
                     "report 6 (the reporting balance, by account).",
                     "1 of 2 comparison(s) broke, the largest involving 1,204,331,902 LCY.   Scope is on Activity for "
                     "each - read it before treating a difference as an error.   (12.4s)", "Break", 0, 1300) +
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
    body = (nav("log", ["Clear activity"]) +
            masthead("Activity", "What the desk did and why, newest first. A file that was not recognised says here "
                     "what was missing.", "The newest entry is at the top. The Desk shows the latest three.", "Idle",
                     0, 1300) +
            table(["When", "Level", "Stage", "What happened", "Which file"], widths, rows, verdict_col=1,
                  mono_cols=(0,), muted_cols=(4,), semi_cols=(2,), band=False))
    return page("Activity", body, 1500)


# ---------------------------------------------------------------------------
#  A built workbook, from the LCR sample
# ---------------------------------------------------------------------------
def balance_pivot():
    import openpyxl
    wb = openpyxl.load_workbook(SAMPLE, read_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    ix = {}
    for i, h in enumerate(hdr):
        ix.setdefault(h, i)
    counts = collections.Counter()
    for r in it:
        counts[r[ix["CURRENCY_NAME"]]] += 1
    local = counts.most_common(1)[0][0]
    wb = openpyxl.load_workbook(SAMPLE, read_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    next(it)
    agg = collections.defaultdict(lambda: [0.0, 0.0])
    for r in it:
        if not r[ix["BUCKET_DISPLAY_NAME"]]:
            continue                                    # the pivot's default: blanks out
        key = (r[ix["COA_BALANCESHEET_CATEGORY"]] or "(no type)",
               r[ix["BALANCESHEET_LINE_NAME"]] or "(no line)",
               r[ix["BALANCESHEET_SUB_LINE_NAME"]] or "(no subline)",
               r[ix["COA_NAME"]] or "(no COA)")
        side = 0 if r[ix["CURRENCY_NAME"]] == local else 1
        agg[key][side] += float(r[ix["CASHFLOW_AMOUNT_LCY_PRE_FACTOR"]] or 0)
    return local, sorted(agg.items(), key=lambda kv: -abs(kv[1][0] + kv[1][1]))


def pivot_sheet():
    local, items = balance_pivot()
    items = items[:16]
    widths = [14, 34, 34, 38, 14, 14, 14]
    rows_html = []
    prev = [None, None, None]
    for i, (key, (lcy, fcy)) in enumerate(items):
        cells = []
        for j in range(4):
            show = key[j] if (j == 3 or key[j] != prev[j] or any(key[k] != prev[k] for k in range(j))) else ""
            cells.append(show)
        prev = list(key[:3])
        vals = [lcy, fcy, lcy + fcy]
        bg = C["MIST"] if i % 2 else C["PAPER"]
        tds = "".join('<td style="border-bottom:1px solid #%s;%s">%s</td>' % (
            C["HAIR"], "font-family:'Selawik Semibold'" if (j == 0 and c) else "", html.escape(str(c)))
            for j, c in enumerate(cells))
        tds += "".join('<td style="border-bottom:1px solid #%s;text-align:right;padding-right:9px;%s">%s</td>' % (
            C["HAIR"], "color:#C00000" if v < 0 else "",
            "-" if round(v) == 0 else ("(%s)" % format(abs(v), ",.0f") if v < 0 else format(v, ",.0f")))
            for v in vals)
        rows_html.append('<tr style="height:24px;background:#%s">%s</tr>' % (bg, tds))
    tot = [sum(v[0] for _, v in items), sum(v[1] for _, v in items)]
    tot.append(tot[0] + tot[1])
    cols = "".join('<col style="width:%dpx">' % colw(w) for w in widths)
    head = ("<tr style='height:22px'><td colspan=4></td>"
            "<td colspan=3 style='background:#000;color:#%s;font-family:\"Selawik Semibold\"'>LCY / FCY</td></tr>"
            "<tr style='height:28px;background:#000'>%s</tr>") % (
        C["SOFT"], "".join("<td style='color:#%s;font-family:\"Selawik Semibold\";border-bottom:1px solid #000'>%s</td>" % (
            C["SOFT"], h) for h in ["Type", "Line", "Subline", "COA name", "LCY", "FCY", "Grand Total"]))
    total_row = ("<tr style='height:26px;background:#%s'><td colspan=4 style='font-family:\"Selawik Semibold\";"
                 "border-top:2px solid #%s'>Grand Total</td>%s</tr>") % (
        C["TINT"], C["BRAND"], "".join("<td style='text-align:right;padding-right:9px;font-family:\"Selawik Semibold\";"
                                        "border-top:2px solid #%s;%s'>%s</td>" % (
                                            C["BRAND"], "color:#C00000" if v < 0 else "",
                                            "(%s)" % format(abs(v), ",.0f") if v < 0 else format(v, ",.0f"))
                                        for v in tot))
    filt = ("<div style='padding:10px 0 12px 12px;font-size:12.6px'><span style='font-family:\"Selawik Semibold\";"
            "color:#%s'>Bucket</span>&nbsp;&nbsp;&nbsp;(Multiple Items)</div>") % C["MUTED"]
    bar = ("<div style='background:#000;height:45px;display:flex;align-items:center;padding-left:16px;"
           "font-family:\"Selawik Semibold\";font-size:11.3px;color:#%s'>&#8249;&nbsp;&nbsp;Start here</div>") % C["M300"]
    body = (bar + masthead("Balance sheet", "The same balances as the balance sheet reads them, down to the COA.",
                           "", "Idle", 0, 0).split('<div style="height:13px"></div>')[0] +
            "<div style='height:13px'></div>" + slicers() + filt +
            "<table>%s%s%s%s</table>" % (cols, head, "".join(rows_html), total_row))
    return page("Balance sheet", body, 1500), local


def slicers():
    def one(title, items, sel):
        chips = "".join(
            "<div style='height:19px;border-radius:3px;margin:3px 0;padding-left:8px;font-size:11px;line-height:19px;"
            "background:%s;color:%s'>%s</div>" % ("#009060" if s else "#1D2B25", "#fff" if s else "#A9BDB3",
                                                 html.escape(t)) for t, s in zip(items, sel))
        return ("<div style='width:204px;height:77px;background:#0C1512;border-radius:4px;padding:6px 8px;"
                "box-sizing:border-box;margin-right:8px;display:inline-block;vertical-align:top;overflow:hidden'>"
                "<div style='color:#F2F7F4;font-family:\"Selawik Semibold\";font-size:11.5px;margin-bottom:2px'>%s</div>"
                "<div style='column-count:2;column-gap:6px'>%s</div></div>") % (title, chips)
    return ("<div style='padding-left:11px;height:90px'>%s%s%s</div>" % (
        one("Type", ["Asset", "Liability", "Off Balance", "Equity"], [1, 1, 1, 1]),
        one("LCY / FCY", ["LCY", "FCY"], [1, 1]),
        one("Rule category", ["OUTFLOW", "(no category)"], [1, 1])))


def guide_sheet(local):
    made = [("Start here", "This index"), ("LCR Output", "Rule-level output"), ("Balance sheet", "Balance sheet"),
            ("Deposits from all institution LCY", "Deposits from all institutions for operational purposes ... - LCY"),
            ("Deposits from all institution FCY", "Deposits from all institutions for operational purposes ... - FCY"),
            ("TOther cash outflows due wi LCY", "TOther cash outflows due within 30 days - LCY"),
            ("TOther cash outflows due wi FCY", "TOther cash outflows due within 30 days - FCY")][1:]
    widths = [38, 110]
    rows = [[a, b] for a, b in made]
    bar = ("<div style='background:#000;height:45px;display:flex;align-items:center;padding-left:16px;"
           "font-family:\"Selawik Semibold\";font-size:10.7px;color:#%s;letter-spacing:.5px'>PIVOTDESK&nbsp;&nbsp;&#183;"
           "&nbsp;&nbsp;START HERE</div>") % C["M300"]
    tbl = table(["Sheet", "What is on it"], widths, rows, muted_cols=(1,))
    tbl = tbl.replace("<td style=\"border-bottom:1px solid #E2E9E5\">", "<td style=\"border-bottom:1px solid #E2E9E5\">")
    # links in deep emerald, semibold
    import re
    tbl = re.sub(r'(<tr style="height:28px;background:#[0-9A-F]+"><td style=")', r"\1color:#%s;font-family:'Selawik Semibold';" % C["DEEP"], tbl)
    notes = [("Amounts", "Pre-factor and post-factor from CASHFLOW_AMOUNT_LCY_PRE_FACTOR  /  CASHFLOW_AMOUNT_LCY_POST_FACTOR."),
             ("Local currency", "\"%s\" - the currency on the most rows. Nothing in the extract says which is local, "
              "so check this before relying on the LCY / FCY split." % local),
             ("Factor", "Post-factor divided by pre-factor, so it always agrees with the two figures beside it."),
             ("Blanks", "Rows with no bucket are excluded from the bucket filter by default. They are still in the "
              "data - clear the filter to see them.")]
    nh = ("<div style='height:24px'></div><div style='width:%dpx;border-bottom:1px solid #%s;padding:0 0 5px 9px;"
          "font-family:\"Selawik Semibold\";font-size:11.3px;color:#%s'>HOW TO READ THIS</div>") % (
        colw(38) + colw(110), C["DEEP"], C["DEEP"])
    nrows = "".join("<div style='display:flex;padding:7px 0'><div style='width:%dpx;padding-left:9px;"
                    "font-family:\"Selawik Semibold\"'>%s</div><div style='width:%dpx;padding-left:9px;color:#%s;"
                    "white-space:normal'>%s</div></div>" % (colw(38) - 9, a, colw(110) - 9, C["MUTED"], html.escape(b))
                    for a, b in notes)
    body = (bar + masthead("LCR  -  MIDBANK  Cairo", "Built by PivotDesk 2.0 on 26 Sep 2026 14:05   -   data as of "
                           "30 Nov 2025", "4,480 rows staged into one pivot cache.  Every sheet below is a live "
                           "PivotTable over it - drag a field, drop a slicer, drill a total.", "OK", 0, 1060) +
            tbl + nh + nrows)
    return page("Start here", body, 1500)


def main():
    os.makedirs(OUT, exist_ok=True)
    jobs = [("sheet-files", files_sheet(), 1500, 560), ("sheet-recon", recon_sheet(), 1500, 470),
            ("sheet-activity", log_sheet(), 1500, 520)]
    local = "Egyptian Pound"
    if SAMPLE and os.path.exists(SAMPLE):
        pv, local = pivot_sheet()
        jobs.append(("built-balance-sheet", pv, 1500, 880))
    jobs.append(("built-start-here", guide_sheet(local), 1500, 700))
    for name, html_, w, h in jobs:
        path = os.path.join(OUT, name + ".png")
        assets.render(html_, path, w, h, scale=1.5, transparent=False)
        print(path)


if __name__ == "__main__":
    main()
