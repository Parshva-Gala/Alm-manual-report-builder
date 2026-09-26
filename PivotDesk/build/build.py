"""
Build dist/Avati.xlsm from sources, without Excel.

    seed workbook  (src/seed/PivotDesk_v1.xlsm)   the package, sheets and styles
    VBA sources    (src/vba/*.bas, *.cls)          the whole program
    Desk design    (build/desk.py + design/assets) the Desk's drawing and artwork

What the build changes in the package, and why:

  * the VBA project is rewritten from src/vba (modules added, changed, removed)
  * the Desk sheet is replaced by the designed canvas and its drawing
  * the console's source sheet becomes the very hidden _Settings sheet - same
    sheet, same code name, so its document module stays attached
  * the 1.0 button rails on Files / Reconciliation / Activity are dropped; the
    code draws the 2.0 bar on first open
  * the workbook theme takes the brand colours, so every built-in style in
    the workbook is emerald rather than Office blue
  * calculation goes back to automatic - 1.0 was saved in manual, which set
    every workbook opened after it in the same Excel to manual as well

Then it checks what it built: the VBA lint, the shape-name contract between
the code and the drawing, every OnAction target, and the package itself.

    python3 build/build.py            build and check
"""

from __future__ import annotations

import datetime as _dt
import glob
import io
import os
import re
import sys
import zipfile
from xml.dom import minidom

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import contrast  # noqa: E402
import desk  # noqa: E402
import ribbon  # noqa: E402
import vbalint  # noqa: E402
import vbaproj  # noqa: E402
from design_tokens import EM, CANVAS, GRID_COL_PT, GRID_ROW_PT  # noqa: E402
from shapes import to_drawingml  # noqa: E402

SEED = os.path.join(ROOT, "src", "seed", "PivotDesk_v1.xlsm")
VBA_DIR = os.path.join(ROOT, "src", "vba")
DIST = os.path.join(ROOT, "dist", "Avati.xlsm")

# Desk columns are exactly 20 px wide at the workbook's 7 px maximum digit
# width (Calibri 11, the Normal style): 2.7890625 characters.
DESK_COL_WIDTH = "2.7890625"


# ============================================================================
#  VBA
# ============================================================================

def read_sources():
    mods = {}
    for path in sorted(glob.glob(os.path.join(VBA_DIR, "*.bas")) + glob.glob(os.path.join(VBA_DIR, "*.cls"))):
        name = os.path.splitext(os.path.basename(path))[0]
        with open(path, encoding="cp1252") as f:
            mods[name] = f.read()
    return mods


def build_vba(seed_bin: bytes, sources: dict) -> bytes:
    proj = vbaproj.VBAProject(vbaproj.read_cfb(seed_bin))
    for m in list(proj.modules):
        if m.name not in sources and m.kind != "doc":
            proj.remove(m.name)
    for name, code in sources.items():
        if proj.has(name):
            proj.set_code(name, code)
        else:
            proj.add_std(name, code)
    return proj.to_bin()


# ============================================================================
#  The Desk sheet
# ============================================================================

def desk_sheet_xml(canvas_xf: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetPr><tabColor rgb="FF%s"/><pageSetUpPr fitToPage="1"/></sheetPr>'
        '<dimension ref="A1"/>'
        '<sheetViews><sheetView showGridLines="0" showRowColHeaders="0" tabSelected="1" zoomScale="100" '
        'zoomScaleNormal="100" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
        '<sheetFormatPr baseColWidth="8" defaultColWidth="%s" defaultRowHeight="15" customHeight="1"/>'
        '<cols><col min="1" max="16384" width="%s" style="%d" customWidth="1"/></cols>'
        '<sheetData/>'
        '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
        '<pageMargins left="0.3" right="0.3" top="0.3" bottom="0.3" header="0.2" footer="0.2"/>'
        '<pageSetup orientation="landscape" fitToHeight="1" fitToWidth="1"/>'
        '<drawing r:id="rId1"/>'
        '</worksheet>' % (EM[500], DESK_COL_WIDTH, DESK_COL_WIDTH, canvas_xf))


def shipped_state():
    st = desk.empty_state(today="PIVOTDESK  ·  ALM DESK", greeting="Welcome.")
    st["macros_banner"] = True
    return st


def desk_drawing(state):
    shapes = desk.desk(state)
    desk.render_icons()
    images = []
    for s in shapes:
        if s.image and s.image not in images:
            images.append(s.image)
    rels, rids, media = [], {}, []
    n = 1
    for base in images:
        png_name = "pdx%d.png" % (len(media) + 1)
        media.append((png_name, base + ".png"))
        png_rid = "rId%d" % n
        rels.append((png_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                     "../media/" + png_name))
        n += 1
        svg_rid = None
        if os.path.exists(base + ".svg"):
            svg_name = "pdx%d.svg" % (len(media) + 1)
            media.append((svg_name, base + ".svg"))
            svg_rid = "rId%d" % n
            rels.append((svg_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                         "../media/" + svg_name))
            n += 1
        rids[base] = (png_rid, svg_rid)
    # the hero is PNG only: it uses blur and masks Office's SVG renderer lacks
    for s in shapes:
        if s.image and not s.image_svg:
            rids[s.image] = (rids[s.image][0], None)
    xml = to_drawingml(shapes, GRID_COL_PT, GRID_ROW_PT, rids)
    rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + "".join('<Relationship Id="%s" Type="%s" Target="%s"/>' % r for r in rels)
                + "</Relationships>")
    used = {name for name, _ in media}
    return xml, rels_xml, media, shapes, used


# ============================================================================
#  Package edits
# ============================================================================

def add_canvas_style(styles: str):
    m = re.search(r'<fills count="(\d+)">(.*?)</fills>', styles, re.S)
    n_fills = int(m.group(1))
    fill = ('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/><bgColor indexed="64"/>'
            '</patternFill></fill>' % CANVAS)
    styles = styles.replace(m.group(0), '<fills count="%d">%s%s</fills>' % (n_fills + 1, m.group(2), fill))
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    n_xf = int(m.group(1))
    xf = '<xf numFmtId="0" fontId="0" fillId="%d" borderId="0" xfId="0" applyFill="1"/>' % n_fills
    styles = styles.replace(m.group(0), '<cellXfs count="%d">%s%s</cellXfs>' % (n_xf + 1, m.group(2), xf))
    return styles, n_xf


def brand_theme(theme: str) -> str:
    colours = {
        "dk2": "0C1512", "lt2": "F4F7F5",
        "accent1": EM[500], "accent2": EM[400], "accent3": EM[300], "accent4": EM[700],
        "accent5": EM[200], "accent6": "5F7068", "hlink": EM[600], "folHlink": EM[800],
    }
    for tag, val in colours.items():
        theme, k = re.subn(r"<a:%s>.*?</a:%s>" % (tag, tag),
                           '<a:%s><a:srgbClr val="%s"/></a:%s>' % (tag, val, tag), theme, count=1, flags=re.S)
        assert k == 1, tag
    theme = theme.replace('<a:clrScheme name="Office">', '<a:clrScheme name="Avati">', 1)
    return theme


def settings_sheet(old: str) -> str:
    head = re.search(r"^(.*?<sheetData>)", old, re.S).group(1)
    tail = re.search(r"(</sheetData>.*)$", old, re.S).group(1)
    head = re.sub(r"<dimension ref=\"[^\"]*\"/>", '<dimension ref="A1:B1"/>', head)
    rows = ('<row r="1"><c r="A1" t="inlineStr"><is><t>Key</t></is></c>'
            '<c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>')
    return head + rows + tail


def drop_drawing(sheet: str) -> str:
    return re.sub(r"<drawing r:id=\"[^\"]*\"/>", "", sheet)


def core_props(old: str, version: str) -> str:
    now = _dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    extra = ("<dc:title>Avati ALM Desk " + version + " - MIDBANK Cairo</dc:title>"
             "<dc:subject>ALM outputs staged into live PivotTables, and reconciled against control reports 3 and 6</dc:subject>"
             "<cp:keywords>ALM; LCR; NSFR; maturity ladder; reconciliation; Avati</cp:keywords>"
             "<dc:description>Open, enable content, and work from the Desk.</dc:description>"
             "<cp:category>ALM desk</cp:category>")
    old = re.sub(r"<dc:title>.*?</dc:title>|<dc:subject>.*?</dc:subject>|<cp:keywords>.*?</cp:keywords>"
                 r"|<dc:description>.*?</dc:description>|<cp:category>.*?</cp:category>", "", old)
    old = re.sub(r"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)", r"\g<1>%s\2" % now, old)
    return old.replace("</cp:coreProperties>", extra + "</cp:coreProperties>")


def build():
    import brand
    if not os.path.exists(os.path.join(ROOT, "design", "brand", "avati-logo.png")):
        brand.build()
    zin = zipfile.ZipFile(SEED)
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]

    sources = read_sources()
    parts["xl/vbaProject.bin"] = build_vba(parts["xl/vbaProject.bin"], sources)

    styles, canvas_xf = add_canvas_style(parts["xl/styles.xml"].decode("utf-8"))
    parts["xl/styles.xml"] = styles.encode("utf-8")
    parts["xl/theme/theme1.xml"] = brand_theme(parts["xl/theme/theme1.xml"].decode("utf-8")).encode("utf-8")

    parts["xl/worksheets/sheet1.xml"] = desk_sheet_xml(canvas_xf).encode("utf-8")
    drawing, drels, media, shapes, _ = desk_drawing(shipped_state())
    parts["xl/drawings/drawing1.xml"] = drawing.encode("utf-8")
    parts["xl/drawings/_rels/drawing1.xml.rels"] = drels.encode("utf-8")
    if "xl/drawings/_rels/drawing1.xml.rels" not in order:
        order.append("xl/drawings/_rels/drawing1.xml.rels")
    for name, src in media:
        with open(src, "rb") as f:
            parts["xl/media/" + name] = f.read()
        order.append("xl/media/" + name)

    for i in (2, 3, 4):
        sp = "xl/worksheets/sheet%d.xml" % i
        parts[sp] = drop_drawing(parts[sp].decode("utf-8")).encode("utf-8")
        for p in ("xl/worksheets/_rels/sheet%d.xml.rels" % i, "xl/drawings/drawing%d.xml" % i):
            parts.pop(p, None)
            if p in order:
                order.remove(p)

    # the PivotDesk tab on the ribbon, for Excel view
    ui, ui_rels, ui_images = ribbon.custom_ui()
    parts[ribbon.PART] = ui.encode("utf-8")
    parts["customUI/_rels/customUI14.xml.rels"] = ui_rels.encode("utf-8")
    order += [ribbon.PART, "customUI/_rels/customUI14.xml.rels"]
    for name, data in ui_images.items():
        parts[name] = data
        order.append(name)
    root = parts["_rels/.rels"].decode("utf-8")
    if ribbon.REL_TYPE not in root:
        root = root.replace("</Relationships>", '<Relationship Id="rIdAvatiUI" Type="%s" Target="%s"/>'
                            "</Relationships>" % (ribbon.REL_TYPE, ribbon.PART))
    parts["_rels/.rels"] = root.encode("utf-8")

    parts["xl/worksheets/sheet5.xml"] = settings_sheet(parts["xl/worksheets/sheet5.xml"].decode("utf-8")).encode("utf-8")

    wb = parts["xl/workbook.xml"].decode("utf-8")
    wb = wb.replace('name="_Console_Src"', 'name="_Settings"')
    wb = wb.replace(' calcMode="manual"', "")
    wb = re.sub(r'<x15ac:absPath[^>]*/>', "", wb)
    # The Desk prints as a one-page status snapshot.
    if "_xlnm.Print_Area" not in wb:
        wb = wb.replace("<definedNames>", '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
                        "Desk!$A$1:$BW$44</definedName>", 1)
    parts["xl/workbook.xml"] = wb.encode("utf-8")

    ct = parts["[Content_Types].xml"].decode("utf-8")
    for i in (2, 3, 4):
        ct = ct.replace('<Override PartName="/xl/drawings/drawing%d.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' % i, "")
    for ext, typ in (("png", "image/png"), ("svg", "image/svg+xml")):
        if 'Extension="%s"' % ext not in ct:
            ct = ct.replace("<Default Extension=\"xml\"", '<Default Extension="%s" ContentType="%s"/><Default Extension="xml"' % (ext, typ), 1)
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    version = re.search(r'TOOL_VERSION As String = "([^"]+)"', sources["modPD_Const"]).group(1)
    parts["docProps/core.xml"] = core_props(parts["docProps/core.xml"].decode("utf-8"), version).encode("utf-8")

    os.makedirs(os.path.dirname(DIST), exist_ok=True)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        first = ["[Content_Types].xml"]
        for name in first + [n for n in order if n not in first and n in parts]:
            z.writestr(name, parts[name])
    with open(DIST, "wb") as f:
        f.write(buf.getvalue())
    return parts, shapes, sources


# ============================================================================
#  Checks
# ============================================================================

def check_package(parts):
    problems = []
    for name, data in parts.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                minidom.parseString(data)
            except Exception as e:  # noqa: BLE001
                problems.append("%s: not well-formed (%s)" % (name, e))
    ct = parts["[Content_Types].xml"].decode("utf-8")
    for name in parts:
        if name == "[Content_Types].xml":
            continue
        ext = name.rsplit(".", 1)[-1]
        if ('PartName="/%s"' % name) not in ct and ('Extension="%s"' % ext) not in ct:
            problems.append("%s: no content type" % name)
    for name, data in parts.items():
        if not name.endswith(".rels"):
            continue
        base = os.path.dirname(os.path.dirname(name))
        for tgt, mode in re.findall(r'Target="([^"]+)"(?: TargetMode="([^"]+)")?', data.decode("utf-8")):
            if mode == "External":
                continue
            path = os.path.normpath(os.path.join(base, tgt)).replace("\\", "/").lstrip("/")
            if tgt.startswith("/"):
                path = tgt.lstrip("/")
            if path not in parts:
                problems.append("%s -> %s: target missing" % (name, tgt))
    return problems


def _code_only(src: str) -> str:
    """The source with comments removed and string literals kept."""
    out = []
    for line in src.split("\n"):
        in_str, cut = False, len(line)
        for i, ch in enumerate(line):
            if ch == '"':
                in_str = not in_str
            elif ch == "'" and not in_str:
                cut = i
                break
        out.append(line[:cut])
    return "\n".join(out)


def check_contract(shapes, sources):
    """Every shape name the code paints, and every macro anything calls."""
    problems = []
    names = {s.name for s in shapes}
    code = "\n".join(_code_only(src) for src in sources.values())

    for lit in sorted(set(re.findall(r'"(pdx_[A-Za-z0-9_]*)"', code))):
        if lit == "pdx_":
            continue                    # the family prefix itself, in a comparison
        # followed by & -> a prefix the code completes at run time
        exact = lit in names
        prefix = any(n.startswith(lit) for n in names)
        pattern = re.search(r'"%s"\s*&' % re.escape(lit), code)
        if pattern:
            if not prefix:
                problems.append("code builds names from %r but no shape starts with it" % lit)
        elif not exact:
            problems.append("code paints %r but the design has no such shape" % lit)
    # the families the code composes, spelled out
    expected = []
    for i in range(1, 6):
        expected += ["pdx_slot%d_dot" % i, "pdx_slot%d_meta" % i, "pdx_slot%d_hit" % i]
    for j in range(1, 4):
        expected += ["pdx_fw%d_%s" % (j, k) for k in ("row", "label", "meta", "track", "knob", "hit")]
    for key, n in (("files", 5), ("fw", 3), ("ctl", 2), ("recon", 1)):
        expected += ["pdx_kpi_%s_value" % key, "pdx_kpi_%s_sub" % key]
        expected += ["pdx_kpi_%s_seg%d" % (key, k) for k in range(1, n + 1)]
    for r in (1, 2):
        expected += ["pdx_m_%d%d" % (r, c) for c in (1, 2, 3)]
        expected += ["pdx_ctl%d_dot" % r, "pdx_ctl%d_meta" % r, "pdx_ctl%d_hit" % r]
    for a in range(1, desk.ACT_ROWS + 1):
        expected += ["pdx_act%d_%s" % (a, k) for k in ("when", "lvl", "stage", "msg")]
    for b in range(1, desk.RB_ROWS + 1):
        expected += ["pdx_rb%d_%s" % (b, k) for k in ("fw", "name", "meta", "open", "hit")]
    expected += ["pdx_tour_pip%d" % k for k in range(1, len(desk.TOUR) + 1)]
    for c in (1, 2, 3):
        expected += ["pdx_card%d_count" % c]
    for n in expected:
        if n not in names:
            problems.append("design is missing %s" % n)

    publics = set()
    for name, src in sources.items():
        for m in re.finditer(r"^Public Sub (\w+)\s*\(\s*\)", src, re.M):
            publics.add(m.group(1).lower())
    called = set(s.macro for s in shapes if s.macro)
    called |= set(re.findall(r'OnAction\s*=\s*"(\w+)"', code))
    called |= set(re.findall(r'Button\(ws, "[^"]*", "(\w+)"', code))
    called |= set(re.findall(r'OnKey "[^"]+", "(\w+)"', code))
    called |= set(re.findall(r'OnTime [^,]+, "(\w+)"', code))
    for mac in sorted(called):
        if mac.lower() not in publics:
            problems.append("nothing answers %s (a button, key or timer calls it)" % mac)
    return problems


def check_palette(sources):
    """The VBA palette and the design tokens are the same colours."""
    import design_tokens as T
    expect = {
        "C_INK": T.INK, "C_CANVAS": T.CANVAS, "C_SURFACE": T.SURFACE, "C_ELEV": T.SURFACE_2,
        "C_HAIR_DARK": T.HAIR, "C_BRAND": T.EM[500], "C_BRAND_BRIGHT": T.EM[400], "C_BRAND_DEEP": T.EM[700],
        "C_BRAND_900": T.EM[900], "C_BRAND_950": T.EM[950], "C_BRAND_SOFT": T.EM[100], "C_BRAND_TINT": T.EM[50],
        "C_TX1": T.TX_1, "C_TX2": T.TX_2, "C_TX3": T.TX_3, "C_TX4": T.TX_4,
        "C_PAPER": T.PAPER, "C_MIST": T.MIST, "C_HAIR": T.LINE, "C_MUTED": T.MUTED, "C_BODY": T.BODY,
        "C_OK_TX": T.OK_TX, "C_OK_BG": T.OK_LT, "C_WARN_TX": T.WARN_TX, "C_WARN_BG": T.WARN_LT,
        "C_BAD_TX": T.BAD_TX, "C_BAD_BG": T.BAD_LT, "C_IDLE_TX": T.IDLE_TX, "C_IDLE_BG": T.IDLE_LT,
        "C_OK_DK": T.OK, "C_WARN_DK": T.WARN, "C_BAD_DK": T.BAD, "C_IDLE_DK": T.IDLE,
    }
    problems = []
    theme = sources.get("modPD_Theme", "")
    got = dict(re.findall(r'\n\s*(C_\w+) = HX\("([0-9A-Fa-f]{6})"\)', theme))
    for name, hexv in expect.items():
        if name not in got:
            problems.append("palette: modPD_Theme has no %s" % name)
        elif got[name].upper() != hexv.upper():
            problems.append("palette: %s is %s in VBA but %s in the tokens" % (name, got[name], hexv))
    # every literal colour anywhere in the code is a token colour
    allowed = {v.upper() for v in expect.values()} | {v.upper() for v in T.EM.values()}
    allowed |= {x.upper() for x in (T.SURFACE_HI, T.SURFACE_3, T.HAIR_2, T.OK_BG, T.WARN_BG, T.BAD_BG, T.IDLE_BG,
                                    "FFFFFF", "000000", "3A4B43", "2A3B33", "17221D", "34443D", "3E5A50",
                                    "F4A29A")}
    for mod, src in sources.items():
        for hexv in re.findall(r'HX\("([0-9A-Fa-f]{6})"\)', _code_only(src)):
            if hexv.upper() not in allowed:
                problems.append("palette: %s uses #%s, which is not a design colour" % (mod, hexv))
    return problems


def check_ribbon(parts, sources):
    """Every ribbon button reaches code that handles it."""
    problems = []
    code = "\n".join(_code_only(src) for src in sources.values())
    if not re.search(r"^Public Sub %s\(control As Object\)" % ribbon.CALLBACK, code, re.M):
        problems.append("ribbon: no Public Sub %s(control As Object) for the buttons to call" % ribbon.CALLBACK)
    for bid in ribbon.ids():
        if not re.search(r'Case "%s":' % bid, code):
            problems.append("ribbon: button %s does nothing - %s has no Case for it" % (bid, ribbon.CALLBACK))
    ui = parts[ribbon.PART].decode("utf-8")
    rels = parts["customUI/_rels/customUI14.xml.rels"].decode("utf-8")
    for rid in set(re.findall(r'image="([^"]+)"', ui)):
        if 'Id="%s"' % rid not in rels:
            problems.append("ribbon: image %s has no relationship" % rid)
    return problems


def check_geometry(sources):
    """The numbers modPD_Desk lays the gap chart and the tour out with are the
    design's, and the tour says in Excel what it says in the preview."""
    problems = []
    src = sources.get("modPD_Desk", "")
    consts = {k: float(v) for k, v in re.findall(r"Private Const (\w+) As (?:Double|Long) = ([0-9.]+)", src)}
    want = {"TOUR_STEPS": len(desk.TOUR), "TOUR_PAD": desk.TOUR_PAD, "TOUR_GAP": desk.TOUR_GAP,
            "DESK_H": desk.H, "RB_ROWS": desk.RB_ROWS}
    for k, v in want.items():
        if k not in consts:
            problems.append("geometry: modPD_Desk has no constant %s" % k)
        elif abs(consts[k] - v) > 1e-9:
            problems.append("geometry: %s is %g in VBA but %g in the design" % (k, consts[k], v))
    fit = re.search(r'FIT_RANGE As String = "A1:([A-Z]+)(\d+)"', src)
    if not fit or int(fit.group(2)) * desk.GRID_ROW_PT != desk.H:
        problems.append("geometry: FIT_RANGE does not cover the %g pt canvas" % desk.H)
    for n, (targets, title, _) in enumerate(desk.TOUR, start=1):
        if ('title = "%s"' % title) not in src:
            problems.append("tour: step %d title %r is not the one modPD_Desk shows" % (n, title))
        if ('targets = "%s"' % ",".join(targets)) not in src:
            problems.append("tour: step %d lights %s in the design but not in modPD_Desk" % (n, targets))
    return problems


def main():
    parts, shapes, sources = build()
    problems = []
    lint = vbalint.analyse(dict(sources)) + vbalint.arity_problems(dict(sources))
    problems += ["vba %s:%d %s" % i for i in lint]
    problems += check_contract(shapes, sources)
    problems += check_package(parts)
    problems += check_palette(sources)
    problems += check_geometry(sources)
    problems += check_ribbon(parts, sources)
    tour = desk.showcase_state()
    tour.update({"toast": None, "tour": 5})
    for label, st in (("shipped", shipped_state()), ("showcase", desk.showcase_state()), ("tour", tour)):
        fails, _ = contrast.check(st, label)
        problems += ["contrast %.2f < %.1f on %s (%r)" % (f[0], f[1], f[3], f[4]) for f in fails]
    fails, _ = contrast.check_sheets()
    problems += ["contrast %.2f < %.1f: %s" % (f[0], f[1], f[2]) for f in fails]
    print("built %s  (%d KB, %d shapes on the Desk, %d modules)" % (
        os.path.relpath(DIST, ROOT), os.path.getsize(DIST) // 1024, len(shapes), len(sources)))
    for p in problems:
        print("  PROBLEM  " + p)
    print("%d problem(s)" % len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
