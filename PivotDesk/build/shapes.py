"""
A small shape model with two outputs from one description:

  to_drawingml()  the Desk's drawing part, as Excel reads it
  to_html()       a page Chromium renders, so the design can be looked at
                  and iterated off Windows

Units are points throughout. The Desk sits on a uniform grid of 15 pt
columns and rows, and every shape is anchored to that grid (twoCellAnchor,
editAs=absolute), so the layout scales with the grid whatever DPI Excel runs
at.
"""

from __future__ import annotations

import html
import math
from dataclasses import dataclass, field
from typing import Optional

EMU = 12700

FONT_PREVIEW = {
    "Segoe UI": "'Selawik'",
    "Segoe UI Semibold": "'Selawik Semibold'",
    "Segoe UI Light": "'Selawik Light'",
    "Segoe UI Semilight": "'Selawik Semilight'",
    "Consolas": "'DejaVu Sans Mono'",
}


# --- paint ------------------------------------------------------------------
@dataclass
class Solid:
    color: str
    alpha: float = 1.0


@dataclass
class Linear:
    angle: float                  # DrawingML sense: 0 = left->right, 90 = top->bottom
    stops: list                   # [(pos 0..1, hex, alpha)]


@dataclass
class Line:
    width: float
    color: str
    alpha: float = 1.0


@dataclass
class Shadow:
    blur: float
    dist: float
    direction: float = 90.0
    color: str = "000000"
    alpha: float = 0.4


@dataclass
class Glow:
    radius: float
    color: str
    alpha: float = 0.4


# --- text -------------------------------------------------------------------
@dataclass
class Run:
    text: str
    size: float = 9.0
    color: str = "FFFFFF"
    font: str = "Segoe UI"
    spc: float = 0.0              # letter spacing in points
    alpha: float = 1.0


@dataclass
class Para:
    runs: list
    align: str = "l"              # l | ctr | r


@dataclass
class Body:
    paras: list
    anchor: str = "t"             # t | ctr | b
    wrap: bool = True
    inset: tuple = (0, 0, 0, 0)   # l t r b
    clip: bool = False


# --- shapes -----------------------------------------------------------------
@dataclass
class Shape:
    name: str
    x: float
    y: float
    w: float
    h: float
    geom: str = "rect"            # rect | roundRect | ellipse
    radius: float = 0.0
    fill: object = None
    line: Optional[Line] = None
    shadow: Optional[Shadow] = None
    glow: Optional[Glow] = None
    body: Optional[Body] = None
    macro: str = ""
    descr: str = ""
    hidden: bool = False
    image: Optional[str] = None   # asset base path (without extension) for pictures
    image_svg: bool = False
    image_alpha: float = 1.0


def T(text, size=9.0, color="FFFFFF", font="Segoe UI", spc=0.0, align="l", alpha=1.0):
    return Para([Run(text, size, color, font, spc, alpha)], align)


# ============================================================================
#  DrawingML
# ============================================================================

def _clr(hexv, alpha=1.0):
    if alpha >= 0.999:
        return '<a:srgbClr val="%s"/>' % hexv
    return '<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr>' % (hexv, round(alpha * 100000))


def _fill(f):
    if f is None:
        return "<a:noFill/>"
    if isinstance(f, Solid):
        return "<a:solidFill>%s</a:solidFill>" % _clr(f.color, f.alpha)
    if isinstance(f, Linear):
        gs = "".join('<a:gs pos="%d">%s</a:gs>' % (round(p * 100000), _clr(c, a)) for p, c, a in f.stops)
        return ('<a:gradFill flip="none" rotWithShape="1"><a:gsLst>%s</a:gsLst>'
                '<a:lin ang="%d" scaled="0"/></a:gradFill>') % (gs, round(f.angle * 60000) % 21600000)
    raise TypeError(f)


def _geom(s: Shape):
    if s.geom == "roundRect":
        adj = min(50000, round(s.radius / max(1e-6, min(s.w, s.h)) * 100000))
        return '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst></a:prstGeom>' % adj
    return '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>' % s.geom


def _effects(s: Shape):
    parts = []
    if s.glow:
        parts.append('<a:glow rad="%d">%s</a:glow>' % (s.glow.radius * EMU, _clr(s.glow.color, s.glow.alpha)))
    if s.shadow:
        sh = s.shadow
        parts.append('<a:outerShdw blurRad="%d" dist="%d" dir="%d" algn="t" rotWithShape="0">%s</a:outerShdw>'
                     % (sh.blur * EMU, sh.dist * EMU, round(sh.direction * 60000), _clr(sh.color, sh.alpha)))
    return "<a:effectLst>%s</a:effectLst>" % "".join(parts) if parts else "<a:effectLst/>"


def _txbody(b: Optional[Body]):
    if b is None:
        return ""
    l, t, r, bb = b.inset
    body = ('<a:bodyPr vertOverflow="%s" horzOverflow="%s" wrap="%s" lIns="%d" tIns="%d" rIns="%d" bIns="%d" '
            'rtlCol="0" anchor="%s"><a:noAutofit/></a:bodyPr><a:lstStyle/>') % (
        "clip" if b.clip else "overflow", "clip" if b.clip else "overflow",
        "square" if b.wrap else "none", l * EMU, t * EMU, r * EMU, bb * EMU, b.anchor)
    ps = []
    for p in b.paras:
        runs = []
        last = None
        for r_ in p.runs:
            last = r_
            runs.append(
                '<a:r><a:rPr lang="en-GB" sz="%d" b="0" i="0" spc="%d" baseline="0">'
                '<a:solidFill>%s</a:solidFill><a:effectLst/><a:latin typeface="%s"/><a:cs typeface="%s"/>'
                '</a:rPr><a:t>%s</a:t></a:r>' % (
                    round(r_.size * 100), round(r_.spc * 100), _clr(r_.color, r_.alpha),
                    r_.font, r_.font, html.escape(r_.text, quote=False)))
        end = ('<a:endParaRPr lang="en-GB" sz="%d"><a:latin typeface="%s"/></a:endParaRPr>'
               % (round(last.size * 100), last.font)) if last else ""
        ps.append('<a:p><a:pPr algn="%s"/>%s%s</a:p>' % (p.align, "".join(runs), end))
    return "<xdr:txBody>%s%s</xdr:txBody>" % (body, "".join(ps))


def _anchor(s: Shape, col_pt, row_pt):
    def cell(v, size):
        i = int(math.floor(v / size + 1e-9))
        return i, round((v - i * size) * EMU)
    c1, o1 = cell(s.x, col_pt)
    r1, p1 = cell(s.y, row_pt)
    c2, o2 = cell(s.x + s.w, col_pt)
    r2, p2 = cell(s.y + s.h, row_pt)
    return ('<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:to>'
            % (c1, o1, r1, p1, c2, o2, r2, p2))


def to_drawingml(shapes, col_pt, row_pt, image_rids):
    """image_rids: asset base -> (png rId, svg rId or None)."""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    for i, s in enumerate(shapes, start=2):
        xfrm = '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (
            round(s.x * EMU), round(s.y * EMU), round(s.w * EMU), round(s.h * EMU))
        nv = '<xdr:cNvPr id="%d" name="%s" descr="%s"%s/>' % (
            i, html.escape(s.name), html.escape(s.descr or s.name), ' hidden="1"' if s.hidden else "")
        if s.image:
            png, svg = image_rids[s.image]
            ext = ""
            if svg:
                ext = ('<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
                       '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="%s"/>'
                       '</a:ext></a:extLst>') % svg
            alpha = "" if s.image_alpha >= 0.999 else '<a:alphaModFix amt="%d"/>' % round(s.image_alpha * 100000)
            body = ('<xdr:pic macro="%s"><xdr:nvPicPr>%s<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                    '<xdr:blipFill><a:blip r:embed="%s">%s%s</a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                    '<xdr:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:pic>'
                    % (html.escape(s.macro), nv, png, alpha, ext, xfrm))
        else:
            ln = ('<a:ln w="%d">%s</a:ln>' % (round(s.line.width * EMU), "<a:solidFill>%s</a:solidFill>" % _clr(s.line.color, s.line.alpha))
                  if s.line else "<a:ln><a:noFill/></a:ln>")
            body = ('<xdr:sp macro="%s" textlink=""><xdr:nvSpPr>%s<xdr:cNvSpPr/></xdr:nvSpPr>'
                    '<xdr:spPr>%s%s%s%s%s</xdr:spPr>%s</xdr:sp>'
                    % (html.escape(s.macro), nv, xfrm, _geom(s), _fill(s.fill), ln, _effects(s),
                       _txbody(s.body)))
        out.append('<xdr:twoCellAnchor editAs="absolute">%s%s<xdr:clientData/></xdr:twoCellAnchor>'
                   % (_anchor(s, col_pt, row_pt), body))
    out.append("</xdr:wsDr>")
    return "".join(out)


# ============================================================================
#  HTML preview
# ============================================================================

def _css_color(hexv, alpha=1.0):
    r, g, b = int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16)
    return "rgba(%d,%d,%d,%.3f)" % (r, g, b, alpha)


def _css_fill(f):
    if f is None:
        return "transparent"
    if isinstance(f, Solid):
        return _css_color(f.color, f.alpha)
    if isinstance(f, Linear):
        stops = ", ".join("%s %.1f%%" % (_css_color(c, a), p * 100) for p, c, a in f.stops)
        return "linear-gradient(%.1fdeg, %s)" % (f.angle + 90, stops)
    raise TypeError(f)


def to_html(shapes, width, height, canvas, assets_dir, scale=4 / 3, title="Desk"):
    px = lambda v: "%.2fpx" % (v * scale)
    parts = ['<!doctype html><html><head><meta charset="utf-8"><title>%s</title><style>'
             'html,body{margin:0;background:#%s;}'
             '.c{position:relative;width:%s;height:%s;overflow:hidden;background:#%s;}'
             '.s{position:absolute;box-sizing:border-box;}'
             '.t{position:absolute;inset:0;display:flex;flex-direction:column;box-sizing:border-box;}'
             '.p{white-space:pre-wrap;line-height:normal;}'
             '.nw .p{white-space:pre;}'
             '</style></head><body><div class="c">' % (title, canvas, px(width), px(height), canvas)]
    for s in shapes:
        if s.hidden:
            continue
        st = ["left:%s" % px(s.x), "top:%s" % px(s.y), "width:%s" % px(s.w), "height:%s" % px(s.h)]
        if s.image:
            src = "%s.%s" % (s.image, "png")
            st.append("opacity:%.3f" % s.image_alpha)
            parts.append('<img class="s" src="%s" style="%s">' % (src, ";".join(st)))
            continue
        st.append("background:%s" % _css_fill(s.fill))
        if s.geom == "roundRect":
            st.append("border-radius:%s" % px(min(s.radius, min(s.w, s.h) / 2)))
        elif s.geom == "ellipse":
            st.append("border-radius:50%")
        if s.line:
            st.append("box-shadow:inset 0 0 0 %s %s" % (px(s.line.width), _css_color(s.line.color, s.line.alpha)))
        shadows = []
        if s.line:
            shadows.append("inset 0 0 0 %s %s" % (px(s.line.width), _css_color(s.line.color, s.line.alpha)))
        if s.shadow:
            dx = s.shadow.dist * math.cos(math.radians(s.shadow.direction))
            dy = s.shadow.dist * math.sin(math.radians(s.shadow.direction))
            shadows.append("%s %s %s %s" % (px(dx), px(dy), px(s.shadow.blur), _css_color(s.shadow.color, s.shadow.alpha)))
        if s.glow:
            shadows.append("0 0 %s %s" % (px(s.glow.radius * 1.4), _css_color(s.glow.color, s.glow.alpha)))
        if shadows:
            st = [x for x in st if not x.startswith("box-shadow")]
            st.append("box-shadow:%s" % ", ".join(shadows))
        inner = ""
        if s.body:
            b = s.body
            l, t, r, bb = b.inset
            just = {"t": "flex-start", "ctr": "center", "b": "flex-end"}[b.anchor]
            ps = []
            for p in b.paras:
                al = {"l": "left", "ctr": "center", "r": "right"}[p.align]
                runs = "".join(
                    '<span style="font-family:%s;font-size:%s;color:%s;letter-spacing:%s">%s</span>' % (
                        FONT_PREVIEW.get(r_.font, "'Selawik'"), px(r_.size), _css_color(r_.color, r_.alpha),
                        px(r_.spc), html.escape(r_.text)) for r_ in p.runs)
                ps.append('<div class="p" style="text-align:%s">%s</div>' % (al, runs))
            inner = ('<div class="t%s" style="justify-content:%s;padding:%s %s %s %s;%s">%s</div>' % (
                " nw" if not b.wrap else "", just, px(t), px(r), px(bb), px(l),
                "overflow:hidden" if b.clip else "", "".join(ps)))
        parts.append('<div class="s" title="%s" style="%s">%s</div>' % (html.escape(s.name), ";".join(st), inner))
    parts.append("</div></body></html>")
    return "".join(parts)
