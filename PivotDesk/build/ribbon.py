"""
The Avati tab on Excel's ribbon, for Excel view.

In app view the ribbon is hidden and the Desk is the interface. Anyone who
prefers Excel's own chrome gets the same actions on a tab of their own,
first on the ribbon, with the Desk's icons. Every button calls one VBA
callback, PD_RibbonClick, which dispatches on the button's id; the build
checks that every id here is one it handles.
"""

from __future__ import annotations

import os

import assets
from design_tokens import EM

REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
PART = "customUI/customUI14.xml"
CALLBACK = "PD_RibbonClick"

#        group        [(id, label, icon, large, tip)]
GROUPS = [
    ("Go to", [
        ("pdDesk", "Desk", "spark", True, "The Desk: what is loaded, what to do next. Ctrl+Shift+D"),
        ("pdFiles", "Files", "files", False, "Every file on the desk and what it was recognised as. Ctrl+Shift+F"),
        ("pdConfig", "Pivot config", "layers", False, "What every framework workbook is built from. Ctrl+Shift+P"),
        ("pdRecon", "Reconciliation", "balance", False, "Outputs against control reports 3 and 6. Ctrl+Shift+R"),
        ("pdLog", "Activity", "pulse", False, "Everything the desk has done, newest first. Ctrl+Shift+A"),
    ]),
    ("Work", [
        ("pdScan", "Scan a folder", "folder", True, "Find every output and control report in a folder."),
        ("pdPick", "Pick files", "upload", True, "Choose files one by one."),
        ("pdBuild", "Build pivots", "pivot", True, "One workbook of live PivotTables per framework switched on."),
        ("pdReconcile", "Reconcile", "check", True, "Compare the outputs with the control reports."),
    ]),
    ("Avati", [
        ("pdApp", "App view", "expand", True, "Hide the ribbon and work from the Desk."),
        ("pdTour", "Tour", "help", True, "A one-minute tour of the Desk. F1"),
    ]),
]


def ids():
    return [b[0] for _, buttons in GROUPS for b in buttons]


def _esc(s):
    return s.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")


def custom_ui():
    """(xml, rels_xml, {part path: png bytes})"""
    images, rels, groups = {}, [], []
    out_dir = os.path.join(assets.OUT, "ribbon")
    os.makedirs(out_dir, exist_ok=True)
    for g, (label, buttons) in enumerate(GROUPS, start=1):
        items = []
        for bid, text, icon, large, tip in buttons:
            png = os.path.join(out_dir, "%s.png" % icon)
            svg = assets.icon_svg(icon, EM[600], stroke=1.9)
            stale = not os.path.exists(png) or not os.path.exists(png + ".svg") or open(png + ".svg").read() != svg
            if stale:
                with open(png + ".svg", "w") as f:
                    f.write(svg)
                page = ('<!doctype html><html><body style="margin:0;background:transparent">'
                        '<div style="width:32px;height:32px">%s</div></body></html>') % svg.replace(
                            'width="24" height="24"', 'width="32" height="32"')
                assets.render(page, png, 32, 32, scale=1)
            rid = "pdImg_%s" % icon
            if rid not in images:
                images[rid] = ("customUI/images/%s.png" % icon, open(png, "rb").read())
                rels.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                            'relationships/image" Target="images/%s.png"/>' % (rid, icon))
            items.append('<button id="%s" label="%s" image="%s" size="%s" onAction="%s" screentip="%s" '
                         'supertip="%s"/>' % (bid, _esc(text), rid, "large" if large else "normal", CALLBACK,
                                              _esc(text), _esc(tip)))
        groups.append('<group id="pdGroup%d" label="%s">%s</group>' % (g, _esc(label), "".join(items)))
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">'
           '<ribbon><tabs><tab id="pdTab" label="Avati" insertBeforeMso="TabHome">%s</tab></tabs></ribbon>'
           '</customUI>' % "".join(groups))
    rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s'
                '</Relationships>' % "".join(rels))
    return xml, rels_xml, dict(images.values())
