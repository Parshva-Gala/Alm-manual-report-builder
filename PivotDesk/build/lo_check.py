"""
Open a workbook in headless LibreOffice and report what it sees of the VBA
project: every module LibreOffice imported, and - for each module - whether
LibreOffice's Basic compiler accepts it (VBA-support mode).

LibreOffice is not Excel. A clean pass here proves the container, the
compression and the dir stream are readable by an implementation that is
not ours, and catches gross syntax slips. It does not prove Excel's
compiler agrees on every construct.

usage:  python3 lo_check.py book.xlsm [probe Module.Proc ...]
"""

import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect(port=2002):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            return resolver.resolve(
                "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % port)
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("soffice did not come up")


def main():
    path = os.path.abspath(sys.argv[1])
    # A run that died leaves LibreOffice's lock beside the file, and the next
    # load then quietly returns nothing.
    lock = os.path.join(os.path.dirname(path), ".~lock." + os.path.basename(path) + "#")
    if os.path.exists(lock):
        os.remove(lock)
    probes = sys.argv[2:]
    port = 2002
    import tempfile
    from pathlib import Path
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    helper = os.environ.get("SOFFICE_HELPER_DIR")
    if helper:
        sys.path.insert(0, helper)
        from office.soffice import get_soffice_env   # sandbox socket shim
        env = get_soffice_env()
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    proc = subprocess.Popen(
        ["soffice", "-env:UserInstallation=" + Path(profile).as_uri(),
         "--headless", "--invisible", "--norestore", "--nologo",
         "--accept=socket,host=localhost,port=%d;urp;" % port],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    try:
        ctx = connect(port)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(path), "_blank", 0,
            (prop("Hidden", True), prop("MacroExecutionMode", 4)))
        libs = doc.BasicLibraries
        print("libraries:", list(libs.getElementNames()))
        for lname in libs.getElementNames():
            libs.loadLibrary(lname)
            lib = libs.getByName(lname)
            for mname in lib.getElementNames():
                src = lib.getByName(mname)
                print("  %-16s %6d chars" % (mname, len(src)))
        sp = doc.getScriptProvider()
        for probe in probes:
            url = "vnd.sun.star.script:VBAProject.%s?language=Basic&location=document" % probe
            try:
                res = sp.getScript(url).invoke((), (), ())
                print("probe %s -> ok %r" % (probe, res[0]))
            except Exception as e:
                print("probe %s -> FAILED %s" % (probe, str(e).splitlines()[0][:300]))
        doc.close(True)
    finally:
        proc.terminate()
        try:
            proc.wait(10)
        except Exception:
            proc.kill()


if __name__ == "__main__":
    main()
