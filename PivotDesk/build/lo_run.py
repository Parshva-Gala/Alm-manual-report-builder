"""
Runs the workbook's own VBA - the shipped file, not the sources - in
LibreOffice's VBA-compatible Basic, the one place it can execute off Windows.

  1. The whole project must compile there. One module LibreOffice cannot parse
     silently disables every macro in it, so this is checked first, with a
     sentinel function, and the lint keeps the three constructs it rejects out.
  2. modPD_Stage.TenorDays is run on bucket labels of every shape and must
     agree with build/tenor.py, which orders the buckets in the previews.
  3. A handful of other pure functions are run against known answers.

LibreOffice has no Scripting.Dictionary (it is a Windows COM object), so
anything built on one cannot run here; that stays on the Windows walk-through.

usage:  python3 lo_run.py ../dist/Avati.xlsm
"""
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import uno

import labels
import lo_check
import tenor

LABELS = ["UPTO 1 MONTH", "1 - 3 MONTHS", "3 - 6 MONTHS", "6 MONTHS - 1 YEAR", "1 - 3 YEARS", "3 - 5 YEARS",
          "OVER 5 YEARS", "NON MATURITY", "(no bucket)", "OVERNIGHT", "2-7 DAYS", "8 DAYS - 1 MONTH",
          "LESS THAN 6 MONTHS", "6 MONTHS TO 1 YEAR", "MORE THAN 1 YEAR", "> 5Y", "3M-6M", "1Y+", "11",
          "01. UPTO 1 MONTH", "02. 1 - 3 MONTHS", "Up to 30 days", "31-90 Days", "Non-Maturity", "Undated",
          "6-12 MONTHS", "1.5 YEARS", "Demand", "At call", "", "1 WEEK", "2W-1M", "5+ YEARS",
          "Within 7 days", "BEYOND 10 YEARS", "O/N", "Perpetual"]

# (expression, expected) - each expression is VBA evaluated in the project
KNOWN = [
    ('modPD_Util.Compact(1234567)', "1.2 m"),
    ('modPD_Util.Compact(-30162552568.72)', "-30.2 bn"),
    ('modPD_Util.Compact(45210)', "45.2 k"),
    ('modPD_Util.Compact(950)', "950"),
    ('modPD_Util.MidTrim("LCR_Output_30Nov2025_final.xlsx", 20)', "LCR_Output…inal.xlsx"),
    ('modPD_Util.MidTrim("short.xlsx", 20)', "short.xlsx"),
    ('modPD_Util.AsOfText(45991)', "30 Nov 2025"),
    ('modPD_Recipe.SplitList("Type, Line ,, Subline", ",").count', "3"),
    ('modPD_Recipe.SplitList("Type, Line ,, Subline", ",").Item(2)', "Line"),
    ('modPD_Recipe.SplitList("Type, Line ,, Subline", ",").Item(3)', "Subline"),
    ('Num("1,500,000")', "1500000|True"),
    ('Num("1.5m")', "1500000|True"),
    ('Num("2bn")', "2000000000|True"),
    ('Num("750k")', "750000|True"),
    ('Num("25%")', ".25|True"),
    ('Num("(5)")', "0|False"),
    ('Num("m")', "0|False"),
    ('modPD_Recipe.UnitFormat("millions")', "#,##0.0,,;[Red](#,##0.0,,);-"),
    ('modPD_Recipe.UnitFormat("")', "#,##0;[Red](#,##0);-"),
    ('modPD_Stage.PeriodOf(46082, "year", "(none)")', "2026"),
    ('modPD_Stage.PeriodOf(46082, "quarter", "(none)")', "2026 Q1"),
    ('modPD_Stage.PeriodOf(46296, "quarter", "(none)")', "2026 Q4"),
    ('Format$(modPD_Stage.PeriodOf(46082, "month", "(none)"), "yyyy-mm-dd")', "2026-03-01"),
    ('modPD_Stage.PeriodOf(Empty, "year", "(none)")', "(none)"),
    ('modPD_Stage.StepOf(1234567, 1000000, False, "(none)")', "1000000"),
    ('modPD_Stage.StepOf(-1234567, 1000000, True, "(none)")', "1000000"),
    ('modPD_Stage.StepOf(-1234567, 1000000, False, "(none)")', "-2000000"),
    ('modPD_Stage.StepOf("n/a", 1000000, False, "(none)")', "(none)"),
]

HARNESS = '''Option VBASupport 1
Function Sentinel()
    Sentinel = 42
End Function
Function Tidy(ByVal lbl As String, ByVal mode As Long) As String
    Tidy = modPD_Util.TidyLabel(lbl, mode, "QNB, CIB")
End Function
Function Num(ByVal s As String) As String
    Dim ok As Boolean, d As Double
    d = modPD_Recipe.NumOf(s, ok)
    Num = Trim$(Str$(d)) & "|" & IIf(ok, "True", "False")
End Function
Function Tenor(ByVal lbl As String) As String
    Dim s As String, d As Double
    d = modPD_Stage.TenorDays(lbl, s)
    Tenor = Trim$(Str$(d)) & "|" & s
End Function
'''


# modPD_Build's family sort is private; its source is copied into the harness
# and run on a family where every tie-break matters.
SORT_TEST = '''
Function SortTest() As String
    Dim idx(0 To 4) As Long, heads(0 To 4) As String, w(0 To 4) As Double, g(0 To 4) As Double, i As Long
    heads(0) = "A": w(0) = 4: g(0) = 10
    heads(1) = "C": w(1) = 10: g(1) = 10
    heads(2) = "B": w(2) = 5: g(2) = 20
    heads(3) = "A": w(3) = 6: g(3) = 10
    heads(4) = "B": w(4) = 15: g(4) = 20
    For i = 0 To 4
        idx(i) = i
    Next i
    MergeSort idx, heads, w, g
    For i = 0 To 4
        SortTest = SortTest & idx(i)
    Next i
End Function
'''


def private_procs(module, names):
    src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "vba", module + ".bas"),
               encoding="cp1252").read()
    out = ""
    for nm in names:
        m = re.search(r"^Private (Sub|Function) %s\(.*?^End \1" % nm, src, re.S | re.M)
        out += "\n" + m.group(0).replace("Private ", "", 1) + "\n"
    return out


def main():
    path = os.path.abspath(sys.argv[1])
    lock = os.path.join(os.path.dirname(path), ".~lock." + os.path.basename(path) + "#")
    if os.path.exists(lock):
        os.remove(lock)
    port = 2009
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    helper = os.environ.get("SOFFICE_HELPER_DIR")
    if helper:
        sys.path.insert(0, helper)
        from office.soffice import get_soffice_env   # sandbox socket shim
        env = get_soffice_env()
    proc = subprocess.Popen(
        ["soffice", "-env:UserInstallation=" + Path(tempfile.mkdtemp(prefix="lo_profile_")).as_uri(),
         "--headless", "--invisible", "--norestore", "--nologo",
         "--accept=socket,host=localhost,port=%d;urp;" % port],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    doc, bad = None, 0
    try:
        ctx = lo_check.connect(port)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0,
                                           (lo_check.prop("Hidden", True), lo_check.prop("MacroExecutionMode", 4)))
        if doc is None:
            print("FAIL  LibreOffice could not open the workbook")
            return 1
        lib = doc.BasicLibraries.getByName("VBAProject")
        body = HARNESS + private_procs("modPD_Build", ("MergeSort", "Before")) + SORT_TEST
        for i, (expr, _) in enumerate(KNOWN):
            body += "Function K%d() As String\n    K%d = CStr(%s)\nEnd Function\n" % (i, i, expr)
        lib.insertByName("zzHarness", body)
        sp = doc.getScriptProvider()

        def call(fn, *args):
            url = "vnd.sun.star.script:VBAProject.zzHarness.%s?language=Basic&location=document" % fn
            return sp.getScript(url).invoke(args, (), ())[0]

        if call("Sentinel") != 42:
            print("FAIL  the project does not compile in LibreOffice - every macro in it is dead there")
            return 1
        print("ok    every module compiles in LibreOffice")

        for lbl in LABELS:
            days, short = call("Tenor", lbl).split("|", 1)
            want = tenor.tenor_days(lbl)
            same = abs(float(days) - want[0]) < 1e-9 and short == want[1]
            bad += not same
            if not same:
                print("FAIL  TenorDays(%r): VBA %s %r, tenor.py %g %r" % (lbl, days, short, want[0], want[1]))
        print("%s  TenorDays agrees with tenor.py on %d labels" % ("ok  " if not bad else "FAIL", len(LABELS)))

        nbad = 0
        for lbl in labels.SAMPLES + ["QNB ALAHLI DEPOSITS", "CIB/QNB SWAP LINES"]:
            for mode in (1, 2, 3):
                got = call("Tidy", lbl, mode)
                want = labels.tidy(lbl, mode, "QNB, CIB")
                if got != want:
                    nbad += 1
                    print("FAIL  TidyLabel(%r, %d): VBA %r, labels.py %r" % (lbl, mode, got, want))
        bad += nbad
        print("%s  TidyLabel agrees with labels.py on %d labels in 3 modes" % ("ok  " if not nbad else "FAIL",
                                                                              len(labels.SAMPLES) + 2))

        got = call("SortTest")
        ok = got == "42301"
        bad += not ok
        print("%s  family sort: biggest family first, then its biggest member, ties A to Z = %r%s" % (
            "ok  " if ok else "FAIL", got, "" if ok else "  (want '42301')"))

        for i, (expr, want) in enumerate(KNOWN):
            got = call("K%d" % i)
            ok = got == want
            bad += not ok
            print("%s  %s = %r%s" % ("ok  " if ok else "FAIL", expr, got, "" if ok else "  (want %r)" % want))
    finally:
        try:
            if doc is not None:
                doc.close(True)
        except Exception:
            pass
        proc.terminate()
        try:
            proc.wait(10)
        except Exception:
            proc.kill()
    print("%d failure(s)" % bad)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
