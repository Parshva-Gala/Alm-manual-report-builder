"""
A compile-shaped lint for the VBA sources, run on every build.

A module that does not compile does not fail loudly in Excel: the first macro
anyone presses stops on "Variable not defined" or "Block If without End If",
and when that happens inside Workbook_Open the whole desk is dead on arrival.
There is no VBA compiler off Windows, so the build checks what it can:

  * every block closes: Sub/Function/Property, If, For, Do, While, With,
    Select, Type, Enum
  * under Option Explicit, every bare identifier is declared somewhere in
    scope - a local, a parameter, a module-level name, a Public name in any
    module, a procedure - or is part of VBA / Excel / Office
  * every xl*/mso*/vb* constant is one known to exist (the ones the shipped
    code already compiled with, plus a reviewed list)
  * no Public procedure name is declared twice across modules, because a
    shape's OnAction is an unqualified name and two of them is "Ambiguous
    name detected" on every press
  * module-level declarations sit above the first procedure

It is a lint, not a compiler: a clean run makes a compile error unlikely, not
impossible.
"""

from __future__ import annotations

import re
import sys

KEYWORDS = set("""
and as byref byval call case const declare dim do each else elseif empty end enum
erase error exit explicit false for friend function get global gosub goto if implements
in is let lib like loop me mod new next not nothing null on option optional or
paramarray preserve private property public redim rem resume return select set static
step stop sub then to true type until wend while with withevents xor base compare text
binary module any typeof addressof lset rset open close print write input output append put get seek lock unlock
access read shared lock len line object
""".split())

TYPES = set("""
string long double boolean integer single date currency byte variant object longptr
longlong decimal collection worksheet workbook range shape shapes listobject pivottable
pivotcache pivotfield pivotitem pivotcaches slicercache slicer hyperlink tablestyle
window chartobject textframe2 fillformat lineformat font interior borders border
application name names sheets worksheets workbooks dictionary pdtrace picture
""".split())

BUILTINS = set("""
abs array asc ascw atn cbool cbyte ccur cdate cdbl cdec chr chrw cint clng clnglng clngptr
command cos createobject csng cstr curdir cvar cverr date dateadd datediff datepart
dateserial datevalue day ddb dir doevents environ eof err error exp filedatetime filelen
fileattr fix format formatcurrency formatdatetime formatnumber formatpercent freefile
getattr getobject hex hour iif imestatus input inputbox instr instrrev int ipmt irr
isarray isdate isempty iserror ismissing isnull isnumeric isobject join lbound lcase
left len loc lof log ltrim mid minute mirr month monthname msgbox now nper npv oct
partition pmt ppmt pv qbcolor rate replace rgb right rnd round rtrim second seek sgn
shell sin sln space spc split sqr str strcomp strconv string strreverse switch syd tab
tan time timer timeserial timevalue trim typename ubound ucase val vartype weekday
weekdayname year kill mkdir rmdir filecopy name chdir chdrive randomize beep sendkeys
appactivate savesetting getsetting deletesetting debug
application activesheet activewindow activecell activeworkbook workbooks worksheets
sheets thisworkbook cells range rows columns selection names windows addins
charts evaluate intersect union
""".split())

# Constants the shipped code compiled with. New ones are added deliberately.
KNOWN_CONSTANTS = set("""
xlsheetveryhidden xlsheetvisible xlup xltoleft xldown xlcenter xlbottom xltop xledgebottom
xledgeleft xlmedium xlthick xlinsidehorizontal xlfreefloating xlsrcrange xlyes xldatabase
xlrowfield xlcolumnfield xlpagefield xlsum xltabularrow xldonotrepeatlabels xlcalculationmanual
xlcalculationautomatic xlwbatworksheet
msofiledialogfolderpicker msofiledialogfilepicker msoshaperoundedrectangle msoshaperectangle
msofalse msotrue msoshadowstyleoutershadow msoanchormiddle msoaligncenter
vbcrlf vbcr vblf vbtextcompare vbinformation vbexclamation vbquestion vbyesno vbyes vbobjecterror
vbnormalfocus vbtab vbnullstring vbbinarycompare vbokonly vbcritical vbno vbokcancel vbok
vbcancel vbhide vbminimizednofocus vbstring vbdouble vblong vbempty vbnull
xlnone xlcontinuous xlthin xlhairline xlleft xlright xlgeneral xlsheethidden
xlmaximized xlnormal xlpasteformats xllandscape xlportrait
msoshapeoval msoanchortop msoalignleft msoalignright msoanchorbottom msotexteffect1
msothemecolortext1
""".split())

# Constants added in 2.0, each checked against the Excel / Office type
# library by name and value before it went in.
REVIEWED_CONSTANTS = {
    "msoautosizenone": 0,                 # MsoAutoSize
    "xlcolumnstacked": 52,                # XlChartType
    "xllegendpositiontop": -4160,         # XlLegendPosition
    "xlvalue": 2,                         # XlAxisType
    "xlcategory": 1,
    "xlticklabelpositionlow": -4134,      # XlTickLabelPosition
    "xlformatfromrightorbelow": 1,        # XlInsertFormatOrigin
    "xlformatfromleftorabove": 0,
    "xlwait": 2,                          # XlMousePointer
    "xldefault": -4143,
    "xlexpression": 2,                    # XlFormatConditionType
    "xlcellvalue": 1,
    "xlsheetvisible": -1,
    "xledgetop": 8,                       # XlBordersIndex
    "xledgeright": 10,
    "xlinsidevertical": 11,
    "xlpagebreakpreview": 2,              # XlWindowView
    "xlnormalview": 1,
    "xlvalues": -4163,
    "xlpart": 2,
    "xlwhole": 1,
    "msoshapeoval": 9,
    "msogradienthorizontal": 1,
    "msoanchorcenter": 2,
    "msotextorientationhorizontal": 1,
    "vbmodeless": 0,
    "vbnormalnofocus": 4,
    "vbretrycancel": 5,
    "vbyesnocancel": 3,
    "vbdirectory": 16,                    # VbFileAttribute
    "vbdate": 7,                          # VbVarType
    "xlunderlinestylenone": -4142,        # XlUnderlineStyle
    "xlvalidatelist": 3,                  # XlDVType
    "xlvalidateinputonly": 0,
    "xlvalidalertstop": 1,                # XlDVAlertStyle
    "xlbetween": 1,                       # XlFormatConditionOperator
    "xlcount": -4112,                     # XlConsolidationFunction
    "xlaverage": -4106,
    "xlmax": -4136,
    "xlmin": -4139,
    "xlpercentofrow": 6,                  # XlPivotFieldCalculation
    "xlpercentofcolumn": 7,
    "xlpercentoftotal": 8,
    "xlascending": 1,                     # XlSortOrder
    "xldescending": 2,
    "xlcompactrow": 0,                    # XlLayoutRowType
    "xloutlinerow": 2,
    "xlrepeatlabels": 2,                  # XlPivotFieldRepeatLabels
}
KNOWN_CONSTANTS |= set(REVIEWED_CONSTANTS)


def strip_strings_and_comments(line: str) -> str:
    out = []
    i = 0
    in_str = False
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
                out.append('"')
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append('"')
            i += 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    s = "".join(out)
    if re.match(r"^\s*Rem(\s|$)", s, re.I):
        return ""
    return s


CONTINUATION_LIMIT = 24     # VBA: "Too many line continuations" beyond this


def continuation_problems(name, code):
    out = []
    run, start = 0, None
    for n, raw in enumerate(code.split("\n"), 1):
        s = strip_strings_and_comments(raw.rstrip("\r"))
        if len(raw) > 1023:
            out.append((name, n, "line longer than VBA's 1023 characters"))
        if re.search(r"\s_\s*$", s):
            if start is None:
                start = n
            run += 1
        else:
            if run > CONTINUATION_LIMIT:
                out.append((name, start, "%d line continuations in one statement; VBA allows %d" % (run, CONTINUATION_LIMIT)))
            run, start = 0, None
    return out


# LibreOffice runs this project's pure functions in build/lo_run.py, which is
# the only way to execute the VBA off Windows. Its compiler rejects three
# things Excel's accepts, and one rejection sinks the whole library, so they
# are kept out: each has a plain equivalent that means the same in Excel.
LO_RESERVED = {"base"}


def lo_compat_problems(name, code):
    out = []
    for n, text in logical_lines(code):
        m = re.match(r"^\s*(?:ElseIf|If)\s+.*?\s+Then\s+(\S.*?)\s+Else\s+", text, re.I)
        if m:
            then = m.group(1)
            is_call = re.match(r"^\.?[A-Za-z_][\w.]*\s+[^=\s:]", then)
            is_assign = re.match(r"^[A-Za-z_][\w.()\s,\"]*=", then)
            if is_call and not is_assign and not re.match(r"^(Exit|GoTo|Set|Call)\b", then, re.I):
                out.append((name, n, "one-line If calling %s with arguments before Else: write it as a block "
                                      "(LibreOffice cannot parse it)" % then.split()[0]))
        if re.search(r"\bOptional\s+(ByVal\s+|ByRef\s+)?\w+\s+As\s+\w+\s*=\s*Nothing\b", text, re.I):
            out.append((name, n, "Optional ... = Nothing: leave the default off, it is Nothing anyway "
                                 "(LibreOffice cannot parse it)"))
        for w in LO_RESERVED:
            if re.search(r"(?<![\w.\"])%s\b(?!\s*\()" % w, text, re.I) and not re.search(r"Option\s+Base", text, re.I):
                out.append((name, n, "%r is reserved in LibreOffice Basic; use another name" % w))
    return out


def logical_lines(code: str):
    """(first physical line number, text) with continuations joined."""
    buf, start = "", None
    for n, raw in enumerate(code.split("\n"), 1):
        s = strip_strings_and_comments(raw.rstrip("\r"))
        if start is None:
            start = n
        if re.search(r"\s_\s*$", s):
            buf += re.sub(r"\s_\s*$", " ", s)
            continue
        buf += s
        yield start, buf
        buf, start = "", None
    if buf:
        yield start, buf


def split_statements(text: str):
    """Split on ':' statement separators, leaving 'label:' and ':=' alone."""
    parts, cur, i = [], [], 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            j = text.find('"', i + 1)
            j = len(text) - 1 if j < 0 else j
            cur.append(text[i:j + 1])
            i = j + 1
            continue
        if ch == ":" and not (i + 1 < len(text) and text[i + 1] == "="):
            parts.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    parts.append("".join(cur))
    return [p.strip() for p in parts if p.strip()]


IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*[$%&#!@]?")
PROC_RE = re.compile(
    r"^(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*(\((.*)\))?",
    re.I)
DECL_RE = re.compile(r"^(?:Dim|Private|Public|Global|Static|Const|ReDim(?:\s+Preserve)?)\s+(.*)$", re.I)


def base(name: str) -> str:
    return name.rstrip("$%&#!@").lower()


def param_names(plist: str):
    names = []
    depth = 0
    cur = ""
    for ch in plist:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            names.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        names.append(cur)
    out = []
    for p in names:
        p = re.sub(r"=.*$", "", p)
        toks = [t for t in re.split(r"\s+", p.strip()) if t]
        toks = [t for t in toks if t.lower() not in ("optional", "byval", "byref", "paramarray")]
        if toks:
            out.append(base(re.sub(r"\(.*", "", toks[0])))
    return out


def decl_names(rest: str):
    rest = re.sub(r"^(?:Const|WithEvents)\s+", "", rest.strip(), flags=re.I)
    # "Private Const X As Long = 1" / "Dim a As Long, b() As String"
    items, depth, cur = [], 0, ""
    for ch in rest:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            items.append(cur)
            cur = ""
        else:
            cur += ch
    items.append(cur)
    out = []
    for it in items:
        it = it.strip()
        it = re.sub(r"^(?:Const|WithEvents)\s+", "", it, flags=re.I)
        m = IDENT.match(it)
        if m:
            out.append(base(m.group(0)))
    return out


class Module:
    def __init__(self, name, code):
        self.name = name
        self.code = code
        self.procs = {}           # lname -> (name, public?, line)
        self.module_names = set()
        self.public_names = set()
        self.issues = []


def analyse(modules: dict[str, str], known_constants=None, extra_names=None):
    known_constants = set(KNOWN_CONSTANTS) | set(known_constants or [])
    extra = set(n.lower() for n in (extra_names or []))
    mods = {n: Module(n, c) for n, c in modules.items()}
    issues = []
    for n, c in modules.items():
        issues += continuation_problems(n, c)
        issues += lo_compat_problems(n, c)

    # ---- pass 1: declarations and block structure ----------------------------
    for m in mods.values():
        stack = []
        in_proc = False
        seen_proc = False
        for ln, text in logical_lines(m.code):
            for st in split_statements(text):
                low = st.lower()
                if re.match(r"^[A-Za-z_]\w*:$", st + ":") and False:
                    pass
                pm = PROC_RE.match(st)
                if pm and not low.startswith(("end ", "exit ")) and not re.match(r"^(public|private)\s+declare", low):
                    kind = pm.group(1).split()[0].lower()
                    name = pm.group(2)
                    is_pub = not low.startswith("private")
                    if base(name) in m.procs and kind != "property":
                        issues.append((m.name, ln, "procedure %s declared twice" % name))
                    m.procs[base(name)] = (name, is_pub, ln)
                    if is_pub:
                        m.public_names.add(base(name))
                    stack.append(("proc:" + kind, ln))
                    in_proc = True
                    seen_proc = True
                    continue
                em = re.match(r"^end\s+(sub|function|property)\b", low)
                if em:
                    if not stack or not stack[-1][0].startswith("proc"):
                        issues.append((m.name, ln, "End %s with no open procedure (open: %s)" %
                                       (em.group(1), stack[-1] if stack else None)))
                    else:
                        stack.pop()
                    in_proc = False
                    continue
                if not in_proc:
                    dm = DECL_RE.match(st)
                    if dm and not re.match(r"^(public|private)\s+(sub|function|property|declare|type|enum)\b", low):
                        if seen_proc:
                            issues.append((m.name, ln, "module-level declaration after the first procedure"))
                        names = decl_names(dm.group(1))
                        m.module_names.update(names)
                        if low.startswith(("public", "global")):
                            m.public_names.update(names)
                    tm = re.match(r"^(?:(public|private)\s+)?(type|enum)\s+(\w+)", low)
                    if tm:
                        m.module_names.add(tm.group(3))
                        if tm.group(1) != "private":
                            m.public_names.add(tm.group(3))
                        stack.append((tm.group(2), ln))
                        continue
                    if re.match(r"^end\s+(type|enum)\b", low):
                        stack.pop()
                        continue
                    if stack and stack[-1][0] in ("type", "enum"):
                        im = IDENT.match(st)
                        if im and stack[-1][0] == "enum":
                            m.module_names.add(base(im.group(0)))
                            m.public_names.add(base(im.group(0)))
                        continue
                    continue
                # blocks inside procedures
                if re.match(r"^if\b", low) and re.search(r"\bthen$", low):
                    stack.append(("if", ln))
                elif re.match(r"^elseif\b|^else$", low):
                    if not stack or stack[-1][0] != "if":
                        issues.append((m.name, ln, "Else/ElseIf outside a block If"))
                elif re.match(r"^end\s+if\b", low):
                    if not stack or stack[-1][0] != "if":
                        issues.append((m.name, ln, "End If without block If (open: %s)" % (stack[-1] if stack else None)))
                    else:
                        stack.pop()
                elif re.match(r"^for\b", low):
                    stack.append(("for", ln))
                elif re.match(r"^next\b", low):
                    if not stack or stack[-1][0] != "for":
                        issues.append((m.name, ln, "Next without For (open: %s)" % (stack[-1] if stack else None)))
                    else:
                        stack.pop()
                elif re.match(r"^do\b", low):
                    stack.append(("do", ln))
                elif re.match(r"^loop\b", low):
                    if not stack or stack[-1][0] != "do":
                        issues.append((m.name, ln, "Loop without Do"))
                    else:
                        stack.pop()
                elif re.match(r"^while\b", low):
                    stack.append(("while", ln))
                elif re.match(r"^wend\b", low):
                    if not stack or stack[-1][0] != "while":
                        issues.append((m.name, ln, "Wend without While"))
                    else:
                        stack.pop()
                elif re.match(r"^with\b", low):
                    stack.append(("with", ln))
                elif re.match(r"^end\s+with\b", low):
                    if not stack or stack[-1][0] != "with":
                        issues.append((m.name, ln, "End With without With"))
                    else:
                        stack.pop()
                elif re.match(r"^select\s+case\b", low):
                    stack.append(("select", ln))
                elif re.match(r"^end\s+select\b", low):
                    if not stack or stack[-1][0] != "select":
                        issues.append((m.name, ln, "End Select without Select"))
                    else:
                        stack.pop()
        for kind, ln in stack:
            issues.append((m.name, ln, "unclosed %s" % kind))

    # Public procedure names across modules: OnAction is unqualified.
    owners = {}
    for m in mods.values():
        for lname, (name, pub, ln) in m.procs.items():
            if pub:
                owners.setdefault(lname, []).append((m.name, ln))
    for lname, where in owners.items():
        if len(where) > 1 and not all(w[0] in ("ThisWorkbook",) for w in where):
            # Same public name in two standard modules is ambiguous when called
            # unqualified. Document/class modules do not collide.
            std = [w for w in where if w[0] not in ("ThisWorkbook", "PDTrace") and not w[0].startswith("Sheet")]
            if len(std) > 1:
                issues.append((std[1][0], std[1][1], "Public %s also declared in %s" % (lname, std[0][0])))

    global_names = set(extra) | set(n.lower() for n in mods)
    for m in mods.values():
        global_names |= m.public_names

    # ---- pass 2: identifiers ---------------------------------------------------
    for m in mods.values():
        local = set()
        labels = set()
        in_proc = False
        proc_lines = []
        for ln, text in logical_lines(m.code):
            for st in split_statements(text):
                low = st.lower()
                pm = PROC_RE.match(st)
                if pm and not low.startswith(("end ", "exit ")) and not re.match(r"^(public|private)\s+declare", low):
                    in_proc = True
                    local = set(param_names(pm.group(4) or ""))
                    local.add(base(pm.group(2)))
                    labels = set()
                    proc_lines = []
                    continue
                if re.match(r"^end\s+(sub|function|property)\b", low):
                    # check the procedure body now that every Dim is known
                    for pln, pst in proc_lines:
                        check_statement(m, pln, pst, local, labels, global_names, known_constants, issues)
                    in_proc = False
                    continue
                if not in_proc:
                    continue
                if re.match(r"^[A-Za-z_]\w*$", st) and text.strip().startswith(st) and text.strip()[len(st):].startswith(":"):
                    labels.add(st.lower())
                    continue
                dm = re.match(r"^(?:Dim|Static|Const|ReDim(?:\s+Preserve)?)\s+(.*)$", st, re.I)
                if dm:
                    local.update(decl_names(dm.group(1)))
                proc_lines.append((ln, st))
    return issues


def check_statement(m, ln, st, local, labels, global_names, known_constants, issues):
    low = st.lower()
    # the part of a declaration after "As" names a type; the name itself is fine
    s = st
    if re.match(r"^(dim|static|const|redim)\b", low):
        s = re.sub(r"\bAs\s+(New\s+)?[\w.]+", " ", s, flags=re.I)
        # drop declared names, keep initialisers and bounds
        s = re.sub(r"^(Dim|Static|Const|ReDim(\s+Preserve)?)\s+", "", s, flags=re.I)
    s = re.sub(r"\bAs\s+(New\s+)?[\w.]+", " ", s, flags=re.I)
    # jump targets
    for tm in re.finditer(r"\b(?:GoTo|GoSub|Resume)\s+([A-Za-z_]\w*)", s, re.I):
        if tm.group(1).lower() not in labels and tm.group(1).lower() not in ("next",) and tm.group(1) != "0":
            pass
    s = re.sub(r"\b(?:GoTo|GoSub)\s+[A-Za-z_]\w*", " ", s, flags=re.I)
    s = re.sub(r"\bResume\s+(Next|[A-Za-z_]\w*)", " ", s, flags=re.I)
    s = re.sub(r'"[^"]*"', '""', s)
    s = re.sub(r"#[^#]*#", "0", s)              # date literals
    s = re.sub(r"&H[0-9A-Fa-f]+&?", "0", s)
    for mt in IDENT.finditer(s):
        tok = mt.group(0)
        start = mt.start()
        prev = s[:start].rstrip()
        if prev.endswith(".") and not prev.endswith(".."):
            continue                              # member access
        after = s[mt.end():].lstrip()
        if after.startswith(":="):
            continue                              # named argument
        name = base(tok)
        if name in KEYWORDS or name in TYPES or name in BUILTINS:
            continue
        if name in local or name in labels or name in m.module_names or name in m.procs:
            continue
        if name in global_names:
            continue
        if re.match(r"^(xl|mso|vb)[a-z0-9]+$", name):
            if name not in known_constants:
                issues.append((m.name, ln, "unknown constant %s" % tok))
            continue
        if re.match(r"^\d", tok):
            continue
        issues.append((m.name, ln, "not declared: %s   [%s]" % (tok, st[:90])))


# ---------------------------------------------------------------------------
#  Arity: a call to one of the project's own procedures passes a number of
#  arguments the signature accepts ("Wrong number of arguments" is a compile
#  error, and it is the one a rename most often leaves behind).
# ---------------------------------------------------------------------------
def _split_args(text):
    args, depth, cur, in_str = [], 0, "", False
    for ch in text:
        if ch == '"':
            in_str = not in_str
        if not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            if ch == "," and depth == 0:
                args.append(cur)
                cur = ""
                continue
        cur += ch
    if cur.strip() or args:
        args.append(cur)
    return [a for a in args]


def signatures(modules):
    sigs = {}
    for mod, code in modules.items():
        for ln, text in logical_lines(code):
            for st in split_statements(text):
                pm = PROC_RE.match(st)
                if not pm or st.lower().startswith(("end ", "exit ")):
                    continue
                name = pm.group(2).lower()
                params = [x for x in _split_args(pm.group(4) or "") if x.strip()]
                req = sum(1 for x in params if not re.match(r"\s*(Optional|ParamArray)\b", x, re.I))
                has_pa = any(re.match(r"\s*ParamArray\b", x, re.I) for x in params)
                mx = 10 ** 6 if has_pa else len(params)
                kind = pm.group(1).split()[0].lower()
                public = not st.lower().startswith("private")
                sigs.setdefault(name, []).append((mod, req, mx, kind, public))
    return sigs


def arity_problems(modules):
    sigs = signatures(modules)
    out = []
    for mod, code in modules.items():
        for ln, text in logical_lines(code):
            for st in split_statements(text):
                if PROC_RE.match(st) or re.match(r"^\s*(Dim|Private|Public|Const|Declare)\b", st, re.I):
                    continue
                s2 = re.sub(r'"[^"]*"', '""', st)
                # statement form:  [Call] [Module.]Name arg, arg
                m = re.match(r"^(?:Call\s+)?(?:(\w+)\.)?(\w+)(?:\s+(.*))?$", s2)
                if m and m.group(2).lower() in sigs and not re.match(r"^\s*\w+(\.\w+)?\s*=", s2):
                    rest = m.group(3)
                    if rest is not None and rest.strip().startswith("="):
                        rest = None
                    if rest is not None and rest.strip().startswith("(") and s2.lower().startswith("call "):
                        rest = rest.strip()[1:-1]
                    n = len([a for a in _split_args(rest or "") if a.strip()]) if rest else 0
                    if not _fits(sigs, m.group(1), m.group(2), n, modules, caller=mod):
                        out.append((mod, ln, "%s called with %d argument(s)" % (m.group(2), n)))
                    continue
                # function form inside expressions:  Name( ... )
                for fm in re.finditer(r"(?<![\w.])(?:(\w+)\.)?(\w+)\(", s2):
                    name = fm.group(2).lower()
                    if name not in sigs:
                        continue
                    depth, i = 1, fm.end()
                    while i < len(s2) and depth:
                        if s2[i] == "(":
                            depth += 1
                        elif s2[i] == ")":
                            depth -= 1
                        i += 1
                    inner = s2[fm.end():i - 1]
                    n = len([a for a in _split_args(inner) if a.strip()])
                    if not _fits(sigs, fm.group(1), fm.group(2), n, modules, as_index=True, caller=mod):
                        out.append((mod, ln, "%s( ) called with %d argument(s)" % (fm.group(2), n)))
    return out


def _fits(sigs, qual, name, n, modules, as_index=False, caller=None):
    cands = sigs.get(name.lower(), [])
    if qual:
        if qual not in modules:
            return True                     # a member of an object, not ours
        cands = [c for c in cands if c[0] == qual]
    else:
        # VBA's scope: this module's own, then any Public one elsewhere
        own = [c for c in cands if c[0] == caller]
        cands = own or [c for c in cands if c[4]]
    if not cands:
        return True
    for mod, req, mx, kind, public in cands:
        if req <= n <= mx:
            return True
        # Name(0) on a function returning an array/collection indexes the result
        if as_index and req == 0 and kind == "function":
            return True
    return False


def main(paths):
    mods = {}
    for p in paths:
        name = re.sub(r"\.(bas|cls)$", "", p.split("/")[-1])
        mods[name] = open(p, encoding="cp1252").read()
    issues = analyse(mods) + arity_problems(mods)
    for mod, ln, msg in issues:
        print("%s:%d  %s" % (mod, ln, msg))
    print("%d issue(s)" % len(issues))
    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
