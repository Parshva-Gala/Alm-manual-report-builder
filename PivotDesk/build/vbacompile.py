"""
Compile errors Excel raises and LibreOffice does not.

Excel compiles a module the first time something in it runs, so an error in
a module the Desk does not touch until a build is found halfway through that
build (3.0 met "Ambiguous name detected: C_LINE" in modPD_Stage while
building the Maturity ladder). LibreOffice's Basic is more forgiving and
compiles all of these, so they are checked here, from the source:

  * a module's member used through the module's name - modPD_X.Name - must
    be Public there ("Method or data member not found");
  * a variable passed to a ByRef parameter of a declared type must be of
    that type ("ByRef argument type mismatch");
  * a For Each variable must be a Variant or an object ("For Each control
    variable must be Variant or Object");
  * a name is declared once in a procedure, parameters included
    ("Duplicate declaration in current scope");
  * a procedure called as a statement takes its arguments without
    parentheses when there is more than one ("Expected: =").

The duplicate Public names that caused the 3.0 error are checked by
vbalint.analyse.
"""
from __future__ import annotations

import os
import re

from vbalint import logical_lines, split_statements, _split_args

PROC_HEAD = re.compile(r"^((?:Public|Private|Friend|Static)\s+)*(Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)"
                       r"\s*\((.*)\)(?:\s+As\s+(\w+(?:\.\w+)?)(\(\))?)?\s*$", re.I)
SCALARS = {"string", "long", "integer", "double", "single", "boolean", "date", "currency", "byte", "longlong",
           "longptr", "decimal"}
KEYWORDS = {"if", "elseif", "else", "case", "select", "for", "next", "do", "loop", "while", "wend", "with", "end",
            "exit", "set", "let", "call", "dim", "redim", "const", "static", "private", "public", "on", "resume",
            "goto", "return", "debug", "option", "erase", "print", "open", "close", "get", "put", "input",
            "line", "stop", "err", "mid", "lset", "rset"}


def _params(plist):
    out = []
    for raw in _split_args(plist or ""):
        p = raw.strip()
        if not p:
            continue
        m = re.match(r"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?(\w+)(\(\))?(?:\s+As\s+(\w+(?:\.\w+)?))?",
                     p, re.I)
        if not m:
            continue
        out.append(dict(name=m.group(4), byref=not (m.group(2) or "").lower().startswith("byval"),
                        type=(m.group(6) or "variant").lower(), array=bool(m.group(5)),
                        optional=bool(m.group(1)), paramarray=bool(m.group(3))))
    return out


def _decls(rest):
    """Dim a As Long, b() As String, c  ->  [(name, type, is_array)]"""
    out = []
    for part in _split_args(rest):
        m = re.match(r"^\s*(?:WithEvents\s+)?(\w+)(\s*\([^)]*\))?(?:\s+As\s+(?:New\s+)?(\w+(?:\.\w+)?))?", part, re.I)
        if m:
            out.append((m.group(1), (m.group(3) or "variant").lower(), bool(m.group(2))))
    return out


class Proc:
    def __init__(self, mod, name, public, params, ln, kind="sub"):
        self.mod, self.name, self.public, self.params, self.ln = mod, name, public, params, ln
        self.kind = kind                # sub, function, property
        self.labels = set()
        self.locals = {}                # lname -> (type, is_array)
        self.body = []                  # (ln, statement)


def parse(modules):
    procs = {}                          # (mod, lname) -> Proc
    mod_vars = {}                       # mod -> {lname: (type, array, public, is_const)}
    for mod, code in modules.items():
        mv = mod_vars.setdefault(mod, {})
        cur = None
        for ln, text in logical_lines(code):
            lab = re.match(r"^\s*([A-Za-z_]\w*):(?!=)(\s|$)", text)
            if lab and cur is not None:
                cur.labels.add(lab.group(1).lower())
            for st in split_statements(text):
                low = st.lower()
                pm = PROC_HEAD.match(st)
                if pm and not low.startswith(("end ", "exit ")) and " declare " not in " " + low:
                    public = "private" not in (pm.group(1) or "").lower()
                    cur = Proc(mod, pm.group(3), public, _params(pm.group(4)), ln, pm.group(2).split()[0].lower())
                    procs[(mod, pm.group(3).lower())] = cur
                    for p in cur.params:
                        cur.locals.setdefault(p["name"].lower(), (p["type"], p["array"]))
                    continue
                if re.match(r"^end\s+(sub|function|property)\b", low):
                    cur = None
                    continue
                if cur is None:
                    m = re.match(r"^(Public|Private|Dim|Global)\s+(Const\s+)?(?!Sub\b|Function\b|Property\b|Declare\b|"
                                 r"Enum\b|Type\b)(.*)$", st, re.I)
                    if m:
                        for nm, ty, arr in _decls(m.group(3)):
                            mv[nm.lower()] = (ty, arr, m.group(1).lower() in ("public", "global"), bool(m.group(2)))
                    continue
                m = re.match(r"^(?:Dim|Static|ReDim(?:\s+Preserve)?)\s+(.*)$", st, re.I)
                if m and not low.startswith("redim"):
                    for nm, ty, arr in _decls(m.group(1)):
                        cur.locals.setdefault(nm.lower(), (ty, arr))
                m = re.match(r"^Const\s+(\w+)(?:\s+As\s+(\w+))?", st, re.I)
                if m:
                    cur.locals.setdefault(m.group(1).lower(), ((m.group(2) or "variant").lower(), False))
                cur.body.append((ln, st))
    return procs, mod_vars


def _resolve(procs, mod_names, caller, qual, name):
    lname = name.lower()
    if qual:
        if qual.lower() not in mod_names:
            return None
        return procs.get((mod_names[qual.lower()], lname))
    own = procs.get((caller, lname))
    if own:
        return own
    pubs = [p for (m, n), p in procs.items() if n == lname and p.public]
    return pubs[0] if len(pubs) == 1 else None


def _type_of(proc, mod_vars, mod, ident):
    li = ident.lower()
    if li in proc.locals:
        return proc.locals[li]
    if li in mod_vars.get(mod, {}):
        ty, arr, pub, const = mod_vars[mod][li]
        return None if const else (ty, arr)
    for m, vs in mod_vars.items():
        if li in vs and vs[li][2]:
            ty, arr, pub, const = vs[li]
            return None if const else (ty, arr)
    return None                          # a function, a constant, a property: an expression


def problems(modules):
    out = []
    procs, mod_vars = parse(modules)
    std = {m for m in modules if not (m == "ThisWorkbook" or m.startswith("Sheet") or m == "PDTrace")}
    mod_names = {m.lower(): m for m in modules}
    public_of = {}
    for (m, n), p in procs.items():
        if p.public:
            public_of.setdefault(m, set()).add(n)
    for m, vs in mod_vars.items():
        for n, (_, _, pub, _) in vs.items():
            if pub:
                public_of.setdefault(m, set()).add(n)

    for (mod, _), proc in procs.items():
        # duplicate declarations in one procedure
        seen = {}
        for p in proc.params:
            if p["name"].lower() in seen:
                out.append((mod, proc.ln, "%s: parameter %s declared twice" % (proc.name, p["name"])))
            seen[p["name"].lower()] = proc.ln
        for ln, st in proc.body:
            m = re.match(r"^(?:Dim|Static)\s+(.*)$", st, re.I)
            if m:
                for nm, _, _ in _decls(m.group(1)):
                    if nm.lower() in seen:
                        out.append((mod, ln, "%s: %s declared twice in one procedure" % (proc.name, nm)))
                    if nm.lower() == proc.name.lower():
                        out.append((mod, ln, "%s: a local variable takes the procedure's own name" % proc.name))
                    seen[nm.lower()] = ln

        # labels, GoTo, Exit, For / Next
        labels = proc.labels
        fors = []
        for ln, st in proc.body:
            s2 = re.sub(r'"[^"]*"', '""', st)
            gm = re.match(r"^(?:On\s+Error\s+)?GoTo\s+(\w+)$", s2, re.I) or re.search(r"\bThen\s+GoTo\s+(\w+)$", s2, re.I)
            if gm and gm.group(1) != "0" and gm.group(1).lower() not in labels:
                out.append((mod, ln, "%s: GoTo %s - no such label" % (proc.name, gm.group(1))))
            rm = re.match(r"^Resume\s+(\w+)$", s2, re.I)
            if rm and rm.group(1).lower() != "next" and rm.group(1).lower() not in labels:
                out.append((mod, ln, "%s: Resume %s - no such label" % (proc.name, rm.group(1))))
            for em in re.finditer(r"\bExit\s+(Sub|Function|Property)\b", s2, re.I):
                if em.group(1).lower() != proc.kind:
                    out.append((mod, ln, "%s: Exit %s inside a %s" % (proc.name, em.group(1), proc.kind)))
            fm = re.match(r"^For\s+(?:Each\s+)?(\w+)\b", s2, re.I)
            if fm and not re.search(r"\bNext\b", s2, re.I):
                fors.append(fm.group(1).lower())
            nm = re.match(r"^Next(?:\s+(\w+))?$", s2, re.I)
            if nm and fors:
                open_var = fors.pop()
                if nm.group(1) and nm.group(1).lower() != open_var:
                    out.append((mod, ln, "%s: Next %s closes For %s" % (proc.name, nm.group(1), open_var)))

        for ln, st in proc.body:
            s2 = re.sub(r'"[^"]*"', '""', st)
            # a Sub used as a value
            for qual, name, args, pos in _fn_calls(s2):
                target = _resolve(procs, mod_names, mod, qual, name)
                if target and target.kind == "sub":
                    whole = re.match(r"^(?:Call\s+)?(?:\w+\.)?\w+\s*\(.*\)$", s2, re.I)
                    if not (whole and pos == s2.find(name)) or (not s2.lower().startswith("call") and len(args) > 1):
                        if not (whole and pos == s2.find(name)):
                            out.append((mod, ln, "%s is a Sub, used as a value" % name))
            # For Each control variables
            fm = re.match(r"^For\s+Each\s+(\w+)\s+In\b", s2, re.I)
            if fm:
                t = _type_of(proc, mod_vars, mod, fm.group(1))
                if t and t[0] in SCALARS:
                    out.append((mod, ln, "For Each variable %s is %s - it must be Variant or an object" %
                                (fm.group(1), t[0])))
            # modX.Name must be public in modX
            for qm in re.finditer(r"(?<![\w.])(modPD_\w+)\.(\w+)", s2):
                target = mod_names.get(qm.group(1).lower())
                if target and target in std and qm.group(2).lower() not in public_of.get(target, set()):
                    out.append((mod, ln, "%s.%s is not Public in %s" % (qm.group(1), qm.group(2), target)))
            # Name(a, b) as a statement
            pm = re.match(r"^((?:\w+\.)*(\w+))\s*\(", s2)
            close = _close(s2, pm.end()) if pm else -1
            if pm and pm.group(2).lower() not in KEYWORDS and close == len(s2.rstrip()) - 1:
                if len([a for a in _split_args(s2[pm.end():close]) if a.strip()]) > 1:
                    out.append((mod, ln, "%s(...) with more than one argument as a statement needs no "
                                         "parentheses, or Call" % pm.group(1)))
            # ByRef argument types
            for qual, name, args in _calls(s2):
                target = _resolve(procs, mod_names, mod, qual, name)
                if not target:
                    continue
                named = {}
                positional = []
                for a in args:
                    nm = re.match(r"^\s*(\w+)\s*:=\s*(.*)$", a)
                    if nm:
                        named[nm.group(1).lower()] = nm.group(2)
                    else:
                        positional.append(a)
                pairs = []
                for i, a in enumerate(positional):
                    if i < len(target.params) and not target.params[i]["paramarray"]:
                        pairs.append((target.params[i], a))
                for p in target.params:
                    if p["name"].lower() in named:
                        pairs.append((p, named[p["name"].lower()]))
                for k in named:
                    if k not in {p["name"].lower() for p in target.params}:
                        out.append((mod, ln, "%s has no argument named %s" % (target.name, k)))
                for p, a in pairs:
                    if not p["byref"] or p["type"] == "variant" and not p["array"]:
                        continue
                    am = re.match(r"^\s*(\w+)\s*$", a)
                    if not am:
                        continue                     # an expression: passed as a temporary
                    t = _type_of(proc, mod_vars, mod, am.group(1))
                    if t is None:
                        continue
                    if t[0] != p["type"] or t[1] != p["array"]:
                        out.append((mod, ln, "ByRef argument type mismatch: %s (%s%s) passed to %s's %s (%s%s)" % (
                            am.group(1), t[0], "()" if t[1] else "", target.name, p["name"], p["type"],
                            "()" if p["array"] else "")))
    return out


def _close(s, start):
    """The index of the ) closing the ( just before start."""
    depth, i = 1, start
    while i < len(s) and depth:
        if s[i] == "(":
            depth += 1
        elif s[i] == ")":
            depth -= 1
        i += 1
    return i - 1 if depth == 0 else -1


def _fn_calls(s2):
    """(qualifier, name, [args], position) for every Name( ... ) in a statement."""
    out = []
    for fm in re.finditer(r"(?<![\w.])(?:(\w+)\.)?(\w+)\(", s2):
        close = _close(s2, fm.end())
        out.append((fm.group(1), fm.group(2), [a for a in _split_args(s2[fm.end():close]) if a.strip()],
                    fm.start(2)))
    return out


def _calls(s2):
    """(qualifier, name, [args]) for every call in a statement, both forms."""
    found = []
    m = re.match(r"^(?:Call\s+)?(?:(\w+)\.)?(\w+)(?:\s+(?!=)(.*))?$", s2)
    if m and m.group(2).lower() not in KEYWORDS and not re.match(r"^\s*[\w.]+\s*(\(.*\))?\s*=", s2):
        rest = m.group(3) or ""
        if s2.lower().startswith("call ") and rest.strip().startswith("("):
            rest = rest.strip()[1:-1]
        found.append((m.group(1), m.group(2), [a for a in _split_args(rest) if a.strip()]))
    for fm in re.finditer(r"(?<![\w.])(?:(\w+)\.)?(\w+)\(", s2):
        depth, i = 1, fm.end()
        while i < len(s2) and depth:
            if s2[i] == "(":
                depth += 1
            elif s2[i] == ")":
                depth -= 1
            i += 1
        found.append((fm.group(1), fm.group(2), [a for a in _split_args(s2[fm.end():i - 1]) if a.strip()]))
    return found


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    vba = os.path.join(os.path.dirname(here), "src", "vba")
    mods = {}
    for f in sorted(os.listdir(vba)):
        if f.endswith((".bas", ".cls")):
            with open(os.path.join(vba, f), encoding="cp1252") as fh:
                mods[f.rsplit(".", 1)[0]] = fh.read()
    probs = problems(mods)
    for p in probs:
        print("%s:%s  %s" % p)
    print("%d problem(s)" % len(probs))
    return 1 if probs else 0


if __name__ == "__main__":
    raise SystemExit(main())
