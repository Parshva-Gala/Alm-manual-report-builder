"""
The tenor reading modPD_Stage.TenorDays does, line for line, so the previews
order buckets exactly as Excel will and the parser can be tested off Windows.
"""

TENOR_NM = 900000.0
TENOR_NONE = 999999.0

UNITS = {}
for words, per, letter in ((("D", "DAY", "DAYS"), 1, "D"), (("W", "WK", "WKS", "WEEK", "WEEKS"), 7, "W"),
                           (("M", "MO", "MOS", "MON", "MTH", "MTHS", "MONTH", "MONTHS"), 30, "M"),
                           (("Y", "YR", "YRS", "YEAR", "YEARS"), 365, "Y")):
    for w in words:
        UNITS[w] = (per, letter)
OVER = {"OVER", "MORE", "ABOVE", "GREATER", "BEYOND", "AFTER"}
UPTO = {"UPTO", "UP", "LESS", "UNDER", "WITHIN", "BELOW"}


def _num(v):
    s = ("%r" % v).rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
    return s


def tenor_days(label):
    """(days, short) - days is -1 when the label names no tenor."""
    u = (label or "").strip().upper()
    if not u or u == "(NO BUCKET)":
        return TENOR_NONE, "—"
    if ("NON" in u and "MAT" in u) or "UNDATED" in u or "NO MATURITY" in u or "INDETERMIN" in u or "PERPETUAL" in u:
        return TENOR_NM, "NM"
    if "OVERNIGHT" in u or u in ("O/N", "ON"):
        return 1, "O/N"
    if "DEMAND" in u or "AT CALL" in u or "SIGHT" in u:
        return 1, "CALL"
    toks, tok, kind = [], "", ""
    over = upto = False
    for ch in u + " ":
        if ch.isdigit() and ch.isascii() or (ch == "." and kind == "n"):
            if kind != "n":
                if tok and kind: toks.append(kind + tok)
                tok, kind = "", "n"
            tok += ch
        elif "A" <= ch <= "Z":
            if kind != "w":
                if tok and kind: toks.append(kind + tok)
                tok, kind = "", "w"
            tok += ch
        else:
            if tok and kind: toks.append(kind + tok)
            tok, kind = "", ""
            if ch in ">+": over = True
            if ch in "<≤": upto = True
    nums, units, uch = [], [], []
    for t in toks:
        if t[0] == "n":
            if len(nums) < 12:
                nums.append(float(t[1:]) if t[1:] not in (".",) else 0.0); units.append(0); uch.append("")
        else:
            w = t[1:]
            per, letter = UNITS.get(w, (0, ""))
            if w in OVER: over = True
            if w in UPTO: upto = True
            if per and nums and units[-1] == 0:
                units[-1], uch[-1] = per, letter
    if not nums:
        return -1, ""
    for i in range(len(nums) - 2, -1, -1):
        if units[i] == 0:
            units[i], uch[i] = units[i + 1], uch[i + 1]
    if units[-1] == 0:
        return -1, ""
    days = [n * p for n, p in zip(nums, units)]
    ilo = min(range(len(days)), key=lambda i: (days[i], i))
    ihi = max(range(len(days)), key=lambda i: (days[i], -i))
    lo, hi = days[ilo], days[ihi]
    f = lambda v: ("%g" % v)
    if over and lo == hi:
        return hi + 0.5, ">" + f(nums[ihi]) + uch[ihi]
    if lo == hi:
        return hi, ("≤" if upto else "") + f(nums[ihi]) + uch[ihi]
    return hi, f(nums[ilo]) + ("" if uch[ilo] == uch[ihi] else uch[ilo]) + "–" + f(nums[ihi]) + uch[ihi]


def bucket_order(labels, avg_days=None):
    """Labels sorted as BucketOrder sorts them; avg_days maps a label to the
    average days to maturity of its rows, for labels with no tenor."""
    avg_days = avg_days or {}

    def key(lbl):
        d, short = tenor_days(lbl)
        if d < 0:
            d = max(0.0, avg_days[lbl]) if lbl in avg_days else TENOR_NM - 1
        return (d, lbl.upper())
    return sorted(labels, key=key)


def short_label(lbl):
    d, short = tenor_days(lbl)
    return short if d >= 0 else lbl.strip().upper()[:5]


if __name__ == "__main__":
    tests = ["UPTO 1 MONTH", "1 - 3 MONTHS", "3 - 6 MONTHS", "6 MONTHS - 1 YEAR", "1 - 3 YEARS", "3 - 5 YEARS",
             "OVER 5 YEARS", "NON MATURITY", "(no bucket)", "OVERNIGHT", "2-7 DAYS", "8 DAYS - 1 MONTH",
             "LESS THAN 6 MONTHS", "6 MONTHS TO 1 YEAR", "MORE THAN 1 YEAR", "> 5Y", "3M-6M", "1Y+", "11",
             "01. UPTO 1 MONTH", "02. 1 - 3 MONTHS", "Up to 30 days", "31-90 Days", "Non-Maturity", "Undated",
             "6-12 MONTHS", "1.5 YEARS", "Demand"]
    for t in tests:
        print("%-22s %10s  %s" % (t, tenor_days(t)[0], tenor_days(t)[1]))
    print(bucket_order(tests))
