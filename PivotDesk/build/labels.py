"""
modPD_Util.TidyLabel, line for line, so the previews show labels as a build
will and build/lo_run.py can hold the VBA to it.
"""

BUILT_IN = set(("ALM ASF AC AED ATM BHD BV CASA CBE CD CDS CHF CNY COA CR DR EAD ECL EGP EMI EU EUR FCY FVOCI FVTOCI "
                "FVTPL FX GBP GL HQLA ID IFRS INT IRS JOD JPY KSA KWD KYC LC LCR LCS LCY LG LGD LGS MBGL MM NBE NII NIM "
                "NPA NPL NSFR OCI OD ODS OMR OVD PD POS QAR ROA ROE RSF SAR SME SMES SPV TB TD TDS UAE UK US USD "
                "VAT").split())
SMALL = set("A AN AND AS AT BY FOR FROM IN INTO OF ON OR PER THE TO VIA VS WITH".split())
MODES = {"As is": 0, "Drop codes": 1, "Drop codes + title case": 2, "Title case": 3}


def _is_code_piece(p):
    return 0 < len(p) <= 12 and all(c.isdigit() and c.isascii() or "A" <= c <= "Z" or c in "_-" for c in p)


def _has_digit(p):
    return any("0" <= c <= "9" for c in p)


def _all_digits(p):
    return bool(p) and all("0" <= c <= "9" for c in p)


def _no_vowel(u):
    return len(u) > 1 and not any(c in "AEIOUY" for c in u)


def drop_code(t):
    u = t.find("_") + 1
    if u > 2 and u < len(t) and _all_digits(t[:u - 1]):
        return t[u:].strip()
    if "." not in t:
        return t
    parts = t.split(".")
    k = -1
    for i in range(len(parts) - 1):
        if not _is_code_piece(parts[i]):
            break
        if _has_digit(parts[i]):
            k = i
    if k < 1:
        return t
    out = ".".join(parts[k + 1:]).strip()
    return out or t


def fix_ampersand(t):
    out = []
    for i, ch in enumerate(t):
        if ch == "&":
            lft = t[i - 1] if i > 0 else " "
            rgt = t[i + 1] if i < len(t) - 1 else " "
            if (lft == " ") != (rgt == " "):
                ch = " & "
        out.append(ch)
    s = "".join(out)
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip()


def kept(u, extra=""):
    if u in BUILT_IN:
        return True
    return u in set(extra.upper().replace(",", " ").replace(";", " ").split())


def case_piece(p, first, after_apos, whole, extra):
    u = p.upper()
    if after_apos:
        return p.lower()
    if kept(u, extra) or _no_vowel(u):
        return u
    if len(u) == 1 and not (whole and u == "A"):
        return u
    if whole and not first and u in SMALL:
        return p.lower()
    return p[:1].upper() + p[1:].lower()


def title_word(w, first, only, extra):
    if w != w.upper() or _has_digit(w):
        return w
    out, piece, after, n = "", "", False, 0
    for ch in list(w) + [""]:
        if ch in ("-", "/", ".", "'", "(", ")", "&", ""):
            if piece:
                n += 1
                out += case_piece(piece, first and n == 1, after, len(w) == len(piece) and not only, extra)
            after = ch == "'"
            out += ch
            piece = ""
        else:
            piece += ch
    return out


def title_words(t, extra=""):
    words = t.split(" ")
    return " ".join(title_word(w, i == 0, len(words) == 1, extra) if w else w for i, w in enumerate(words))


def tidy(s, mode, extra=""):
    if isinstance(mode, str):
        mode = MODES.get(mode, 0)
    if mode <= 0 or not s:
        return s
    t = s.strip()
    if mode in (1, 2):
        t = drop_code(t)
    t = fix_ampersand(t)
    if mode >= 2:
        t = title_words(t, extra)
    return t


SAMPLES = ["1.07.00.MBGL.1360.LOANS TO CUSTOMERS", "4.01.00.MBGL.6510.CONTINGENT LIABILITIES& COMMITMENTS",
           "COA_MBGL.7210.BID BOND", "COA_MBGL.7245.LG against syndicated LC'S", "1.02.00.MBGL.0050.CASH",
           "1.07.11.CORPORATE LOANS - Direct Loans", "1.04.02.Advance Revenue for T.Bills & T.Bonds OCI ",
           "1.05.01.3180.AC INVESTMENTS", "1.04.04.FVTOCI INVESTMENTS - Shares", "3.01.00.MBGL.6440.TOTAL OWNERS EQUITY",
           "4.01.14.MBGL.9998.CATCH ALL OFF BALANCESHEET", "2.5% RESERVE", "T.Bills", "P&L ACCOUNT",
           "NON-PERFORMING LOANS", "DUE FROM BANKS", "ECL FOR AC INVESTMENTS", "A", "OF", "Off Balance Sheet",
           "(no line)", "NATIONAL BANK OF EGYPT", "US DOLLAR", "4.01.10.MBGL.7260.ISSUED LETTER OF GUARANTEES CONTRA",
           "3686_LTL UNSECURED LEASING", "16007_OTHER Credit Balances", "51000_INT Income",
           "COA_MBGL.3607.Accrule CORPORATE", "COA_MBGL.3625.Other Assets&AL POSITION", "3582_CBE LTL initive unsecured",
           "Hazem Muhammad Helmy Abdel AL HAQ AL HALWAGY", "BY", "MY ACCOUNT"]

if __name__ == "__main__":
    for s in SAMPLES:
        print("%-58s | %-45s | %s" % (s, tidy(s, 1), tidy(s, 2)))
