"""
The Avati design tokens - one place for every colour, size and font.

MIDBANK's mark is jet black, white, and the emerald the word MID is set in.
The system keeps that discipline: black is the field, emerald is the ONLY
accent, and everything else is a tone of one or the other. Status colours
exist for status and are never decoration.

The VBA theme module mirrors the same values (modPD_Theme); the build checks
they agree, so the sheets and the Desk cannot drift apart.
"""

# --- the field --------------------------------------------------------------
INK = "000000"            # the logo's own black: app bar, table heads
CANVAS = "060B09"         # the Desk behind everything - black with a green cast
SURFACE = "0C1512"        # cards
SURFACE_HI = "0F1B16"     # card top edge of the gradient
SURFACE_2 = "111F19"      # raised inside a card (chips, tiles)
SURFACE_3 = "16271F"      # pressed / selected surface
HAIR = "1A2922"           # hairlines on dark
HAIR_2 = "243A31"         # stronger borders on dark
BAR_WELL = "0B1310"       # the well a group of pills sits in (nav, tabs)

# --- emerald, the one accent -----------------------------------------------
EM = {
    50: "E6F6EF", 100: "C2EBDA", 200: "8FDBBE", 300: "4FC79C", 400: "16B07F",
    500: "009060",   # the brand emerald, sampled off the mark
    600: "00794F", 700: "006141", 800: "004A32", 900: "003323", 950: "001D14",
}

# --- type on dark ------------------------------------------------------------
TX_1 = "F2F7F4"           # primary
TX_2 = "A9BDB3"           # secondary
TX_3 = "7E9388"           # tertiary / labels
TX_4 = "4A5E55"           # disabled, placeholders

# --- status (on dark) --------------------------------------------------------
OK = "2FC48D"
OK_BG = "0D2A20"
WARN = "F2B544"
WARN_BG = "2A2310"
BAD = "FF6B5E"
BAD_BG = "2E1614"
IDLE = "7E9388"
IDLE_BG = "121D19"

# --- the light canvas the tables sit on (Files / Reconciliation / Activity) --
PAPER = "FFFFFF"
MIST = "F4F7F5"
LINE = "E2E9E5"
MUTED = "5F7068"
BODY = "18241F"
OK_TX, OK_LT = "0B6B47", "DDF5EA"
WARN_TX, WARN_LT = "8A5A00", "FFF1CF"
BAD_TX, BAD_LT = "B42318", "FDE5E2"
IDLE_TX, IDLE_LT = "5E6F67", "EDF2EF"

# --- Avati ------------------------------------------------------------------
# The mark itself is a picture (build/brand.py); this is its middle blue, for
# the word when the picture cannot be placed.
AVATI_BLUE = "1C8CCB"

# --- type -------------------------------------------------------------------
# Segoe UI ships with every Windows since Vista; nothing here depends on a
# cloud font arriving.
F_BODY = "Segoe UI"
F_SEMI = "Segoe UI Semibold"
F_LIGHT = "Segoe UI Light"
F_SEMILIGHT = "Segoe UI Semilight"
F_MONO = "Consolas"

# --- the Desk canvas, in points ---------------------------------------------
W, H = 1120, 660
GRID_COL_PT = 15.0        # every Desk column is 20 px = 15 pt
GRID_ROW_PT = 15.0
MARGIN = 28
GUTTER = 16
