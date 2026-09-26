# Avati ALM Desk — look & feel plan

(PivotDesk until 3.0.)

One product, not two. The console window is gone; its best ideas (the three
action cards, readiness pills, framework chips, one obvious next action, the
dark emerald look) move into the workbook, and the workbook runs as an
application: when the Desk is in front, Excel's chrome steps aside.

Status: `[x]` done and checked here · `[~]` done, needs a look in Excel on
Windows to confirm (cannot be run off Windows) · `[ ]` open · `[-]` dropped,
with the reason.

---

## A. Design system — the foundation everything else draws from

- [x] A01 Colour tokens from the MIDBANK mark: black field, emerald as the only accent, white canvas for numbers
- [x] A02 Emerald tonal ramp 50–950 around the sampled #009060
- [x] A03 Neutral ramp on dark with a faint green cast (text 1–4, surfaces 0–3, hairlines)
- [x] A04 Semantic status palette — OK / Check / Break / Idle — in dark and light variants
- [x] A05 Contrast checker: every text/background pair the build uses, WCAG AA (4.5:1 body, 3:1 large); build fails below
- [x] A06 Tokens are the single source of truth; the VBA palette block is generated from them
- [x] A07 Type scale: Display 28 Light · Title 15 Semibold · Heading 10.5 Semibold · Body 9–10.5 · Caption 8 · Overline 7.5 caps +150 tracking · Mono 8
- [x] A08 Font stack that ships with Windows (Segoe UI family, Consolas) — no cloud-font dependency
- [x] A09 4-pt spacing scale (4 8 12 16 20 24 28 32)
- [x] A10 Radius scale (4 small · 7 button · 10 tile · 14 card · 16 hero · full pill)
- [x] A11 Elevation on dark: flat · card · raised tile · toast, each with a defined shadow
- [x] A12 Hairline rules: 0.75 pt, token colours, never pure grey
- [x] A13 Icon set, one hand: 24-grid, 1.75 stroke, round caps and joins
- [x] A14 Icons: upload, folder, files, pivot grid, balance (reconcile), pulse (activity), check, arrow, reset, expand, info, warning, error, clock, database, calendar, spark
- [x] A15 Icons shipped as SVG (crisp at any zoom) with PNG fallback for older Office
- [x] A16 Cairo motif: star-and-cross lattice (the khatam pattern of the city's mashrabiya screens), used only in the hero and empty states
- [x] A17 Logo mark: an eight-point khatam star on an emerald tile
- [x] A18 Hero art rendered at 2× for high-DPI laptops
- [x] A19 Subtle film grain on large dark fields so gradients do not band
- [x] A20 Design grid for the Desk: 1120 × 660 pt (630 in 2.0), 28 margin, 16 gutter, 3 × 344 cards
- [x] A21 Rule: dark chrome, light data — no balance is ever read off a dark field
- [x] A22 One number format everywhere (existing NUM_FMT) plus a compact form for KPI tiles (4.5 k, 30.1 bn)  
      _compact form used where amounts appear in sentences (the largest break)_
- [x] A23 Date conventions: "30 Nov 2025" for data, "26 Sep 14:05" for events
- [x] A24 Voice: plain, specific, action first; counts in words when small
- [~] A25 Motion within Excel's limits: press feedback, toast in/out, no gratuitous animation  
      _toast in/out, press feedback (buttons sink a point), busy overlay_

## B. Application shell — the workbook behaves like an app

- [x] B01 Remove the HTA console: module, page, button, temp-file exchange
- [x] B02 Repurpose the console's source sheet as a very-hidden `_Settings` store (keeps its code-name link)
- [~] B03 App mode on the Desk: ribbon, formula bar, headings, gridlines, sheet tabs hidden
- [~] B04 Excel's chrome restored whenever the user leaves the workbook (deactivate, close) — never leave their Excel altered
- [~] B05 "Excel view" toggle on the app bar for when they want the ribbon back
- [~] B06 Window caption "PivotDesk · MIDBANK Cairo"
- [~] B07 Zoom-to-fit the Desk canvas on open, on activate and on window resize
- [~] B08 Scroll area locked to the canvas
- [x] B09 Every cell beyond the canvas painted canvas colour — no white letterbox at any window shape
- [~] B10 Cursor parked on a hidden cell so no selection box shows on the Desk
- [~] B11 One navigation bar on every sheet: same items, same order, active state
- [~] B12 Keyboard shortcuts while the workbook is active: Ctrl+Shift+D/F/R/A for Desk/Files/Reconciliation/Activity; removed on leave
- [~] B13 Status bar progress for long runs: `PivotDesk ▰▰▰▱▱ 60%  LCR · staged 240,000 of 400,000 rows`
- [~] B14 Busy cursor during work
- [~] B15 Status bar and cursor always restored, including on failure paths
- [x] B16 Macros-off banner on the Desk, shown by the saved file and hidden once macros run
- [~] B17 Opens on the Desk, every time
- [~] B18 Desk protected against stray typing (UserInterfaceOnly so the code can still paint)
- [x] B19 Document properties: title, subject, keywords, category
- [x] B20 Workbook-level theme set to the brand, so built-in table and slicer styles pick up emerald

## C. Desk — app bar

- [x] C01 Black bar with an emerald hairline — the masthead the sheets always had, now the app's title bar
- [x] C02 Logo tile with the khatam star
- [x] C03 Wordmark "Pivot" + "Desk" in emerald, "MIDBANK · CAIRO" overline
- [x] C04 Segmented navigation — Desk · Files · Reconciliation · Activity — with an active pill
- [~] C05 Data as-of chip, from the loaded outputs
- [x] C06 Right-hand actions: Excel view, Reset
- [x] C07 Version stamp

## D. Desk — hero

- [x] D01 Hero panel art: deep gradient, aurora glow, lattice fading in from the right, grain, inner hairline
- [x] D02 Date overline
- [x] D03 Time-of-day greeting
- [~] D04 One sentence that says where the desk stands: what is loaded, what is ready, what is missing
- [~] D05 Next-best-action button whose label and action follow the state (add files → build → reconcile → open results)
- [x] D06 Secondary action beside it
- [x] D07 Four KPI tiles: Files, Frameworks, Controls, Last reconciliation
- [x] D08 Segment meters under the counts (5 / 3 / 2 segments)
- [~] D09 Reconciliation tile takes its verdict colour
- [x] D10 Compact figures in tiles

## E. Desk — card 1, Add files

- [x] E01 Step badge, title, two-line description, icon tile
- [~] E02 Five slot rows: status dot, label, file name, rows, as-of
- [~] E03 Click a slot row to choose the file for exactly that slot
- [x] E04 Empty slot copy: "Not added · click to choose"
- [~] E05 Missing file state: red dot, "moved or renamed"
- [~] E06 Forced-mismatch state: amber dot when a file was put in a slot its columns disagree with
- [x] E07 "Scan a folder" and "Pick files" actions
- [x] E08 Link through to the Files sheet
- [~] E09 Long file names shortened in the middle, extension kept
- [x] E10 "3 of 5" count in the header

## F. Desk — card 2, Build pivots

- [x] F01 Framework rows with switches: on, off, unavailable
- [~] F02 Per-framework detail: rows, as-of
- [~] F03 Default selection: every loaded framework
- [~] F04 Selection remembered in `_Settings`
- [x] F05 Button reads what it will do: "Build 2 workbooks"
- [x] F06 Disabled look and explanation when nothing can be built
- [~] F07 Last output folder remembered; the folder picker opens there
- [~] F08 Last build: when, how many, where; "Open folder"
- [-] F09 Open a built workbook straight from the Desk  
      _dropped: "Open folder" opens Explorer on the output folder with every workbook in it; a per-file list did not earn its space on the card_
- [~] F10 Completion reported on the Desk, not in a message box

## G. Desk — card 3, Reconcile

- [x] G01 Control-report tiles (3 by COA, 6 by account) with status and key counts
- [x] G02 Verdict matrix: controls × frameworks, each cell coloured and labelled
- [x] G03 Overall verdict pill
- [~] G04 Last run time
- [x] G05 "Reconcile now" and "Open results"
- [~] G06 Disabled state that says what to add
- [~] G07 Largest difference called out

## H. Desk — activity, footer, connective tissue

- [~] H01 Recent activity: three newest events with time, level pill, stage, message
- [x] H02 "All activity" link
- [x] H03 Empty-state copy
- [x] H04 Footer line with what the tool is for and the version
- [x] H05 Step connectors between the three cards
- [x] H06 Hairline dividers inside cards aligned to one rhythm

## I. Feedback and states

- [~] I01 Toast on the Desk (success / warning / error) instead of information message boxes
- [~] I02 Toast dismisses itself after a few seconds, or on click
- [~] I03 Message boxes kept only for confirmations and hard failures
- [~] I04 Every failure message names the next step
- [~] I05 Press feedback on Desk buttons
- [~] I06 Buttons that cannot act look it and say why
- [x] I07 Empty desk is a designed state, not a blank one

## K. The data sheets — Files, Reconciliation, Activity

- [~] K01 App bar row with the same navigation as the Desk
- [~] K02 Title band: title, one-line description, emerald rule
- [~] K03 Status banner: tinted field, accent bar, glyph, sentence
- [~] K04 Table header: black, emerald underline, caps 8 pt with tracking
- [~] K05 Row height 20 pt, vertical centre, 1-level indent
- [~] K06 Subtle banding
- [~] K07 Verdict and status cells as pills with a leading dot glyph
- [~] K08 Figures right-aligned in the one number format
- [~] K09 Freeze panes under the header
- [~] K10 AutoFilter on the header row
- [~] K11 Gridlines and headings off
- [-] K12 Files: file cell links to the file  
      _dropped: the source files are 60-80 MB, and a link a person can click by accident opens one_
- [~] K13 Files: contextual actions — Scan a folder, Pick files, Use a file for this row, Clear this row
- [~] K14 Reconciliation: difference column carries in-cell data bars
- [~] K15 Activity: timestamp in mono, level pill, newest first
- [~] K16 Print setup: landscape, fit to width, repeating header, footer with page and date
- [~] K17 Empty-state line on each sheet
- [~] K18 Tab colours from tokens

## L. The workbooks it builds

- [~] L01 Brand theme (colours and fonts) applied to every framework workbook
- [~] L02 Custom "PivotDesk" PivotTable style: black header, emerald accents, quiet banding, strong grand totals
- [~] L03 Slicers styled to match
- [~] L04 "Start here" guide redesigned: masthead, key facts as tiles, sheet index with links, notes
- [~] L05 Every pivot sheet: masthead band and a "← Start here" link
- [~] L06 Tab colours by kind: guide, overviews, rule sheets
- [~] L07 Gridlines off on pivot sheets
- [~] L08 Document properties on the output
- [~] L09 Print setup on pivot sheets
- [~] L10 Guide states staged rows, local currency, as-of, amount field

## M. Words

- [x] M01 One vocabulary: Desk, Add files, Build pivots, Reconcile, outputs, control reports
- [x] M02 No copy anywhere still mentions the console
- [x] M03 Labels action-first and short
- [x] M04 Card descriptions two lines at most
- [x] M05 Status sentences are sentences, with counts

## N. Accessibility

- [x] N01 AA contrast verified by the build
- [x] N02 Status never by colour alone: a word always travels with the dot
- [~] N03 Nothing below 7.5 pt on the Desk; zoom-to-fit clamped so text stays legible  
      _zoom-to-fit clamped at 60%; below that the window scrolls_
- [x] N04 Alt text on every Desk shape
- [x] N05 Shortcuts listed on the Desk footer
- [x] N06 Click targets at least 22 pt tall

## O. Robustness of the interface code

- [x] O01 Painting the Desk can never break an operation
- [x] O02 A missing shape is skipped, not an error
- [x] O03 The code never adds shapes to the Desk; it only paints the ones designed
- [~] O04 Rebuilding a data sheet is idempotent and keeps its data
- [x] O05 Reset leaves the design intact

## P. Engineering

- [x] P01 Reproducible build: seed + sources + design → dist/PivotDesk.xlsm, no Excel needed
- [x] P02 VBA project writer (source only, spec-compliant)
- [x] P03 VBA lint gate (blocks, Option Explicit, constants, ambiguous names)
- [x] P04 Shape-name contract: every name the code paints exists in the drawing
- [x] P05 Every OnAction target exists as a Public Sub
- [x] P06 Package validation: XML well-formed, content types, relationships
- [x] P07 Opens in LibreOffice with the VBA project intact
- [x] P08 Previews rendered: populated desk, empty desk
- [x] P09 README: build, design notes, what changed
- [x] P10 Contrast report committed

## R. Found on the way — fixed

- [x] R01 Log rows were inserted formatted like the black table header (Insert copies the row above); now formatted from below and dressed
- [x] R02 DressTable toggled the Files AutoFilter off on every other refresh; it now only ever turns it on
- [x] R03 The as-of date was written as a serial number (45991) — .Value2 hands dates over as numbers; now written as "30 Nov 2025"
- [x] R04 Replacing a file kept the previous file's row count, as-of and amount field; a new file now starts clean
- [x] R05 The workbook was saved in manual calculation, which switched every workbook opened after it to manual too
- [x] R06 The as-of date is read the moment a file is placed, not only after a build
- [x] R07 Opening the workbook and closing it straight away no longer asks to save
- [x] R08 A pending toast timer is cancelled on close, so Excel never reopens the workbook to run it
- [x] R09 Guide and small emerald text on white moved to the deeper emerald: #009060 is 4.1:1 on white, below AA for small text
- [x] R10 The tertiary grey on dark and the muted grey on white were both just under AA; retuned to pass everywhere they are used
- [x] R11 "Use a file for this row" was described on the Files sheet but had no button; it now has one, beside Clear this row
- [x] R12 Slicers were placed at a fixed 74 pt, which the taller 2.0 masthead would have overlapped; they now get a band of their own
- [x] R13 1.0 made an FCY sheet for every rule even when the rule had no FCY rows; PickOne then failed silently and the sheet showed every side unfiltered. Recipe families only build combinations that exist
- [x] R14 The field catalog was first laid out under the recipes and would have shared their 7-wide first column; caught in preview and moved to its own sheet
- [x] R15 A sheet name with an apostrophe ("Customer's deposits") broke its Start here link; now escaped
- [x] R16 A deleted Pivot config sheet would have made a build produce nothing; it now comes back as the defaults
- [x] R17 First open added and dressed sheets with events on, so each new sheet was "activated" half-built; events and repainting now pause until the desk is ready

## S. 2.1 — a few notches higher

- [x] S01 Desk canvas grows to 1120 × 660; the bottom row becomes Recent activity (four rows) beside a Maturity gap card (in 3.0: Recent builds)
- [-] S02 Maturity gap card: net pre-factor in each bucket of the last build, bars around a zero line, emerald above and grey below, "(no bucket)" dim at the end, more than twelve buckets folded into the last bar — dropped in 3.0: the maturity gap was removed at the bank's request
- [-] S03 Net, gross and weighted factor beside the bars; a chip switches between the frameworks built — dropped in 3.0: the maturity gap was removed at the bank's request
- [-] S04 Before any build: the ghost of a gap under an opaque pill that says what will be drawn (the zero line no longer runs through the words) — dropped in 3.0: the maturity gap was removed at the bank's request
- [-] S05 Staging measures the gap in its one pass: net, gross and maturity dates per bucket; nothing is read twice — dropped in 3.0: the maturity gap was removed at the bank's request
- [x] S06 Bucket labels read as tenors (TenorDays): up to / over / ranges / days, weeks, months, years / overnight / demand / non-maturity; a label it cannot read is placed by its rows' average maturity date. Run as VBA in LibreOffice on 37 labels, and agrees with build/tenor.py on every one
- [~] S07 Every pivot with Bucket on rows or columns orders the buckets by tenor, in the 1.0 layout and in recipes; a recipe's own sort still wins
- [~] S08 Guided tour: six steps, the desk veiled around one part at a time with a ring and a card, Back / Next / Skip and progress pips; starts by itself the first time the workbook opens; F1 or ? replays it
- [x] S09 The tour's placement is one piece of arithmetic in the design and in the VBA; the build checks the titles, the targets and the geometry constants agree
- [~] S10 Motion: framework switches slide over four frames, easing out; toasts fade in
- [~] S11 The status bar's progress says how long is left once it can tell
- [~] S12 Pivot config answers as you type, in the status bar, without writing a cell, so Excel's undo keeps working; the Check column catches up when you leave the sheet
- [~] S13 Add a pivot: a working recipe, switched off, on the next free row
- [~] S14 A PivotDesk tab on the ribbon for Excel view: go to any sheet, scan, pick, build, reconcile, app view, tour
- [~] S15 Start here opens on At a glance: rows staged, gross and net pre-factor, weighted factor, local currency, as-of date; a long figure steps down in size rather than spilling out
- [-] S16 Start here draws the maturity gap as a live PivotChart on the workbook's one cache: tenor order, LCY and FCY stacked, no field buttons, axis in bn or m, columns kept column-shaped when there are only one or two buckets — dropped in 3.0: the maturity gap was removed at the bank's request
- [x] S17 Chart and tile colours measured: text 4.5:1, graphics 3:1 (FCY moved from a pale mint that measured 1.5:1 to deep emerald at 10.4:1)
- [x] S18 The shipped workbook's VBA runs in LibreOffice (build/lo_run.py): the whole project compiles there, and its pure functions are executed against known answers
- [x] S19 The lint keeps out the three constructs LibreOffice cannot parse, each rewritten to a plain equivalent Excel reads the same
- [x] S20 Build gate: every ribbon button reaches PD_RibbonClick and every ribbon image has its relationship
- [x] S21 Fixed: framework rows and file slots read "412,806 rowsTxt"
- [x] S22 Fixed: the Start here preview listed FCY sheets the recipe engine does not build, under names longer than Excel allows
- [x] S23 Previews: desk-showcase, desk-empty, desk-tour, built-start-here (the LCR sample's real figures), built-start-here-ladder (illustrative figures; removed in 3.0 with the gap)

## T. 3.0 — Avati, one look everywhere, reports without limits

Asked for: the generated sheets in the Desk's theme, pivots that no longer
look messy, far more pivot customisation and advanced report forms, custom
charts, the Avati name and mark, the maturity gap removed, and the maturity
ladder as one workbook per currency with sheets per rule.

Brand and shell

- [x] T01 PivotDesk becomes Avati ALM Desk: the name, version 3.0, the ribbon tab, document properties, footers
- [x] T02 The Avati mark composed from the supplied artwork (build/brand.py), carried in the workbook, drawn on every app bar; the word AVATI in Avati blue if the image cannot be placed
- [~] T03 The mark decoded at run time from the workbook itself (MSXML), so the built books need no file beside them
- [x] T04 Maturity gap removed: the Desk card, the staging measures, the Start here chart, the tour step; Recent builds takes its place

One look

- [~] T05 Every tool sheet on the Desk's dark theme: app bar with the mark and navigation, title block, toolbar of tabs and actions, status line, dark banded tables with hairlines
- [~] T06 The Normal style is dark in the Desk and in every workbook it writes, so there is no white at the edges
- [~] T07 Reports is a tabbed section: Pivots, Charts, Workbooks, Fields, Gallery

Built workbooks

- [~] T08 Every sheet: the Avati bar, Start here / Previous / Next, the title block, live tiles of each value's total (GETPIVOTDATA, so a slicer moves them)
- [~] T09 Filters on rows of their own, clear of the table; drill buttons off; "Total"; a dark Avati pivot style and slicer style
- [x] T10 Labels tidied while staging: codes dropped, title case, acronyms kept; run as VBA in LibreOffice and agrees with build/labels.py on 35 labels in 3 modes
- [~] T11 Start here: six tiles, the charts grid, the index and the notes

Pivot config (31 columns)

- [~] T12 Values as % of parent / parent row / parent column, running total, % running total, rank (asc / desc), change, % change, index - along a field named with "in"
- [~] T13 Top / value filter: top / bottom N and N%, comparisons and between, on the first row field or a named one; numbers as 1.5m, 2bn, 750k, 5%
- [~] T14 Label rules in Show only / hide: contains, begins with, ends with and their opposites
- [~] T15 Group: dates by year / quarter / month / day, numbers by a step, staged into a column of their own (Excel's grouping would group the shared cache)
- [~] T16 Subtotals at, Total label, Blank line, Values in, Expand to, Units, Highlight (data bars / heatmap / negatives / top N, on one value or all), Tiles
- [~] T17 Gross pre-factor and gross post-factor fields; Calculated fields (Haircut, Effective factor) worked out by the pivot
- [x] T18 Every new column parsed and checked with a message that points at the cell; run through the same live check and Check as before
- [x] T19 Recipe reading moved to modPD_Recipe, so no module grows unwieldy
- [x] T20 Build gate: every heading a default recipe, default chart or gallery card names is a column of its sheet; column constants point at the right headings

Charts

- [~] T21 Chart config: 13 types, categories and series, values / filters / groups / top-N as on Pivot config, sort, placement, size, labels, legend, units (Auto), palettes
- [~] T22 Each chart a PivotChart of its own hidden pivot over the shared cache; drawn on the dark surface, hairline grid, no field buttons
- [~] T23 Start here grid of thirds, halves, two thirds and full widths; or a sheet of its own in the index
- [x] T24 Palettes measured: every colour 3:1 against the chart surface, checked by the build

Workbooks

- [~] T25 Workbooks sheet: one workbook per value of a field, per framework; file-name template, only these, max workbooks
- [~] T26 The Maturity ladder by default one workbook per currency, each with one sheet per rule
- [~] T27 The output read once and held open; each part stages only its own rows

Gallery

- [~] T28 18 designed cards (12 pivots, 6 charts): Add / Switch on / Open, repainted on the way in
- [x] T29 The top counterparties, heatmap, product concentration and large exposures cards hide the blank counterparty; the sample showed "(no counterparty)" holding 54% of the top 25

Accessibility and verification

- [x] T30 Contrast pairs brought up to the dark theme: 46 pairs, including data bars (one emerald, 00794F, that is 3.4:1 against the rows and carries white figures at 5:1), highlights, slicers, charts and the gallery
- [x] T31 LibreOffice runs the shipped VBA: all 19 modules compile; number parsing, unit formats, date periods and number steps run against known answers
- [x] T32 Previews of Chart config, Workbooks, the Gallery, Start here with its charts and the top counterparties pivot, from the LCR sample staged as the VBA stages it

## Q. Configurable pivots (separate workstream)

User guide: [PIVOT_CONFIG.md](PIVOT_CONFIG.md).

- [~] Q01 Pivot config sheet: one row per pivot, or one sheet per value of a field
- [~] Q02 Pivot fields sheet: a catalog of every output column under a plain name, 37 shipped, any can be added
- [~] Q03 Rows, columns and values from any field in the catalog
- [~] Q04 Values: sum, count, average, max, min, and %row / %col / %total, each with its own caption
- [~] Q05 Captions that collide with a field name are made safe automatically (Excel refuses them silently)
- [~] Q06 Show only / hide item rules, on whichever axis the field is on, or as a report filter
- [~] Q07 One sheet per one or two fields: biggest first, capped, blank values skipped
- [~] Q08 Layout: Tabular, Outline, Compact
- [~] Q09 Subtotals: none, all, or per row field
- [~] Q10 Grand totals: both, bottom row, right column, none
- [~] Q11 Repeat item labels
- [~] Q12 Sort by label or by any value, ascending or descending
- [~] Q13 Column widths per field, and one width for the figures
- [~] Q14 Number format per pivot, per field; percent values formatted as percent
- [~] Q15 Tab colour per pivot, or automatic by kind
- [~] Q16 Slicers per pivot (left off one-sheet-per families, with the reason)
- [~] Q17 Frameworks per pivot; {fw} in names
- [~] Q18 Staging carries only the extra fields a recipe uses, typed as text, number or date
- [~] Q19 Check: OK / Break beside every row with the reason; the build runs it first and stops
- [x] Q20 Cell-by-cell help: every column carries an input message explaining its syntax, lists carry dropdowns
- [x] Q21 Grouped column bands (Which pivot / What it shows / How it looks / The sheet / Check), frozen first two columns
- [x] Q22 Defaults reproduce 1.0 exactly; a switched-off worked example shows a recipe of your own
- [~] Q23 Engine switch: build from the sheet, or with the proven 1.0 layout
- [~] Q24 Restore defaults
- [~] Q25 A recipe Excel refuses is logged and skipped; the rest of the workbook is still built
- [~] Q26 Start here says what the workbook was built from, and which fields the file lacked
- [x] Q27 Desk: Pivot config in the navigation, a count and link on Build pivots, Ctrl+Shift+P
- [x] Q28 Desk: the next action points at Pivot config when a recipe that is on will not build
- [x] Q29 VBA lint now also enforces VBA's 24-continuation and 1023-character limits
- [x] Q30 Previews of both sheets, drawn from the defaults in the VBA source itself
- [x] Q31 Four ready-made recipes ship switched off: counterparty by product, maturity profile, top counterparties, one sheet per product
- [x] Q32 VBA lint checks every call's argument count against the procedure's signature
