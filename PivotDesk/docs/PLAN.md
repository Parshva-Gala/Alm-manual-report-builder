# PivotDesk 2.0 — look & feel plan

One product, not two. The console window is gone; its best ideas (the three
action cards, readiness pills, framework chips, one obvious next action, the
dark emerald look) move into the workbook, and the workbook runs as an
application: when the Desk is in front, Excel's chrome steps aside.

Status: `[x]` done and checked here · `[~]` done, needs a look in Excel on
Windows to confirm (cannot be run off Windows) · `[ ]` open · `[-]` dropped,
with the reason.

---

## A. Design system — the foundation everything else draws from

- [ ] A01 Colour tokens from the MIDBANK mark: black field, emerald as the only accent, white canvas for numbers
- [ ] A02 Emerald tonal ramp 50–950 around the sampled #009060
- [ ] A03 Neutral ramp on dark with a faint green cast (text 1–4, surfaces 0–3, hairlines)
- [ ] A04 Semantic status palette — OK / Check / Break / Idle — in dark and light variants
- [ ] A05 Contrast checker: every text/background pair the build uses, WCAG AA (4.5:1 body, 3:1 large); build fails below
- [ ] A06 Tokens are the single source of truth; the VBA palette block is generated from them
- [ ] A07 Type scale: Display 28 Light · Title 15 Semibold · Heading 10.5 Semibold · Body 9–10.5 · Caption 8 · Overline 7.5 caps +150 tracking · Mono 8
- [ ] A08 Font stack that ships with Windows (Segoe UI family, Consolas) — no cloud-font dependency
- [ ] A09 4-pt spacing scale (4 8 12 16 20 24 28 32)
- [ ] A10 Radius scale (4 small · 7 button · 10 tile · 14 card · 16 hero · full pill)
- [ ] A11 Elevation on dark: flat · card · raised tile · toast, each with a defined shadow
- [ ] A12 Hairline rules: 0.75 pt, token colours, never pure grey
- [ ] A13 Icon set, one hand: 24-grid, 1.75 stroke, round caps and joins
- [ ] A14 Icons: upload, folder, files, pivot grid, balance (reconcile), pulse (activity), check, arrow, reset, expand, info, warning, error, clock, database, calendar, spark
- [ ] A15 Icons shipped as SVG (crisp at any zoom) with PNG fallback for older Office
- [ ] A16 Cairo motif: star-and-cross lattice (the khatam pattern of the city's mashrabiya screens), used only in the hero and empty states
- [ ] A17 Logo mark: an eight-point khatam star on an emerald tile
- [ ] A18 Hero art rendered at 2× for high-DPI laptops
- [ ] A19 Subtle film grain on large dark fields so gradients do not band
- [ ] A20 Design grid for the Desk: 1120 × 630 pt, 28 margin, 16 gutter, 3 × 344 cards
- [ ] A21 Rule: dark chrome, light data — no balance is ever read off a dark field
- [ ] A22 One number format everywhere (existing NUM_FMT) plus a compact form for KPI tiles (4.5 k, 30.1 bn)
- [ ] A23 Date conventions: "30 Nov 2025" for data, "26 Sep 14:05" for events
- [ ] A24 Voice: plain, specific, action first; counts in words when small
- [ ] A25 Motion within Excel's limits: press feedback, toast in/out, no gratuitous animation

## B. Application shell — the workbook behaves like an app

- [ ] B01 Remove the HTA console: module, page, button, temp-file exchange
- [ ] B02 Repurpose the console's source sheet as a very-hidden `_Settings` store (keeps its code-name link)
- [ ] B03 App mode on the Desk: ribbon, formula bar, headings, gridlines, sheet tabs hidden
- [ ] B04 Excel's chrome restored whenever the user leaves the workbook (deactivate, close) — never leave their Excel altered
- [ ] B05 "Excel view" toggle on the app bar for when they want the ribbon back
- [ ] B06 Window caption "PivotDesk · MIDBANK Cairo"
- [ ] B07 Zoom-to-fit the Desk canvas on open, on activate and on window resize
- [ ] B08 Scroll area locked to the canvas
- [ ] B09 Every cell beyond the canvas painted canvas colour — no white letterbox at any window shape
- [ ] B10 Cursor parked on a hidden cell so no selection box shows on the Desk
- [ ] B11 One navigation bar on every sheet: same items, same order, active state
- [ ] B12 Keyboard shortcuts while the workbook is active: Ctrl+Shift+D/F/R/A for Desk/Files/Reconciliation/Activity; removed on leave
- [ ] B13 Status bar progress for long runs: `PivotDesk ▰▰▰▱▱ 60%  LCR · staged 240,000 of 400,000 rows`
- [ ] B14 Busy cursor during work
- [ ] B15 Status bar and cursor always restored, including on failure paths
- [ ] B16 Macros-off banner on the Desk, shown by the saved file and hidden once macros run
- [ ] B17 Opens on the Desk, every time
- [ ] B18 Desk protected against stray typing (UserInterfaceOnly so the code can still paint)
- [ ] B19 Document properties: title, subject, keywords, category
- [ ] B20 Workbook-level theme set to the brand, so built-in table and slicer styles pick up emerald

## C. Desk — app bar

- [ ] C01 Black bar with an emerald hairline — the masthead the sheets always had, now the app's title bar
- [ ] C02 Logo tile with the khatam star
- [ ] C03 Wordmark "Pivot" + "Desk" in emerald, "MIDBANK · CAIRO" overline
- [ ] C04 Segmented navigation — Desk · Files · Reconciliation · Activity — with an active pill
- [ ] C05 Data as-of chip, from the loaded outputs
- [ ] C06 Right-hand actions: Excel view, Reset
- [ ] C07 Version stamp

## D. Desk — hero

- [ ] D01 Hero panel art: deep gradient, aurora glow, lattice fading in from the right, grain, inner hairline
- [ ] D02 Date overline
- [ ] D03 Time-of-day greeting
- [ ] D04 One sentence that says where the desk stands: what is loaded, what is ready, what is missing
- [ ] D05 Next-best-action button whose label and action follow the state (add files → build → reconcile → open results)
- [ ] D06 Secondary action beside it
- [ ] D07 Four KPI tiles: Files, Frameworks, Controls, Last reconciliation
- [ ] D08 Segment meters under the counts (5 / 3 / 2 segments)
- [ ] D09 Reconciliation tile takes its verdict colour
- [ ] D10 Compact figures in tiles

## E. Desk — card 1, Add files

- [ ] E01 Step badge, title, two-line description, icon tile
- [ ] E02 Five slot rows: status dot, label, file name, rows, as-of
- [ ] E03 Click a slot row to choose the file for exactly that slot
- [ ] E04 Empty slot copy: "Not added · click to choose"
- [ ] E05 Missing file state: red dot, "moved or renamed"
- [ ] E06 Forced-mismatch state: amber dot when a file was put in a slot its columns disagree with
- [ ] E07 "Scan a folder" and "Pick files" actions
- [ ] E08 Link through to the Files sheet
- [ ] E09 Long file names shortened in the middle, extension kept
- [ ] E10 "3 of 5" count in the header

## F. Desk — card 2, Build pivots

- [ ] F01 Framework rows with switches: on, off, unavailable
- [ ] F02 Per-framework detail: rows, as-of
- [ ] F03 Default selection: every loaded framework
- [ ] F04 Selection remembered in `_Settings`
- [ ] F05 Button reads what it will do: "Build 2 workbooks"
- [ ] F06 Disabled look and explanation when nothing can be built
- [ ] F07 Last output folder remembered; the folder picker opens there
- [ ] F08 Last build: when, how many, where; "Open folder"
- [ ] F09 Open a built workbook straight from the Desk
- [ ] F10 Completion reported on the Desk, not in a message box

## G. Desk — card 3, Reconcile

- [ ] G01 Control-report tiles (3 by COA, 6 by account) with status and key counts
- [ ] G02 Verdict matrix: controls × frameworks, each cell coloured and labelled
- [ ] G03 Overall verdict pill
- [ ] G04 Last run time
- [ ] G05 "Reconcile now" and "Open results"
- [ ] G06 Disabled state that says what to add
- [ ] G07 Largest difference called out

## H. Desk — activity, footer, connective tissue

- [ ] H01 Recent activity: three newest events with time, level pill, stage, message
- [ ] H02 "All activity" link
- [ ] H03 Empty-state copy
- [ ] H04 Footer line with what the tool is for and the version
- [ ] H05 Step connectors between the three cards
- [ ] H06 Hairline dividers inside cards aligned to one rhythm

## I. Feedback and states

- [ ] I01 Toast on the Desk (success / warning / error) instead of information message boxes
- [ ] I02 Toast dismisses itself after a few seconds, or on click
- [ ] I03 Message boxes kept only for confirmations and hard failures
- [ ] I04 Every failure message names the next step
- [ ] I05 Press feedback on Desk buttons
- [ ] I06 Buttons that cannot act look it and say why
- [ ] I07 Empty desk is a designed state, not a blank one

## K. The data sheets — Files, Reconciliation, Activity

- [ ] K01 App bar row with the same navigation as the Desk
- [ ] K02 Title band: title, one-line description, emerald rule
- [ ] K03 Status banner: tinted field, accent bar, glyph, sentence
- [ ] K04 Table header: black, emerald underline, caps 8 pt with tracking
- [ ] K05 Row height 20 pt, vertical centre, 1-level indent
- [ ] K06 Subtle banding
- [ ] K07 Verdict and status cells as pills with a leading dot glyph
- [ ] K08 Figures right-aligned in the one number format
- [ ] K09 Freeze panes under the header
- [ ] K10 AutoFilter on the header row
- [ ] K11 Gridlines and headings off
- [ ] K12 Files: file cell links to the file
- [ ] K13 Files: contextual actions — Scan a folder, Pick files, Use a file for this row, Clear this row
- [ ] K14 Reconciliation: difference column carries in-cell data bars
- [ ] K15 Activity: timestamp in mono, level pill, newest first
- [ ] K16 Print setup: landscape, fit to width, repeating header, footer with page and date
- [ ] K17 Empty-state line on each sheet
- [ ] K18 Tab colours from tokens

## L. The workbooks it builds

- [ ] L01 Brand theme (colours and fonts) applied to every framework workbook
- [ ] L02 Custom "PivotDesk" PivotTable style: black header, emerald accents, quiet banding, strong grand totals
- [ ] L03 Slicers styled to match
- [ ] L04 "Start here" guide redesigned: masthead, key facts as tiles, sheet index with links, notes
- [ ] L05 Every pivot sheet: masthead band and a "← Start here" link
- [ ] L06 Tab colours by kind: guide, overviews, rule sheets
- [ ] L07 Gridlines off on pivot sheets
- [ ] L08 Document properties on the output
- [ ] L09 Print setup on pivot sheets
- [ ] L10 Guide states staged rows, local currency, as-of, amount field

## M. Words

- [ ] M01 One vocabulary: Desk, Add files, Build pivots, Reconcile, outputs, control reports
- [ ] M02 No copy anywhere still mentions the console
- [ ] M03 Labels action-first and short
- [ ] M04 Card descriptions two lines at most
- [ ] M05 Status sentences are sentences, with counts

## N. Accessibility

- [ ] N01 AA contrast verified by the build
- [ ] N02 Status never by colour alone: a word always travels with the dot
- [ ] N03 Nothing below 7.5 pt on the Desk; zoom-to-fit clamped so text stays legible
- [ ] N04 Alt text on every Desk shape
- [ ] N05 Shortcuts listed on the Desk footer
- [ ] N06 Click targets at least 22 pt tall

## O. Robustness of the interface code

- [ ] O01 Painting the Desk can never break an operation
- [ ] O02 A missing shape is skipped, not an error
- [ ] O03 The code never adds shapes to the Desk; it only paints the ones designed
- [ ] O04 Rebuilding a data sheet is idempotent and keeps its data
- [ ] O05 Reset leaves the design intact

## P. Engineering

- [ ] P01 Reproducible build: seed + sources + design → dist/PivotDesk.xlsm, no Excel needed
- [ ] P02 VBA project writer (source only, spec-compliant)
- [ ] P03 VBA lint gate (blocks, Option Explicit, constants, ambiguous names)
- [ ] P04 Shape-name contract: every name the code paints exists in the drawing
- [ ] P05 Every OnAction target exists as a Public Sub
- [ ] P06 Package validation: XML well-formed, content types, relationships
- [ ] P07 Opens in LibreOffice with the VBA project intact
- [ ] P08 Previews rendered: populated desk, empty desk
- [ ] P09 README: build, design notes, what changed
- [ ] P10 Contrast report committed

## Q. Configurable pivots (separate workstream)

Tracked in [PIVOT_CONFIG.md](PIVOT_CONFIG.md).
