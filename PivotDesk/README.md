# Avati ALM Desk 3.1

A desk for daily ALM analysis at MIDBANK Cairo. Point it at the LCR, NSFR and
maturity-ladder outputs and control reports 3 and 6. It recognises each file
by its columns, builds workbooks of live PivotTables and PivotCharts for each
framework, and reconciles the outputs against the control reports.

It runs as one application inside Excel. While the workbook is in front, the
ribbon, formula bar and sheet tabs step aside and the **Desk** fills the
window. When another workbook comes to the front, Excel's chrome comes back.

![The Desk](preview/desk-showcase.png)

## Using it

1. Open `dist/Avati.xlsm` and choose **Enable Content**.
2. On the Desk, **Scan a folder** or **Pick files**.
3. Switch frameworks on or off, then **Build**. Each workbook opens on a
   *Start here* sheet with its figures, its charts and an index of its sheets.
4. **Reconcile** once a control report is on the desk.

What a build makes is set under **Reports**, a tabbed section of five sheets:

| Tab | What it decides |
|---|---|
| **Pivots** | Every pivot sheet: one row per pivot, 31 columns of options ([reference](docs/PIVOT_CONFIG.md)) |
| **Charts** | Every chart: on Start here or on a sheet of its own ([reference](docs/REPORTS.md#chart-config)) |
| **Workbooks** | How many files a framework becomes. The Maturity ladder is one workbook per currency ([reference](docs/REPORTS.md#workbooks)) |
| **Fields** | The output columns the pivots and charts may name, and how their labels read |
| **Gallery** | 18 ready-made reports and charts, each one click from the next build ([reference](docs/REPORTS.md#gallery)) |

The first time the workbook opens, a six-step tour walks through the Desk.
**F1** or the **?** on the app bar plays it again.

| Shortcut | Goes to |
|---|---|
| F1 | The tour |
| Ctrl+Shift+D | Desk |
| Ctrl+Shift+F | Files |
| Ctrl+Shift+P | Pivot config |
| Ctrl+Shift+R | Reconciliation |
| Ctrl+Shift+A | Activity |

**Excel view** on the app bar brings the ribbon back, with an **Avati** tab
first on it. **App view** hides it again.

## What changed in 3.1

The working sheets now carry the Desk's full layout. Files, all five Reports
pages, Reconciliation and Activity have the same inset emerald hero, star
lattice, rounded workflow cards, prominent next action and elevated toolbar.
The numbered workflow cards navigate between Files, Reports and Reconciliation;
the highlighted card identifies the current section, not completion status.
Excel/App view and Desk help are available from every page. Gallery cards use
the same gradient surface as the Desk.

The update preserves the existing data layout: status remains on row 5,
headings on row 7 and records from row 8. Existing recipes, validation,
reconciliation logic, filters and workbook outputs are unchanged. Version 3.1
triggers the existing one-time restyling on open.

![Files with the shared Desk layout](preview/sheet-files.png)

The preview images are generated layout previews, not screenshots from Excel.
The build checks VBA source, shape actions, package integrity, palette and
contrast. Native Excel interaction remains a separate acceptance check.

## What changed in 3.0

**Avati.** The product is Avati ALM Desk. The Avati mark leads every app bar,
on the Desk, on every tool sheet and on every sheet of every workbook a build
writes.

**One look, everywhere.** The Desk's dark emerald theme is now on every
sheet:

- Files, Reconciliation, Activity and all the Reports sheets share the app bar,
  a title block, a toolbar row and a status line. Tables sit on dark banded
  rows with hairlines, and verdicts show as pills.
- The workbook's Normal style is dark, so there are no white cells at the
  edges.

**Built workbooks that read cleanly.**

- Every pivot sheet has the Avati bar, with Start here, Previous and Next.
- Under the title, tiles show each value's live total. They are linked to the
  pivot, so a slicer or a filter moves them.
- Filters sit on rows of their own, clear of the table.
- Pivots use a dark "Avati" style. The +/- buttons are off, and the total row
  reads *Total*.
- Ledger codes are dropped from labels, and names are title-cased where
  Pivot fields says so. *1.07.00.MBGL.1360.LOANS TO CUSTOMERS* reads
  *Loans to Customers*.

![A built pivot: the top counterparties](preview/built-top-counterparties.png)

**Pivot config: far more customisation.** A row now has 31 columns.

- Values can be shown as:
  - % of row, column, total or parent
  - a running total, or % running total
  - a rank
  - a change or % change from the previous item
  - an index

  Each can run along a field you name, for example
  `Gross pre-factor %running in Counterparty as Cumulative`.
- **Top / value filter**: `Top 25 by Exposure`, `Top 10 Counterparty by Exposure`,
  `Top 5% by Share`, `Exposure > 1m`, `Exposure between 1m and 5m`.
- **Show only / hide** also takes label rules: *contains*, *begins with*,
  *ends with* and their opposites.
- **Group**: dates by year, quarter, month or day, and numbers by a step. The
  grouping is done while staging, so it never breaks the shared cache.
- Presentation options:
  - **Subtotals at** top or bottom
  - **Total label**
  - **Blank line** after groups
  - **Values in** rows or columns
  - **Expand to**: open folded to a level
  - **Units**: thousands, millions or billions
  - **Highlight**: data bars, heatmap, negatives, top N
  - **Tiles** on or off
- New fields:
  - *Gross pre-factor* and *Gross post-factor*: the amounts without their sign.
  - *Calculated* fields worked out by the pivot, such as *Haircut* and
    *Effective factor*.

**Custom charts.** Chart config draws PivotCharts:

- 13 types, from column and bar to doughnut and column + line.
- Categories and an optional series field.
- The same values, filters, groups and top-N as a pivot.
- Placement on Start here in a grid of thirds, halves and full widths, or on
  a sheet of its own.
- Options for labels, legend, units (Auto picks bn, m or k) and three
  palettes.

Each chart is a live PivotChart over the workbook's one cache, so a refresh
redraws it.

![Start here, LCR, with the default charts](preview/built-start-here.png)

**One workbook per currency for the Maturity ladder.**

- Workbooks splits a framework into a file per value of a field. The ladder's
  default is Currency.
- The output is read once and held open. Each workbook stages only its own
  rows, so every total, tile and chart in it belongs to that currency.
- Inside each ladder workbook, the sheets are one per rule, as for LCR and NSFR.
- Prefer one workbook, with the currency picked on each sheet? Clear *One
  workbook per* on the ladder's Workbooks row, and put `Currency` under the new
  **Report filters** column on its Pivot config row
  ([how](docs/PIVOT_CONFIG.md#report-filters)).

**The Gallery.** 18 designed cards:

- top counterparties, counterparty heatmap, product concentration, maturity
  by year, rule ranking, sector league
- currency by bucket, haircut by rule, local/foreign mix, bank counterparties,
  large exposures, balance summary
- six charts

*Add* writes the row switched on. A card already on a sheet offers *Switch on*
or *Open*.

![The Gallery](preview/sheet-gallery.png)

**Removed.** The maturity-gap card, its chart and the gap figures are gone,
as asked. The Desk's bottom right now lists **Recent builds**, each with an
*Open* link.

## Design

- **Colour.** The palette is MIDBANK's black, white and emerald (#009060),
  with its tones, the Avati blue of the mark, and separate status colours.
- **Chart colours.** Every chart palette colour clears 3:1 against the chart
  surface. The build checks this.
- **Data bars** use one emerald (#00794F). It stands out from the rows at
  3.4:1, and white figures on it read at 5:1.
- **Contrast.** Every text and background pair on the Desk and the sheets is
  measured for WCAG AA ([docs/CONTRAST.md](docs/CONTRAST.md)).
- **Type.** Segoe UI, which ships with Windows.

## Building

The workbook is built from source without Excel:

```
python3 build/build.py
```

| Input | Path |
|---|---|
| Seed package | `src/seed/PivotDesk_v1.xlsm` |
| VBA | `src/vba/*.bas`, `*.cls` (19 modules) |
| Desk design | `build/desk.py` (one spec, rendered to DrawingML for Excel and to HTML for previews) |
| Brand | `build/brand.py` → `design/brand/` (the Avati mark, composed from the supplied artwork) |

The build fails on any of these:

- **VBA lint.** Blocks must close, and every identifier must be declared.
  Every Excel constant must be known. No Public name may be declared twice.
  Every call must match its procedure's argument count. The three constructs
  LibreOffice cannot parse are kept out.
- **Excel compile checks** (`build/vbacompile.py`). LibreOffice compiles
  things Excel refuses, and Excel only reports them when a module is first
  used, often mid-build. The check covers:
  - a public name declared in two modules, of any kind;
  - a private member used through its module's name;
  - `ByRef argument type mismatch`;
  - `For Each` over a typed variable;
  - duplicate declarations in one procedure;
  - a Sub used as a value;
  - parentheses around a statement call's arguments;
  - `Exit` of the wrong kind, and `Next` closing the wrong `For`;
  - `GoTo` to a missing label;
  - unknown named arguments.
- **Config agreement.** Every heading named by a default pivot, a default
  chart or a gallery template must be a column of its sheet. The column
  constants must point at the right headings.
- **Shape-name contract, geometry, ribbon, palette.** The VBA and the design
  must agree.
- **WCAG contrast.** This covers the Desk, every sheet pair and every chart
  colour.
- **Package validity.**

Previews are rendered with the Chromium on the build machine:

```
python3 build/preview.py                               # the Desk
LCR_SAMPLE=<lcr sample.xlsx> python3 build/preview_sheets.py   # Files, Recon, Activity, a Balance sheet pivot
CONFIG_ONLY=1 python3 build/preview_sheets.py          # Pivot config, Pivot fields
LCR_SAMPLE=<lcr sample.xlsx> python3 build/preview_v3.py       # Chart config, Workbooks, Gallery, Start here, top counterparties
```

The built-workbook previews use the LCR sample's real figures, staged the way
the VBA stages them. The Reports previews read their rows from the VBA
source.

## Verification, and what still needs Windows

Checked here:

- The build passes every gate above.
- `build/lo_run.py` opens the shipped workbook in LibreOffice's
  VBA-compatible Basic.
  - All 19 modules compile there.
  - The tenor parser and label tidying agree with their Python mirrors: 37
    and 35 labels.
  - These run against known answers: number parsing (`1.5m`, `2bn`, `5%`),
    unit formats, date periods and number steps.
- The VBA project round-trips byte for byte.

```
python3 build/lo_run.py dist/Avati.xlsm
```

**Not checked here: anything that needs Excel itself.** LibreOffice has no
`Scripting.Dictionary`, which recipe parsing is built on, and no Excel
object model. That leaves out:

- pivots, slicers and PivotCharts;
- value filters, calculated fields, conditional formats;
- the logo, which is decoded with MSXML.

Before anyone relies on 3.0, open it once on Windows:

1. Enable Content. The tour should start, and the Desk should carry the Avati
   mark.
2. Load the LCR, NSFR and ladder outputs and build all three. You should get
   LCR, NSFR and one Maturity Ladder workbook per currency.
3. In LCR:
   - Start here shows three charts.
   - *LCR top counterparties* has data bars, Share and Cumulative.
   - Buckets run shortest first.
4. Gallery: add *Haircut by rule* and *Large exposures*, then build LCR again.
   Check the calculated fields and the folded outline.
5. Reports:
   - On Pivot config, type a bad value. The status bar should say why, and
     Ctrl+Z should still undo it.
   - On Chart config, **Check** should mark every row OK.
6. Choose Excel view. The Avati tab should lead the ribbon.

Anything Excel refuses is logged on Activity with the Pivot config or Chart
config row it came from. The rest of the workbook is still built.
