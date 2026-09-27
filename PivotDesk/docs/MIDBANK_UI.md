# Midbank experience, version 3.2

The reference is Avati Risk Studio's Midbank presentation: a dark emerald
workspace, clear hierarchy, rounded surfaces, restrained borders and a useful
primary action. The Excel implementation keeps native cells, PivotTables,
filters and charts usable, and retains Segoe UI so no font installation is
needed.

## Coverage

| Surface | Shared experience |
|---|---|
| Desk | Operational headline, next action, live status, framework selection, metric cards, recent activity and builds |
| Files | Compact context header, workflow navigation, file actions and typed status table |
| Pivot config, Chart config, Workbooks, Pivot fields | Shared Reports tabs, context header, actions and configuration tables |
| Gallery | Reports navigation and consistent report/chart cards |
| Reconciliation, Activity | Shared header, status hierarchy and readable detail tables |
| Generated Start here | Midbank report header, summary tiles, charts and a clickable report index |
| Generated Output, Balance sheet, rule/currency/recipe pivots | Report header, linked totals, filter guidance, native pivot styling and workbook navigation |
| Generated standalone charts | Report header, framed chart, workbook navigation and bounded printing |

The private staging, chart-support and settings sheets remain internal. Excel's
own double-click drill-through creates a new native worksheet outside the
generator; macro-free generated workbooks cannot automatically intercept and
restyle a worksheet Excel creates later.

Tool print areas refresh before printing so new Activity and Reconciliation rows
are included. Generated workbooks include their visible report artwork and
slicers in the saved print area. If a later pivot refresh grows a macro-free
export beyond its original bounds, reset its print area in Excel.

## Preserved contracts

- Existing worksheet names and data addresses stay intact. Tool status is on
  row 5, column headings on row 7 and records begin on row 8.
- Main Desk action names, framework switches, keyboard navigation and tour
  targets remain connected to the existing macros.
- Pivot recipes, the Report filters column, currency handling, shared caches,
  calculations and the top-customer performance changes remain intact.
- Summary tiles retain their existing live formulas. Source selection changes
  mark retained reconciliation results as needing a rerun.
- Generated workbooks remain macro-free `.xlsx` files. The application remains
  `dist/Avati.xlsm`, built through the repository's VBA-preserving package builder.

## Verification

The repository build checks VBA structure and calls, static Excel compilation
rules, configuration columns, action targets, drawing geometry, palette,
contrast, ribbon contracts and workbook package integrity. Preview HTML mirrors
the runtime design and is reviewed in the browser, including synthetic generated
workbook examples.

The 3.2 build passed with **0 problems across 19 VBA modules**, including the
Desk shape/action contract, contrast and package checks. Seventeen browser
viewport captures cover the three Desk states, the working sheets and generated
report examples. Long or wide worksheets continue beyond the captured viewport;
the [report index preview](../preview/built-report-index.png) shows the lower
part of Start here separately. These examples use synthetic demonstration data.

Native Excel execution is a separate acceptance step. The browser previews and
static checks do not establish PivotTable, slicer, PivotChart or printing behavior
inside Excel. No macro security or device policy has been changed.

Open `dist/Avati.xlsm` in Excel, load representative inputs and generate an LCR,
NSFR and currency-split maturity workbook. Review Start here, an Output pivot,
a Balance sheet, a filtered currency report and an own-sheet chart. Check that
filter changes update totals, navigation reaches the named sheets and Print
Preview includes the report without helper columns or blank pages. On the Desk,
change a source selection after reconciling and confirm the rerun state.
