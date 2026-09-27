# Midbank experience, version 3.3

The reference is Avati Risk Studio's Midbank presentation: a dark emerald
workspace, clear hierarchy, rounded surfaces, restrained borders and a useful
primary action. The Excel implementation keeps native cells, PivotTables,
filters and charts usable, and retains Segoe UI so no font installation is
needed.

## Coverage

| Surface | Shared experience |
|---|---|
| Desk | Operational headline, next action, live status, framework selection, metric cards, recent activity and builds |
| Files | Compact context header, one file-action toolbar, Fit columns and typed status table |
| Pivot config, Chart config, Workbooks, Pivot fields | Shared Reports tabs, focused toolbar, separate reset actions and editable configuration tables |
| Gallery | Reports navigation and consistent report/chart cards |
| Reconciliation, Activity | Shared header, status hierarchy and readable detail tables |
| Generated Start here | Midbank report header, summary tiles, charts and a clickable report index |
| Generated Output, Balance sheet, rule/currency/recipe pivots | Report header, linked totals, filter guidance, native pivot styling and workbook navigation |
| Generated standalone charts | Report header, framed chart, workbook navigation and bounded printing |

Working sheets use a 124-point header across rows 2–3, a single 44-point
toolbar and 34-point wrapped table headings. The header explains the work;
actions live in the toolbar. Native row and column handles remain visible for
resizing. **Fit columns** is available on relevant toolbars and the Avati ribbon.
Generated reports fit native displayed values automatically, including formats
that need more than the former 30-character column-width cap. Fitting leaves
values, totals, filters and the current selection/view unchanged.

Generated titles, descriptive text and metric labels stay inside their own
surfaces. Linked totals preserve their values. Chart layout reserves every
required row before drawing, respecting Excel's row-height ceiling. Repeating
guide placement clears only its prior chart/panel group; the report index and
unrelated content remain intact.

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

The 3.3 build passed with **19 release VBA modules, 239 Desk shapes and 0
problems**. All **seven focused layout tests** passed. They exercise the
production chart-layout procedures through a limited syntax adapter; they do
not simulate native Excel rendering.

Native Excel acceptance completed on **27 September 2026**, with **38
pivot/layout assertions** plus chart checks. Actual Balance sheet and Output
PivotTables were generated from synthetic staging data. Fresh reports fitted
without numeric overflow; deliberately narrowed fixtures then produced these
results:

| Native fixture | Cells displaying `####` before → after fit |
|---|---:|
| Large contra balances | 12 → 0 |
| Twelve decimal places | 36 → 0 |
| Long currency format, exceeding the old width cap | 19 → 0 |
| Long date and time labels | 15 → 0 |
| Active report filter | 6 → 0 |

Each fixture preserved every pivot value and total, native field captions,
column handles, selection, zoom and scroll. The selected report filter also
remained unchanged. Ten mixed-size native guide charts passed inset,
non-overlap and index-bound checks. Repeated placement retained exactly ten
charts and ten panels with stable bounds, preserving an unrelated shape and
the index. A 520-point standalone PivotChart reserved at least 536 points
across bounded rows.

Native review also caught Excel replacing a card's dark gradient endpoint with
white. Initializing gradient mode before assigning its colours fixes the pale
edges around charts and KPI tiles; the repeated native run checks the exact
dark endpoints as well as chart geometry.

To reproduce, run from the PivotDesk directory:

```
python build/build.py
python build/chart_layout_check.py
python build/layout_acceptance.py --output qa/Avati-Layout-QA.xlsm
```

Choose a new filename containing `QA`. Open the disposable `.xlsm` in Excel,
enable its macros and run **PD_LayoutAcceptance**. The harness writes
`layout-acceptance.log` and a synthetic `.xlsx` result beside the QA copy. Its
two optional acceptance modules are excluded from the release workbook.
Add `--auto-run` when packaging to schedule the test after the QA copy opens.

Preview HTML and images mirror the design using synthetic examples; they are
not native acceptance evidence. The native run verifies the scenarios above,
not full acceptance of bank extracts, every recipe, slicer interaction or
printer configuration. Operational review still needs representative LCR,
NSFR and currency-split maturity data, control reconciliation and Print Preview.
