# Validation record

Validated on 26 September 2026. This record applies to the delivered ESG Banking Product workbook, identified by the SHA-256 below.

## Results

219 checks passed across the module, integrated workbook and native Excel test suites. Some principles are deliberately exercised in more than one engine; the count is test executions, not a count of independent model risks covered.

| Test suite | Executed | Passed |
| --- | ---: | ---: |
| Integrated workbook | 56 | 56 |
| Financial engine | 38 | 38 |
| E&S and taxonomy | 42 | 42 |
| Emissions | 55 | 55 |
| Microsoft Excel | 28 | 28 |

Microsoft Excel 16.0 opened the workbook normally and performed full recalculation in a read-only test session. The saved workbook contains 28 worksheets, 21 filterable tables and 4 native charts. Both the exported formula-error scan and native Excel scan found **zero formula errors**. No external workbook formula links were found. No business calculation depends on the terminal Controls worksheet.

## What was tested

- Accounting scenario weights; Stage 1, 2 and 3 treatment; independent annual-bucket ECL checks; recovery-delay direction; lifetime effects beyond the five-year display; and separation of accounting ECL from conditional stress.
- Annual bank balance sheets in all four cases; facility roll-forwards; noncash write-offs; immediate default settlement; and exact capital/cash reconciliation when management actions are switched off.
- Default and SICR backstops, invalid overrides, duplicate identifiers, the next prepared facility row, later-period assumptions, missing selected versus unselected case inputs, and propagation of unavailable results to headline outputs.
- Country selection, internal versus Jordan taxonomy routing, safeguards, allocation ceilings, missing/expired evidence, enhanced diligence, independently verified closure, critical taxonomy actions and E&S summary counts.
- Class-specific financed-emissions methods, original asset values, reported versus activity data, unit/year/country checks, missing versus zero, separate borrower Scope 3, exposure coverage, operational Scope 2 and category exclusions.
- Native Excel dropdown source, country routing, renovation Amber classification, missing-Scope-2 reconciliation across combined emissions views, and incomplete-scenario headline gating.
- Rendered worksheet views, readable headings and units, navigation targets, cached chart data and saved-file structure.

## Boundaries

The checks verify implementation behavior for the tested conditions. They do not validate the synthetic data, establish predictive performance, certify regulatory compliance, or replace an independent bank model-validation process. The capital and liquidity implementation boundaries, simplified repayment/recovery conventions and national applicability qualifications are described in Product_Guide.md and the framework.

Test mutations were confined to temporary in-memory/read-only sessions and restored before the delivered export. The workbook opens with Egypt, Delayed transition and Demo Nile Cement selected.

SHA-256: `321dd78f9a2f5a78023d7852ceacf8e73b883a147be9cbea8cd994849b3c7653`
