# Pivot config

Every pivot sheet a build writes is one row of **Pivot config**, or a family
of sheets, one per value of a field. The fields a row may name are listed on
**Pivot fields**. Change a row and the next build changes. Nothing about a
built workbook is fixed in code.

Charts, the number of workbooks and the gallery are on their own tabs; see
[REPORTS.md](REPORTS.md).

![Pivot config](../preview/sheet-pivot-config.png)

## A row

The columns are banded in four groups, plus the check. Select any cell and
Excel shows what goes in it. The columns with a fixed set of choices have a
dropdown.

**Which pivot**

| Column | What it takes | Example |
|---|---|---|
| **On** | Yes or No | Yes |
| **Pivot** | The sheet name. `{fw}` becomes the framework. With *One sheet per*, write `{split}`. | `{fw} Output` → *LCR Output* |
| **Frameworks** | All, or any of LCR, NSFR, Maturity ladder | `LCR, NSFR` |
| **One sheet per** | Optional. One or two text fields: one sheet per value, biggest first. | `Rule name, LCY / FCY` |

**What it shows**

| Column | What it takes | Example |
|---|---|---|
| **Rows** | Fields down the side, in order | `Type, Line, Subline, COA name` |
| **Columns** | Fields across the top | `Bucket` |
| **Report filters** | Fields that filter the whole pivot from a dropdown over it (see [Report filters](#report-filters)) | `Currency` |
| **Values** | One or more, separated by `;` (see [Values](#values)) | `Gross pre-factor sum as Exposure` |
| **Show only / hide** | Item and label rules, separated by `;` (see [Show only / hide](#show-only--hide)) | `Bucket <> (no bucket)` |
| **Top / value filter** | Top or bottom N, or a comparison, by a value (see [Top / value filter](#top--value-filter)) | `Top 25 by Exposure` |
| **Group** | A date by a period, or a number by a step (see [Group](#group)) | `Maturity date by year` |
| **Slicers** | Fields to put slicers on | `LCY / FCY` |

**How it looks**

| Column | What it takes | Default |
|---|---|---|
| **Layout** | Tabular, Outline or Compact | Tabular |
| **Subtotals** | None, All, or the row fields to subtotal | None |
| **Subtotals at** | Top or Bottom. This applies to Outline and Compact only; Tabular always totals below. | Excel's choice |
| **Grand totals** | Both, Bottom row, Right column, None | Both |
| **Total label** | What the grand totals are called | Total |
| **Repeat labels** | Yes repeats a row label on every line it covers | No |
| **Blank line** | Yes leaves a blank line after each group of the outer row fields | No |
| **Values in** | Columns, or Rows to stack two or more values under the row labels | Columns |
| **Expand to** | A row field to fold the table to. The fields below it start closed, with +/- to open them. | Everything open |
| **Sort** | `label asc` / `label desc`, or a value caption then `asc` / `desc` | The data's order |
| **Units** | As is, Thousands, Millions, Billions. The sheet says *IN MILLIONS* over its title. | As is |
| **Number format** | Any Excel format. It overrides Units. | `#,##0;[Red](#,##0);-` in the chosen unit |
| **Highlight** | None, Data bars, Heatmap, Negatives, Top N. Add `on` and a caption for one value only. | None |
| **Widths** | `field=width` and `values=width`, separated by `;` | Fitted to the longest label |
| **Tiles** | Yes shows each plain value's live total in a tile over the pivot | Yes |

**The sheet**

| Column | What it takes | Default |
|---|---|---|
| **Tab** | Auto, Emerald, Deep, Slate, Black | Auto |
| **Max sheets** | For *One sheet per*: the most sheets | 120 |
| **Description** | Shown under the title and in Start here's index | |

**Check** and **What to fix** are written by **Check**.

## Values

Each value is written as:

```
field  [aggregation]  [calculation]  [in field]  [as caption]
```

- **Aggregation**: `sum` (the default), `count`, `average`, `max`, `min`.
  - Text can only be counted.
  - A date can be counted, or its min or max taken.
  - A calculated field can only be summed.
- **Calculation**: how the value is shown.

  | Word | Shows | Runs along (unless `in` says) |
  |---|---|---|
  | `%row`, `%col`, `%total` | share of its row, column, grand total | |
  | `%parent` | share of its parent item | the first row field |
  | `%parentrow`, `%parentcol` | share of the parent row or column | |
  | `running` | running total | the last row field |
  | `%running` | running total as a share | the last row field |
  | `rank`, `rank asc`, `rank desc` | rank, largest first by default | the last row field |
  | `diff`, `%diff` | change from the previous item | the first column field |
  | `index` | Excel's index | |

- **in field**: what a running total, rank, share or change runs along. It
  must be on the rows or columns.
- **as caption**: the column's heading. If it is left out, the heading is
  *Sum of …*, *Running total of …* and so on. A caption that equals a field
  name is made safe automatically; Excel would otherwise refuse it silently.

The words are read from the end, and only while what is left is not already
a field. A field called *Price index* or *Days in arrears* reads as itself.

```
Pre factor amount sum as Pre-factor; Post factor amount sum as Post-factor
Gross pre-factor sum as Exposure; Gross pre-factor %total as Share
Gross pre-factor %running in Counterparty as Cumulative
Gross pre-factor rank as Rank
Haircut sum as Taken off; Effective factor sum as Applied factor
Account count as Accounts
```

## Show only / hide

| Rule | Does |
|---|---|
| `Field = a \| b` | shows only those items |
| `Field <> a \| b` | hides them |
| `Field contains text` | shows items whose label contains the text |
| `does not contain`, `begins with`, `does not begin with`, `ends with`, `does not end with` | as they say |

- Separate several rules with `;`.
- An item rule on a field that is not on the rows or columns becomes a report
  filter over the pivot, on a row of its own.
- A label rule needs the field on the rows or columns.
- Blank cells show as the field's *Blank shows as* label, for example
  `(no counterparty)`. That label is a real item, so it can be hidden
  reliably. The top counterparties default does exactly that.

```
Bucket <> (no bucket); Counterparty <> (no counterparty)
Counterparty contains BANK
Type = Asset | Liability
```

## Report filters

Each field listed becomes a dropdown over the pivot that filters the whole
table. Separate several with commas.

| Written | Starts on |
|---|---|
| `Currency` | every value |
| `Currency = USD` | USD, one value at a time |
| `Currency = All` | every value, always |

The dropdown offers every value; the start only decides what the sheet
opens on.

**A Currency filter and native amounts.** Some outputs carry amounts in each
row's own currency as well as in local currency. Staging uses those native
amounts when it finds them, and adding dollars to pounds means nothing. So
when a file has them, a plain `Currency` filter starts on the local
currency, and Start here says why. `Currency = All` overrides this. On a file
with local-currency amounts only, which is what the LCR sample has, every
currency is already in EGP and `Currency` starts on every value.

**Maturity ladder: currency as a filter instead of a workbook each.**

1. On **Workbooks**, clear *One workbook per* on the Maturity ladder row, and
   set *File name* to `{fw}`.
2. On **Pivot config**, on the ladder's `{split}` row (One sheet per = Rule
   name), write `Currency` under **Report filters**.

The ladder then builds as one workbook with one sheet per rule, and each
sheet has a Currency dropdown over it. To get the per-currency workbooks
back, undo both changes.

## Top / value filter

| Written | Keeps |
|---|---|
| `Top 25 by Exposure` | the 25 largest items of the first row field, by the value captioned *Exposure* |
| `Bottom 10 by Net` | the 10 smallest |
| `Top 5% by Exposure` | the items making up the top 5% |
| `Top 10 Counterparty by Exposure` | the top 10 of a named field. With Rows = Product, Counterparty, that is the top 10 within each product. |
| `Exposure > 1m` | items whose value is over a million; also `>=`, `<`, `<=`, `=`, `<>` |
| `Exposure between 1m and 5m` | also `not between` |
| `Counterparty: Exposure > 100m` | the same on a named field |

- Numbers may be written `1500000`, `1,500,000`, `1.5m`, `2bn`, `750k` or
  `5%`. They are compared with the full amounts, whatever **Units** shows.
- One value filter is allowed per field. It can sit alongside a show-only or
  hide rule on the same field.
- Excel's totals count only the items shown, so a share of the top 25 is a
  share of the 25.

## Group

| Written | Makes |
|---|---|
| `Maturity date by year` | 2025, 2026, … |
| `… by quarter` | 2026 Q1, 2026 Q2, … |
| `… by month` | Mar 2026, Apr 2026, … |
| `… by day` | the date without its time |
| `Interest rate by 0.5`, `Gross pre-factor by 1m` | the step each value falls in, named by where it starts |

The grouped field must be on the rows or columns. It is grouped while the
output is staged, into a column of its own called, for example, *Maturity
date by year*. Excel's own grouping is not used: it would group the cache
that every sheet shares, and it gives up on a column with one blank in it.
Everything that names the field (a running total `in Maturity date`, a
filter, a width) follows it to the grouped column.

## One sheet per

With `One sheet per = Rule name, LCY / FCY`, a build makes one sheet for each
rule-and-side pair found in the data.

- The pairs come biggest rule first, and a rule's LCY and FCY sheets sit side
  by side.
- A pair with no rows gets no sheet.
- Rows with no value, such as *(no rule)*, get no sheet of their own. They
  are still in the overview pivots.
- Slicers are left off these sheets, because one slicer would filter every
  sheet in the family at once.

## The defaults

These rows are switched on:

| Pivot | Frameworks | One sheet per | Rows | Columns | Values |
|---|---|---|---|---|---|
| `{fw} Output` | All | | Rule order, Rule category, Rule name, Factor | LCY / FCY | Pre-factor; Post-factor |
| Balance sheet | All | | Type, Line, Subline, COA name | LCY / FCY | Pre-factor, subtotalled by Type |
| `{split}` | LCR, NSFR | Rule name, LCY / FCY | Type, Line, Subline, COA name | Bucket | Pre-factor |
| `{split}` | Maturity ladder | Rule name | Type, Line, Subline, COA name | Bucket | Pre-factor |
| `{fw} top counterparties` | All | | Counterparty | | Exposure; Share (%total); Cumulative (%running), top 25, in millions, data bars, LCY / FCY slicer |

The ladder is built as one workbook per currency (see [Workbooks](REPORTS.md#workbooks)),
so inside each workbook its sheets are one per rule, as for LCR and NSFR.

Five more rows ship switched off, ready to turn on:

- counterparty heatmap
- product concentration
- maturity by year
- rule ranking
- one sheet per product

The [Gallery](REPORTS.md#gallery) adds more with one click each.

## Pivot fields

![Pivot fields](../preview/sheet-pivot-fields.png)

The first thirteen fields are built in. They cannot be renamed. Every other
row names a column of the output, or something worked out from the output:

| Column | What it holds |
|---|---|
| **Source column** | The header in the file; case and punctuation do not matter. `=\|PRE\|` and `=\|POST\|` are the amounts without their sign. For a Calculated field, a formula over other fields' names in single quotes. |
| **Kind** | Text, Number, Date or Calculated |
| **Labels** | As is, Drop codes, Drop codes + title case, Title case |
| **Blank shows as** | The label for empty cells |
| **Width** | The column width when the field is on the rows |
| **Number format** | For Number, Date and Calculated fields |

- **Words kept in capitals**, beside the table, lists acronyms that title
  case must leave alone, such as *QNB* or *CIB*.
- Two calculated fields ship. The pivot works them out from the sums of
  whatever a row adds up to:
  - *Haircut*: `='Pre factor amount' - 'Post factor amount'`
  - *Effective factor*: `=IF('Pre factor amount'=0,0,'Post factor amount'/'Pre factor amount')`
- Only the fields a recipe or chart actually uses are staged. A calculated
  field stages the fields its formula reads. If a file lacks a field's
  column, the field is staged blank and Start here says so.

## Adding, checking, undoing

- **Add a pivot** writes a working recipe, switched off, on the next free
  row.
- **As you type**, the status bar says whether the row will build and, if not,
  what to fix. Nothing is written while you edit, so **Undo** keeps working.
  The Check column catches up when you leave the sheet.
- **Check** writes *OK* or *Break* beside every row, with the reason. A build
  runs the same check first, and stops before it starts if a row that is on
  will not build.
- **A row Excel refuses** at build time is logged on Activity with its row
  number. The rest of the workbook is still built.
- **Use the 1.0 layout** builds what 1.0 built, whatever the sheet says;
  **Build from this sheet** switches back.
- **Restore defaults** puts Pivot config and Pivot fields back to the rows
  above.
