# Stress Testing Tool (V15)

An Excel VBA tool that reconciles a bank's stress testing system. It does not
produce the stress test. It reads the same input files the system uses (ECL,
CapRWA, LCR, liquidity ladder, NSFR, capital components), rebuilds each figure
along the chain

scenario → test case → base → pre-shock (filtered by test case) → shock → post-shock

and reconciles every figure against the system's output, with tolerances, a
results table and a review pack.

## Using it

The workbook opens on **Home**, which has four steps:

1. **Load**: load the input files.
2. **Check setup**: test cases and value sources.
3. **Run**: rebuild base and pre-shock and reconcile every check.
4. **Review and hand over**: the results table and the review pack, which
   also saves `Reconciliation_evidence.xlsx`.

Five sheets are visible: Home, Inputs, Test cases, Value sources and Results.
Configuration (metrics, fields, grouping rules) sits behind Configuration.

## What is in this folder

| Path | What it is |
|---|---|
| `VBA/` | Every standard module and class, plus `ThisWorkbook`, `Sheet1` and `Sheet2` |
| `shared/` | The review console, filter builder and progress windows (HTA), their bridge modules, and the report template source |
| `build/` | `build.ps1` (import, lint, compile), `package.ps1` and `template.ps1` |
| `Rebuild.ps1` | One-command build of `JKB_Stress_Testing_Tool_V15.xlsm` |
| `Manual_Report_Template.xltm` | The empty report template the build regenerates |

## What is not in this folder

This repository is public, so nothing that carries the bank's data is here:

- the build input workbook (`JKB_Stress_Testing_Tool_V14_input.xlsm`) and any
  built workbook, because their element sheets and output sheet hold the
  bank's figures;
- `shock_catalog.tsv`, which is taken from the bank's design document;
- the sample input files and the design documents.

Keep those on your own machine and copy them next to `Rebuild.ps1` before you
build. `.gitignore` in this folder blocks workbooks, CSV files and the shock
catalog so they are not committed by accident.

## Build (Windows with Excel)

1. In Excel, go to File > Options > Trust Center > Trust Center Settings >
   Macro Settings, and tick "Trust access to the VBA project object model".
2. Put `JKB_Stress_Testing_Tool_V14_input.xlsm` and `shock_catalog.tsv` in this
   folder.
3. Run `powershell -ExecutionPolicy Bypass -File .\Rebuild.ps1` from this folder.
4. Open `JKB_Stress_Testing_Tool_V15.xlsm`.

The build lints the VBA before it starts Excel and stops with the module and
line on any compile error. The input workbook is only read.
