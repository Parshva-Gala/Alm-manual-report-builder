<#
    Builds JKB_Stress_Testing_Tool_V15.xlsm (the UI and UX redesign) from these sources.

    Requirements (the same as V13's build kit):
      - Windows with Microsoft Excel installed
      - PowerShell 5 or later
      - In Excel: File > Options > Trust Center > Trust Center Settings >
        Macro Settings > tick "Trust access to the VBA project object model"

    What it does:
      1. Rebuilds the report template from shared\report so generated review
         workbooks get the new palette.
      2. Imports VBA\ and shared\ into a copy of JKB_Stress_Testing_Tool_V14_input.xlsm,
         injects the new console, filter builder and progress window, compiles the
         whole project (stops loudly on any compile error), then runs
         JKB_ApplyReleaseSetup once. That writes the ECL and RWA group setup into the
         config and creates Config_ValueSources, which links each base and pre-shock
         value to the input files.
      3. Packages the result as JKB_Stress_Testing_Tool_V15.xlsm and reopens it with
         alerts on to prove Excel does not want to repair it.

    The input workbook is never changed.

    Usage, from the folder this file is in:
        powershell -ExecutionPolicy Bypass -File .\Rebuild.ps1
#>
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

$src  = Join-Path $here "JKB_Stress_Testing_Tool_V14_input.xlsm"
$dest = Join-Path $here "JKB_Stress_Testing_Tool_V15.xlsm"
$out  = Join-Path $here "_st_build_out.xlsm"
$tpl  = Join-Path $here "Manual_Report_Template.xltm"
$mod  = Join-Path $here "shared\report\modReport.bas"
$doc  = Join-Path $here "shared\report\ThisWorkbook.cls"

if (-not (Test-Path $src)) { throw "Missing input workbook: $src" }
if (-not (Test-Path $mod)) { throw "Missing report module: $mod" }
Get-ChildItem -Path $here -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

Write-Output "=== 1/3 rebuilding the report template ==="
& (Join-Path $here "build\template.ps1") -Module $mod -Output $tpl -DocModule $doc

Write-Output ""
Write-Output "=== 2/3 importing, compiling and applying the config setups ==="
& (Join-Path $here "build\build.ps1") -Workbook $src -Output $out -Template $tpl -Compile `
    -Run "JKB_ApplyReleaseSetup"

Write-Output ""
Write-Output "=== 3/3 packaging the new workbook ==="
& (Join-Path $here "build\package.ps1") -Source $out -Dest $dest

if (Test-Path $out) { Remove-Item $out -Force }
Write-Output ""
Write-Output ("Done. The new workbook is: {0}" -f $dest)
