<#
    Builds the macro-enabled template that generated manual report workbooks are
    born from.

    A report workbook is handed to a reviewer who will open it on their own
    machine, months later, with no copy of the tool. Its buttons therefore have
    to travel with it, and a workbook cannot be given a VBA project after the
    fact without the user turning on "Trust access to the VBA project object
    model" - which is not something to ask of a reviewer.

    So the project is put into a TEMPLATE here, at build time, where that trust
    already exists. Workbooks.Add(Template:=...) then produces workbooks that
    carry it, with no trust setting and no prompt anywhere near the reviewer.

    Usage:
        template.ps1 -Module <modReport.bas> -Output <Manual_Report_Template.xltm>
#>
param(
  [Parameter(Mandatory=$true)][string]$Module,
  [Parameter(Mandatory=$true)][string]$Output,
  [string]$DocModule
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Module)) { throw "Module not found: $Module" }
$name = [System.IO.Path]::GetFileNameWithoutExtension($Module)

# ThisWorkbook cannot be imported, only replaced line for line - same rule as the
# main build harness. It carries Workbook_Open, which is what lets a generated
# report freeze its own control band on a machine this script will never see.
if (-not $DocModule) {
  $cand = Join-Path (Split-Path -Parent $Module) "ThisWorkbook.cls"
  if (Test-Path $cand) { $DocModule = $cand }
}

function Strip-Attributes([string]$text) {
  $lines = $text -split "`r?`n"
  $out = New-Object System.Collections.ArrayList
  $inHeader = $true
  foreach ($l in $lines) {
    if ($inHeader -and ($l -match '^\s*Attribute\s+VB_' -or $l -match '^\s*(VERSION 1\.0 CLASS|BEGIN|END|\s+MultiUse)')) { continue }
    $inHeader = $false
    [void]$out.Add($l)
  }
  return ($out -join "`r`n")
}

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$owned = @($after | Where-Object { $before -notcontains $_ })
try { $xl.VBE.MainWindow.Visible = $false } catch {}

try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets($wb.Worksheets.Count).Delete() }

  # CRLF, or VBA imports it as the wrong kind of component. Same rule as build.ps1.
  $normDir = Join-Path ([System.IO.Path]::GetTempPath()) ("tplimport_" + $PID)
  if (-not (Test-Path $normDir)) { New-Item -ItemType Directory -Path $normDir | Out-Null }
  $normPath = Join-Path $normDir ([System.IO.Path]::GetFileName($Module))
  $raw = [System.IO.File]::ReadAllText($Module)
  $raw = $raw -replace "`r`n", "`n"
  $raw = $raw -replace "`n", "`r`n"
  [System.IO.File]::WriteAllText($normPath, $raw, (New-Object System.Text.UTF8Encoding($false)))

  $imported = $wb.VBProject.VBComponents.Import($normPath)
  if ($imported.Name -ne $name) { throw "imported as '$($imported.Name)' instead of '$name'" }
  Write-Output ("  imported {0} ({1} lines)" -f $imported.Name, $imported.CodeModule.CountOfLines)

  if ($DocModule -and (Test-Path $DocModule)) {
    $cm = $wb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString((Strip-Attributes ([System.IO.File]::ReadAllText($DocModule))))
    Write-Output ("  replaced ThisWorkbook ({0} lines)" -f $cm.CountOfLines)
  }

  # A blank template opens showing nothing useful, so say what it is.
  $sh = $wb.Worksheets(1)
  $sh.Name = "Sheet1"
  $sh.Cells(1,1).Value2 = "This is the template the ALM manual report workbooks are generated from. It carries their button strip and nothing else."

  # The whole project, compiled here rather than the first time a reviewer
  # presses a button on a machine that is not this one.
  try {
    $ctl = $xl.VBE.CommandBars.FindControl(1, 578)
    if ($null -ne $ctl -and $ctl.Enabled) {
      $ctl.Execute()
      Start-Sleep -Milliseconds 400
      if ($ctl.Enabled) { throw "the template's VBA project did not compile" }
    }
    Write-Output "  compiled"
  } catch { Write-Output "  (compile step: $($_.Exception.Message))" }

  if (Test-Path $Output) { Remove-Item $Output -Force }
  $wb.SaveAs($Output, 53)        # xlOpenXMLTemplateMacroEnabled
  $wb.Close($false)
  Write-Output ("Built {0} ({1:N0} bytes)" -f $Output, (Get-Item $Output).Length)
}
finally {
  try { $xl.Quit() } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  Start-Sleep -Milliseconds 700
  foreach ($procId in $owned) {
    if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }
  }
}
