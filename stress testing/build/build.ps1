<#
    Applies the VBA sources in .\VBA to a workbook and saves the result.

    Standard modules (.bas) and real class modules are removed and re-imported.
    Document modules (ThisWorkbook, SheetN) cannot be imported, so their code
    module contents are replaced line-for-line instead.

    Usage:
        build.ps1 -Workbook <source .xlsm> -Output <dest .xlsm> [-Run <macro>] [-RunArgs <a1,a2>]
#>
param(
  [Parameter(Mandatory=$true)][string]$Workbook,
  [Parameter(Mandatory=$true)][string]$Output,
  [string]$SourceDir,
  [string]$SharedDir,
  [string]$Template,
  [string]$Run,
  [string[]]$RunArgs = @(),
  [int]$RunTimeoutSec = 900,
  [switch]$Compile,
  [switch]$KeepOpen
)
$ErrorActionPreference = "Stop"

if (-not $SourceDir) { $SourceDir = Join-Path (Split-Path -Parent $PSScriptRoot) "VBA" }
if (-not (Test-Path $SourceDir)) { throw "VBA source directory not found: $SourceDir" }

# Modules and pages both tools use, kept in one place so there is one copy to
# improve. The console and the condition builder are the two surfaces a person
# actually touches, and neither belongs to one tool.
if (-not $SharedDir) {
  $p = Split-Path -Parent $PSScriptRoot
  foreach ($cand in @((Join-Path $p "shared"), (Join-Path (Split-Path -Parent $p) "shared"))) {
    if (Test-Path $cand) { $SharedDir = $cand; break }
  }
}

# shared\report holds the module that goes into the TEMPLATE the generated report
# workbooks are born from - never into either tool. It lives under shared because
# both tools' templates are built from it; it is excluded here because a tool
# carrying RptZoomIn and a ThisWorkbook.Worksheets("Data") reference is carrying
# code that can only ever be dead.
function Get-VbaSources([string]$dir) {
  return @(Get-ChildItem -Path $dir -Include *.bas,*.cls -File -Recurse |
           Where-Object { $_.DirectoryName -notmatch '[\\/]report$' })
}

$DOC_MODULES = @('ThisWorkbook')   # plus anything matching Sheet\d+

function Is-DocModule([string]$name) {
  return ($DOC_MODULES -contains $name) -or ($name -match '^Sheet\d+$')
}

function Strip-Attributes([string]$text) {
  $lines = $text -split "`r?`n"
  $out = New-Object System.Collections.ArrayList
  $inHeader = $true
  foreach ($l in $lines) {
    if ($inHeader -and $l -match '^\s*Attribute\s+VB_') { continue }
    $inHeader = $false
    [void]$out.Add($l)
  }
  return ($out -join "`r`n")
}

# ---------------------------------------------------------------- lint ------
# VBA accepts module-level declarations only in the header block, before the first
# procedure. One placed between procedures compiles to "Variable not defined", which
# surfaces as a modal dialog inside an invisible Excel - a hang, not an error. Catch
# it here, before Excel is ever launched.
function Test-VbaModuleLayout([string]$path) {
  $problems = @()
  $depth = 0
  $seenProcedure = $false
  $n = 0
  foreach ($line in (Get-Content $path)) {
    $n++
    $t = $line.Trim()
    if ($t -match '^\s*(Public\s+|Private\s+|Friend\s+|Static\s+)*(Sub|Function|Property\s+(Get|Let|Set))\s+\w+') {
      $depth++; $seenProcedure = $true; continue
    }
    if ($t -match '^\s*End\s+(Sub|Function|Property)\s*$') { $depth--; continue }
    # A single-line If cannot carry ElseIf. VBA only reports it when the procedure is
    # first executed (compile on demand), so it can sit unnoticed for a long time.
    if ($t -match '^If\s+.+\s+Then\s+\S.*\bElseIf\b') {
      $problems += "  line ${n}: single-line If cannot have ElseIf - use a block If: $t"
    }
    # A reserved word used as a variable name is a compile error VBA reports only
    # as "Syntax error" on the Dim line, with no clue which identifier is at fault.
    # `Any` (the Declare parameter type) and `Name` (the Name statement) are the
    # two that read as ordinary English and get typed without a second thought.
    foreach ($rw in @('Any','Name','To','Then','Error','Print','Open','Close','Stop','Next','Loop','Step','Set','Let','Get','Option','Type','Resume','Declare')) {
      if ($t -match "(?i)\b(Dim|ByVal|ByRef|Static|Const)\s+(\w+\s+As\s+\w+\s*,\s*)*$rw\s+As\b") {
        $problems += "  line ${n}: '$rw' is a VBA reserved word and cannot be a variable name: $t"
      }
    }
    if ($depth -eq 0 -and $seenProcedure) {
      if ($t -match '^(Public|Private|Global|Dim)\s+(Const\s+)?\w+' -and $t -notmatch '^\s*''') {
        $problems += "  line ${n}: module-level declaration after a procedure - move it to the header: $t"
      }
      if ($t -match '^Option\s+') {
        $problems += "  line ${n}: Option statement after a procedure: $t"
      }
    }
  }
  if ($depth -ne 0) { $problems += "  unbalanced Sub/Function blocks (depth $depth at end of file)" }

  # Module-level state is named mXxx by convention throughout this project. Under
  # Option Explicit an undeclared one is a compile error, which Excel reports only as a
  # modal "Variable not defined" dialog with no location. Check it statically instead.
  $text = Get-Content $path -Raw
  $header = ($text -split '(?m)^\s*(Public|Private|Friend|Static)?\s*(Sub|Function|Property)\s+', 2)[0]
  $declared = @{}
  foreach ($m in [regex]::Matches($header, '(?m)^\s*(?:Public|Private|Global|Dim)\s+(?:Const\s+|WithEvents\s+)?(.+)$')) {
    $decl = $m.Groups[1].Value
    # `Private a As Long, b As Long` declares BOTH. Capturing only the first is how
    # this check came to report a declared variable as undeclared, which is worse
    # than not checking at all - it teaches you to ignore it.
    foreach ($d in [regex]::Matches($decl, '(?:^|,)\s*([A-Za-z_]\w*)\s*(?:\(\s*[^)]*\))?\s+As\s+')) {
      $declared[$d.Groups[1].Value] = $true
    }
    if ($decl -match '^\s*([A-Za-z_]\w*)') { $declared[$Matches[1]] = $true }
  }
  $used = @{}
  foreach ($m in [regex]::Matches($text, '\bm[A-Z]\w*')) { $used[$m.Value] = $true }
  foreach ($name in $used.Keys) {
    if (-not $declared.ContainsKey($name)) {
      $problems += "  module-level variable '$name' is used but never declared in the header block"
    }
  }
  return $problems
}

Write-Output "Linting VBA sources..."
$lintFailed = $false
$lintTargets = @(Get-VbaSources $SourceDir)
if ($SharedDir -and (Test-Path $SharedDir)) { $lintTargets += @(Get-VbaSources $SharedDir) }
foreach ($f in ($lintTargets | Sort-Object Name)) {
  # An empty source file is almost always a mistyped path from some other command,
  # and it would otherwise be imported as a silently empty module.
  if ($f.Length -eq 0) {
    $lintFailed = $true
    Write-Output "$($f.Name):"
    Write-Output "  the file is empty - delete it, or it will be imported as an empty module"
    continue
  }
  # VBA module names cap at 31 characters. A longer one imports as "Module1"
  # with no warning, so every reference to it then fails to compile.
  $modName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
  if ($modName.Length -gt 31) {
    $lintFailed = $true
    Write-Output "$($f.Name):"
    Write-Output "  module name is $($modName.Length) characters; VBA allows 31 and silently renames anything longer to Module1"
  }
  $problems = Test-VbaModuleLayout $f.FullName
  if ($problems.Count -gt 0) {
    $lintFailed = $true
    Write-Output "$($f.Name):"
    $problems | ForEach-Object { Write-Output $_ }
  }
}

# Standard modules reach the sheet code modules by name (Sheet1.Foo). VBA only rejects an
# unknown member when that procedure is first executed, so a call left behind after moving
# a routine can reach the user. Resolve every such reference against the real module.
$docMembers = @{}
foreach ($f in (Get-ChildItem -Path $SourceDir -Filter *.cls -File -Recurse)) {
  $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
  $members = @{}
  foreach ($m in [regex]::Matches((Get-Content $f.FullName -Raw), '(?m)^\s*(?:Public\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)')) {
    $members[$m.Groups[1].Value] = $true
  }
  $docMembers[$name] = $members
}
# Every procedure this project defines, and where it lives.
$projectProcs = @{}
foreach ($f in (Get-ChildItem -Path $SourceDir -Include *.bas,*.cls -File -Recurse)) {
  $owner = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
  foreach ($m in [regex]::Matches((Get-Content $f.FullName -Raw), '(?m)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)')) {
    $projectProcs[$m.Groups[1].Value] = $owner
  }
}
foreach ($f in (Get-ChildItem -Path $SourceDir -Include *.bas,*.cls -File -Recurse)) {
  $n = 0
  foreach ($line in (Get-Content $f.FullName)) {
    $n++
    if ($line -match "^\s*'") { continue }
    foreach ($m in [regex]::Matches($line, '\b(Sheet\d+|ThisWorkbook)\.(\w+)')) {
      $target = $m.Groups[1].Value; $member = $m.Groups[2].Value
      # Only our own procedures are checked. Excel's own members (FollowHyperlink,
      # Worksheets, Save, ...) are not ours to validate and are left alone.
      if ($docMembers.ContainsKey($target) -and $projectProcs.ContainsKey($member) -and -not $docMembers[$target].ContainsKey($member)) {
        $lintFailed = $true
        Write-Output "$($f.Name):"
        Write-Output "  line ${n}: '$target.$member' is qualified onto $target, but '$member' lives in $($projectProcs[$member])"
      }
    }
  }
}
if ($lintFailed) { throw "VBA problems found; Excel was not launched." }
Write-Output "  layout OK"

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
# The process this script owns, so it can be torn down unconditionally even when a
# modal VBA dialog has made Quit() impossible.
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$ownedPid = @($after | Where-Object { $before -notcontains $_ })
try { $xl.VBE.MainWindow.Visible = $false } catch {}

$wb = $null
try {
  $wb = $xl.Workbooks.Open($Workbook)
  try { $xl.Calculation = -4135 } catch {}   # xlCalculationManual

  $proj = $wb.VBProject
  $files = @(Get-VbaSources $SourceDir)
  if ($SharedDir -and (Test-Path $SharedDir)) {
    $files += @(Get-VbaSources $SharedDir)
  }
  $files = $files | Sort-Object Name

  foreach ($f in $files) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $text = Get-Content $f.FullName -Raw

    if (Is-DocModule $name) {
      $comp = $null
      try { $comp = $proj.VBComponents.Item($name) } catch {}
      if ($null -eq $comp) { Write-Output "  SKIP (no document module named $name)"; continue }
      $cm = $comp.CodeModule
      if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
      $cm.AddFromString((Strip-Attributes $text))
      Write-Output ("  replaced document module {0} ({1} lines)" -f $name, $cm.CountOfLines)
    }
    else {
      # A UTF-8 BOM hides the leading `Attribute VB_Name` line, and VBA then silently
      # names the component ModuleN. Fail loudly rather than shipping a renamed module.
      $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
      if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        throw "$($f.Name) starts with a UTF-8 BOM; save it without one or VBA will rename the module."
      }
      $comp = $null
      try { $comp = $proj.VBComponents.Item($name) } catch {}
      if ($null -ne $comp) { $proj.VBComponents.Remove($comp) }
      # VBA's Import only recognises the "VERSION 1.0 CLASS" header when the file
      # uses CRLF. With LF-only endings a class module is silently imported as a
      # STANDARD module, `New <Class>` then fails to compile, and the whole
      # project stops running with "the macro may not be available" - which names
      # neither the class nor the line. Normalise before importing.
      $normDir = Join-Path ([System.IO.Path]::GetTempPath()) ("vbaimport_" + $PID)
      if (-not (Test-Path $normDir)) { New-Item -ItemType Directory -Path $normDir | Out-Null }
      $normPath = Join-Path $normDir $f.Name
      $raw = [System.IO.File]::ReadAllText($f.FullName)
      $raw = $raw -replace "`r`n", "`n"
      $raw = $raw -replace "`n", "`r`n"
      [System.IO.File]::WriteAllText($normPath, $raw, (New-Object System.Text.UTF8Encoding($false)))
      $imported = $proj.VBComponents.Import($normPath)
      if ($imported.Name -ne $name) {
        throw "$($f.Name) imported as '$($imported.Name)' instead of '$name' - check its Attribute VB_Name line."
      }
      # 1 = standard module, 2 = class module. A .cls that lands as a standard
      # module is the silent failure described above, so fail loudly instead.
      if ($f.Extension -eq '.cls' -and $imported.Type -ne 2) {
        throw "$($f.Name) imported as a standard module, not a class. Its header must start with 'VERSION 1.0 CLASS'."
      }
      Write-Output ("  imported {0} ({1} lines, type {2})" -f $imported.Name, $imported.CodeModule.CountOfLines, $imported.Type)
    }
  }

  # The two pages a person actually touches - the condition builder and the console -
  # live as ordinary .hta files and are injected into very hidden sheets. That keeps
  # them editable as files and side-steps VBA's 1023-character line and 25-continuation
  # limits, which either comfortably exceeds.
  # A page only one tool has is simply not found for the other, and skipped.
  $pages = @(
    @{ file = "builder.hta";  sheet = "_Builder_Src" },
    @{ file = "console.hta";  sheet = "_Console_Src" },
    @{ file = "progress.hta"; sheet = "_Progress_Src" },
    @{ file = "loader.hta";   sheet = "_Loader_Src" }
  )
  foreach ($page in $pages) {
    $htaPath = Join-Path $SourceDir $page.file
    if (-not (Test-Path $htaPath) -and $SharedDir) { $htaPath = Join-Path $SharedDir $page.file }
    if (-not (Test-Path $htaPath)) { continue }
    $sheetName = $page.sheet
    $sh = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq $sheetName) { $sh = $w } }
    if ($null -eq $sh) {
      $sh = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets($wb.Worksheets.Count))
      $sh.Name = $sheetName
    }
    $sh.Visible = -1          # must be visible to write to it
    $sh.Cells.Clear()
    $htaLines = @(Get-Content $htaPath)
    # One cell at a time: a few hundred writes at build time, and it avoids the
    # multi-dimensional-array marshalling that PowerShell handles poorly here.
    for ($i = 0; $i -lt $htaLines.Count; $i++) {
      $sh.Cells($i + 1, 1).Value2 = $htaLines[$i]
    }
    $sh.Visible = 2           # xlSheetVeryHidden
    Write-Output ("  injected {0} ({1} lines)" -f $page.file, $htaLines.Count)
  }

  # The manual-report template travels INSIDE the workbook, the same way the three
  # pages above do, so that shipping a tool is shipping one file. It is binary, so
  # it goes in base64 - split across rows because a cell holds 32,767 characters.
  if (-not $Template) {
    $p = Split-Path -Parent $PSScriptRoot
    foreach ($cand in @((Join-Path $p "Manual_Report_Template.xltm"),
                        (Join-Path (Split-Path -Parent $p) "Manual_Report_Template.xltm"),
                        (Join-Path $p "ALM_V10\Manual_Report_Template.xltm"))) {
      if (Test-Path $cand) { $Template = $cand; break }
    }
  }
  if ($Template -and (Test-Path $Template)) {
    $b64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Template))
    $sh = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq "_Template_Src") { $sh = $w } }
    if ($null -eq $sh) {
      $sh = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets($wb.Worksheets.Count))
      $sh.Name = "_Template_Src"
    }
    $sh.Visible = -1
    $sh.Cells.Clear()
    # Text format before writing: a long run of base64 is still a string, but a
    # General cell is entitled to decide otherwise and there is nothing to gain
    # from finding out which runs it decides that about.
    $sh.Columns(1).NumberFormat = "@"
    $chunk = 20000
    $row = 1
    for ($i = 0; $i -lt $b64.Length; $i += $chunk) {
      $len = [Math]::Min($chunk, $b64.Length - $i)
      $sh.Cells($row, 1).Value2 = $b64.Substring($i, $len)
      $row++
    }
    $sh.Visible = 2
    Write-Output ("  embedded {0} ({1:N0} bytes, {2:N0} base64 chars in {3} row(s))" -f
                  (Split-Path -Leaf $Template), (Get-Item $Template).Length, $b64.Length, ($row - 1))
  }
  else {
    Write-Output "  (no manual-report template found to embed)"
  }

  if ($Run -or $Compile) {
    Write-Output "  starting dialog watchdog ..."
    # Excel is invisible here, so a VBA compile or run-time error opens a modal dialog
    # nobody can see and the COM call blocks forever. A watchdog in a separate process
    # reads that dialog and dismisses it, turning a silent hang into a real error.
    $watchdog = Start-Job -ArgumentList $PID, $RunTimeoutSec -ScriptBlock {
      param($parentPid, $timeoutSec)
      Add-Type @"
using System;using System.Text;using System.Runtime.InteropServices;
public class Watch{
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
 public delegate bool EnumProc(IntPtr h, IntPtr l);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
 [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
 [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, StringBuilder l);
 [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
}
"@
      $deadline = (Get-Date).AddSeconds($timeoutSec)
      while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 700
        $excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' }
        foreach ($p in $excel) {
          $found = [IntPtr]::Zero
          $cb = [Watch+EnumProc]{ param($h,$l)
            $q = 0; [Watch]::GetWindowThreadProcessId($h, [ref]$q) | Out-Null
            if ($q -eq $p.Id) {
              $cn = New-Object System.Text.StringBuilder 128
              [Watch]::GetClassName($h, $cn, 128) | Out-Null
              if ($cn.ToString() -eq '#32770') { $script:found = $h }
            }
            return $true }
          [Watch]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
          if ($found -ne [IntPtr]::Zero) {
            $text = ''
            $cb2 = [Watch+EnumProc]{ param($h,$l)
              $sb = New-Object System.Text.StringBuilder 2048
              [Watch]::SendMessage($h, 0x000D, [IntPtr]2048, $sb) | Out-Null
              if ($sb.Length -gt 0) { $script:text += $sb.ToString() + ' | ' }
              return $true }
            [Watch]::EnumChildWindows($found, $cb2, [IntPtr]::Zero) | Out-Null
            Write-Output "VBA DIALOG: $text"
            [Watch]::PostMessage($found, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null  # WM_CLOSE
            return
          }
        }
      }
    }
    try {
      if ($Compile) {
        # VBA compiles procedures on demand, so a syntax error in a procedure nothing has
        # called yet stays invisible until a user finally reaches it. Compile the whole
        # project here (VBE "Compile VBAProject", control id 578) so it fails at build time.
        Write-Output "  compiling VBA project ..."
        try {
          $ctl = $xl.VBE.CommandBars.FindControl(1, 578)
          if ($null -eq $ctl) { Write-Output "  (compile control unavailable - skipped)" }
          elseif (-not $ctl.Enabled) { Write-Output "  already compiled" }
          else {
            $ctl.Execute()
            # A clean compile DISABLES the Compile VBAProject control. If it is still
            # enabled the compile failed - and the watchdog will have quietly closed
            # the dialog that said so, which is how a real error hid behind the word
            # "compiled" until a run finally tripped over it.
            Start-Sleep -Milliseconds 400
            if ($ctl.Enabled) {
              $where = ''
              try {
                $pane = $xl.VBE.ActiveCodePane
                if ($null -ne $pane) {
                  $sl = 0; $sc = 0; $el = 0; $ec = 0
                  $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                  $where = "$($pane.CodeModule.Name) line $sl"
                  $src = $pane.CodeModule.Lines($sl, 1)
                  $where += ":  $($src.Trim())"
                }
              } catch {}
              $compileFailed = "COMPILE FAILED at $where"
              Write-Output "  !! $compileFailed"
            } else {
              Write-Output "  compiled"
            }
          }
        } catch { Write-Output "  (compile step failed: $($_.Exception.Message))" }
      }
      if ($Run) {
        Write-Output "  running $Run ..."
        switch ($RunArgs.Count) {
          0 { $xl.Run($Run) | Out-Null }
          1 { $xl.Run($Run, $RunArgs[0]) | Out-Null }
          2 { $xl.Run($Run, $RunArgs[0], $RunArgs[1]) | Out-Null }
          default { throw "Run supports at most 2 arguments" }
        }
      }
    }
    finally {
      $dialogText = @(Receive-Job $watchdog -ErrorAction SilentlyContinue)
      Stop-Job $watchdog -ErrorAction SilentlyContinue
      Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
      if ($dialogText) { $dialogText | ForEach-Object { Write-Output "  !! $_" } }
    }
    if ($compileFailed) { throw "$compileFailed - the workbook was not written." }
  }

  if (Test-Path $Output) { Remove-Item $Output -Force }
  $wb.SaveAs($Output, 52)
  Write-Output ("Built {0} ({1:N0} bytes)" -f $Output, (Get-Item $Output).Length)
}
finally {
  if (-not $KeepOpen) {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    # Quit() cannot succeed while a VBA dialog is up, and a surviving process leaves a
    # stray Excel/VBE window on the user's screen. Never leave one behind.
    Start-Sleep -Milliseconds 800
    foreach ($procId in $ownedPid) {
      $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
      if ($p) { Write-Output "  (force-closing leftover Excel process $procId)"; Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }
    }
  }
}
