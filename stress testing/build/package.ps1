<#
    Produces the shipping workbook and proves it opens without Excel repairing it.

    Repair happens when a package violates the OOXML schema - which is exactly how the
    V5 file was damaged. The final file is written by Excel's own SaveAs, so it is
    canonical by construction; this script then re-opens it with alerts ENABLED and a
    dialog watchdog watching, so a repair prompt would be caught rather than silently
    accepted, and runs a structural check over the package parts.
#>
param(
  [Parameter(Mandatory=$true)][string]$Source,
  [Parameter(Mandatory=$true)][string]$Dest
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

Copy-Item $Source $Dest -Force
Unblock-File $Dest -ErrorAction SilentlyContinue

# --- rebind + final save under the shipping file name -----------------------
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$owned = @($after | Where-Object { $before -notcontains $_ })
try {
  $wb = $xl.Workbooks.Open($Dest)
  # Button OnAction strings embed the workbook name, so rebind under the final name.
  $xl.Run("InitializeJKBTool")
  $wb.Save()
  $wb.Close($false)
} finally {
  try { $xl.Quit() } catch {}
  Start-Sleep -Milliseconds 800
  foreach ($procId in $owned) { if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue } }
}
Write-Output ("Packaged {0} ({1:N0} bytes)" -f $Dest, (Get-Item $Dest).Length)

# --- structural check: the things that actually trigger a repair -------------
$ORDER = @(
  'sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr',
  'sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate',
  'customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations',
  'hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks',
  'customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing',
  'legacyDrawingHF','drawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst')
$rank = @{}; for ($i=0; $i -lt $ORDER.Count; $i++) { $rank[$ORDER[$i]] = $i }

$z = [System.IO.Compression.ZipFile]::OpenRead($Dest)
$issues = @()
$names = @($z.Entries | ForEach-Object { $_.FullName })
foreach ($e in $z.Entries) {
  $s = $e.Open(); $r = New-Object System.IO.StreamReader($s); $txt = $r.ReadToEnd(); $r.Dispose(); $s.Dispose()
  if ($e.FullName -match '\.(xml|rels)$') {
    try { $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($txt) }
    catch { $issues += "malformed XML: $($e.FullName) - $($_.Exception.Message)"; continue }
    if ($e.FullName -match '^xl/worksheets/sheet\d+\.xml$') {
      $last = -1
      foreach ($k in $doc.DocumentElement.ChildNodes) {
        $idx = if ($rank.ContainsKey($k.LocalName)) { $rank[$k.LocalName] } else { 999 }
        if ($idx -lt $last) { $issues += "element order: $($e.FullName) has <$($k.LocalName)> out of sequence" }
        $last = $idx
      }
    }
  }
  if ($e.FullName -match '_rels/.*\.rels$') {
    $baseDir = ''
    if ($e.FullName -match '^(.*/)_rels/[^/]+\.rels$') { $baseDir = $Matches[1] }
    foreach ($m in [regex]::Matches($txt, 'Target="([^"]+)"')) {
      $t = $m.Groups[1].Value
      if ($t -match '^https?://' -or $t -match '^file:') { continue }
      $full = if ($t.StartsWith('/')) { $t.TrimStart('/') } else {
        $parts = New-Object System.Collections.ArrayList
        foreach ($p in ($baseDir + $t).Split('/')) {
          if ($p -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count-1) } }
          elseif ($p -ne '.' -and $p -ne '') { [void]$parts.Add($p) }
        }
        ($parts -join '/')
      }
      if ($names -notcontains $full) { $issues += "dangling relationship: $($e.FullName) -> $t" }
    }
  }
}
$z.Dispose()

if ($issues.Count -eq 0) { Write-Output "Structural check: OK (schema element order, XML well-formedness, relationship targets)" }
else { Write-Output "Structural check FOUND $($issues.Count) issue(s):"; $issues | Select-Object -First 20 | ForEach-Object { Write-Output "  $_" } }

# --- open with alerts ENABLED and watch for a repair prompt ------------------
Write-Output "Repair check: re-opening with alerts enabled ..."
$job = Start-Job -ScriptBlock {
  Add-Type @"
using System;using System.Text;using System.Runtime.InteropServices;
public class RW{
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
 public delegate bool EnumProc(IntPtr h, IntPtr l);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
 [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
 [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, StringBuilder l);
 [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
}
"@
  $deadline = (Get-Date).AddSeconds(120)
  while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' })) {
      $found = [IntPtr]::Zero
      $cb = [RW+EnumProc]{ param($h,$l)
        $q=0; [RW]::GetWindowThreadProcessId($h,[ref]$q)|Out-Null
        if ($q -eq $p.Id) { $cn=New-Object System.Text.StringBuilder 128; [RW]::GetClassName($h,$cn,128)|Out-Null
          if ($cn.ToString() -eq '#32770') { $script:found = $h } }
        return $true }
      [RW]::EnumWindows($cb,[IntPtr]::Zero)|Out-Null
      if ($found -ne [IntPtr]::Zero) {
        $text=''
        $cb2=[RW+EnumProc]{ param($h,$l)
          $sb=New-Object System.Text.StringBuilder 2048
          [RW]::SendMessage($h,0x000D,[IntPtr]2048,$sb)|Out-Null
          if ($sb.Length -gt 0) { $script:text += $sb.ToString() + ' | ' }
          return $true }
        [RW]::EnumChildWindows($found,$cb2,[IntPtr]::Zero)|Out-Null
        Write-Output "DIALOG: $text"
        [RW]::PostMessage($found,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        return
      }
    }
  }
}
$before2 = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl2 = New-Object -ComObject Excel.Application
$xl2.Visible = $false
$xl2.DisplayAlerts = $true      # a repair prompt must NOT be suppressed here
$xl2.EnableEvents = $false
$xl2.AutomationSecurity = 1
$after2 = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$owned2 = @($after2 | Where-Object { $before2 -notcontains $_ })
$opened = $false
try {
  $wb2 = $xl2.Workbooks.Open($Dest)
  $opened = $true
  Write-Output "  opened: $($wb2.Worksheets.Count) sheets, $($wb2.VBProject.VBComponents.Count) VBA components"
  $wb2.Close($false)
} catch {
  Write-Output "  OPEN FAILED: $($_.Exception.Message)"
} finally {
  $dlg = @(Receive-Job $job -ErrorAction SilentlyContinue)
  Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue
  try { $xl2.Quit() } catch {}
  Start-Sleep -Milliseconds 800
  foreach ($procId in $owned2) { if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue } }
  if ($dlg) { Write-Output "  REPAIR/ERROR PROMPT SEEN:"; $dlg | ForEach-Object { Write-Output "    $_" } }
  elseif ($opened) { Write-Output "  no repair prompt - file opens clean" }
}
