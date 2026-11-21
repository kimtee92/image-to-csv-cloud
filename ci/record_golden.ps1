# Record what Excel 365 computes for the recalc check's test exports.
#
# Run on a Windows PC with Excel and Python 3.11, from the repo root:
#   powershell -ExecutionPolicy Bypass -File ci\record_golden.ps1
# It writes ci\golden\<template>.json: every formula cell on the filled rows of
# the Timekeeping sheet and its Excel value. Commit those files. Re-record only
# when a change is MEANT to alter computed values, and review the diff.
param([string]$Python = "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe")
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("golden-" + [guid]::NewGuid())
& $Python (Join-Path $repo 'ci\recalc_check.py') export $tmp
if ($LASTEXITCODE) { throw "export failed" }
$manifest = Get-Content (Join-Path $tmp 'manifest.json') -Raw | ConvertFrom-Json
$errText = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?';
              -2146826288 = '#NULL!'; -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!' }
$outDir = Join-Path $repo 'ci\golden'
New-Item -ItemType Directory -Force $outDir | Out-Null
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  foreach ($p in $manifest.PSObject.Properties) {
    $tpl = $p.Name; $m = $p.Value
    $wb = $xl.Workbooks.Open((Join-Path $tmp "$tpl.xlsx"), 0, $true)
    $xl.CalculateFull()
    $ws = $wb.Worksheets.Item($m.sheet)
    $vals = [ordered]@{}
    # The cell list comes from the file (see recalc_check.export): Excel can't
    # enumerate formula cells on a protected sheet, but can read their values.
    $lastCol = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1
    $block = $ws.Range($ws.Cells.Item($m.first_row, 1), $ws.Cells.Item($m.last_row, $lastCol)).Value2
    foreach ($addr in $m.cells) {
      $cell = $ws.Range($addr)
      $x = $block[($cell.Row - $m.first_row + 1), $cell.Column]
      if ($x -is [int] -and $errText.ContainsKey($x)) { $x = $errText[$x] }
      elseif ($null -eq $x) { $x = '' }
      $vals[$addr] = $x
    }
    if ($vals.Count -ne @($m.cells).Count) { throw "$($tpl): recorded $($vals.Count) of $(@($m.cells).Count) cells" }
    $json = $vals | ConvertTo-Json -Compress
    [IO.File]::WriteAllText((Join-Path $outDir "$tpl.json"), $json)
    "{0,-13} {1,6} formula cells recorded" -f $tpl, $vals.Count
    $wb.Close($false)
  }
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
