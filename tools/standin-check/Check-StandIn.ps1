<#
.SYNOPSIS
  Runs the same RAS programs on the stand-in for Excel (ras_book.node) and on Excel itself
  (ras_com.node), and compares what they print and the workbooks they leave.

.EXAMPLE
  pwsh -File tools\standin-check\Check-StandIn.ps1
  pwsh -File tools\standin-check\Check-StandIn.ps1 -Only clipboard,more -KeepExcel
#>
param(
    [string]$Ras = "",
    [string]$Target = "",
    [string]$Out = "",
    [string[]]$Only = @(),
    [switch]$KeepExcel
)
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$repo = (Resolve-Path "$here\..\..").Path
if (-not $Target) { $Target = Join-Path $repo "target\release" }
if (-not $Ras) { $Ras = Join-Path $Target "ras.exe" }
if (-not $Out) { $Out = Join-Path $here "out" }
$comDll = Join-Path $Target "ras_com.dll"
$bookDll = Join-Path $Target "ras_book.dll"
foreach ($f in $Ras, $comDll, $bookDll) { if (-not (Test-Path $f)) { throw "$f がありません（cargo build --release）" } }

# Each scenario: the project, the entry point, and whether the workbooks are compared.
$scenarios = @(
    @{ name = "rebuild"; project = "cells"; entry = "Main.rebuild" }
    @{ name = "dates"; project = "cells"; entry = "Main.dates" }
    @{ name = "sheets"; project = "cells"; entry = "Main.sheets" }
    @{ name = "failing"; project = "cells"; entry = "Main.failing" }
    @{ name = "insertAndDelete"; project = "cells"; entry = "Main.insertAndDelete" }
    @{ name = "clipboard"; project = "cells"; entry = "Main.clipboard" }
    @{ name = "sortAndFilter"; project = "cells"; entry = "Main.sortAndFilter" }
    @{ name = "formats"; project = "cells"; entry = "Main.formats" }
    @{ name = "manual"; project = "cells"; entry = "Main.manual" }
    @{ name = "more"; project = "cells"; entry = "Main.more" }
    @{ name = "files"; project = "cells"; entry = "Tools.files" }
    @{ name = "dictionary"; project = "cells"; entry = "Tools.dictionary" }
    @{ name = "regex"; project = "cells"; entry = "Tools.regex" }
    # The same questions asked of VBA in Excel (Probe.bas), which is what both should answer.
    @{ name = "probe"; project = "cells"; entry = "Probe.probe"; vba = "Probe.RunProbe" }
    # The order of the events: the VBA build in Excel against the stand-in.
    @{ name = "events.run"; project = "events"; entry = "Main.run"; packed = $true }
    @{ name = "events.sheets"; project = "events"; entry = "Main.sheets"; packed = $true }
)
$Only = @($Only | ForEach-Object { $_ -split "," } | Where-Object { $_ })   # -File passes `a,b` as one string
if ($Only.Count -gt 0) { $scenarios = $scenarios | Where-Object { $Only -contains $_.name } }

Remove-Item $Out -Recurse -Force -ErrorAction SilentlyContinue
$work = Join-Path $Out "work"
New-Item -ItemType Directory -Force $work, "$Out\standin", "$Out\excel" | Out-Null
$bookAddon = Join-Path $Out "ras_book.node"
Copy-Item $bookDll $bookAddon
# The addons as Node loads them: copies of what cargo built (a `.node` left in target\ may be old).
$comAddon = Join-Path $Out "ras_com.node"
Copy-Item $comDll $comAddon

function Invoke-Node([string]$dir, [string]$entry, [hashtable]$vars) {
    $psi = [System.Diagnostics.ProcessStartInfo]::new("node")
    $psi.ArgumentList.Add("--enable-source-maps")
    $psi.ArgumentList.Add((Join-Path $dir "build\debug\ras-main.js"))
    $psi.ArgumentList.Add($entry)
    $psi.WorkingDirectory = $dir
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    foreach ($k in "RAS_COM_ADDON", "RAS_BOOK_ADDON", "RAS_MOCK_HOST", "RAS_WORKBOOKS", "RAS_DIALOGS", "RAS_RUN_REPORT", "RAS_ANSWERS") { $psi.Environment.Remove($k) | Out-Null }
    foreach ($k in $vars.Keys) { $psi.Environment[$k] = $vars[$k] }
    $p = [System.Diagnostics.Process]::Start($psi)
    $err = $p.StandardError.ReadToEndAsync()
    $stdout = $p.StandardOutput.ReadToEnd()
    if (-not $p.WaitForExit(120000)) { $p.Kill(); $stdout += "`n<timeout>" }
    $lines = @($stdout -split "`r?`n" | Where-Object { $_ -ne "" -and -not $_.StartsWith("[ras]") })
    @{ lines = $lines; stderr = $err.Result.Trim(); code = $p.ExitCode }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$results = @()
try {
    # The workbook the programs start from, written by Excel.
    $seed = Join-Path $work "seed.xlsx"
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    # As formulas (text): PowerShell 7 does not pass a number to a property it has passed a string to.
    $rows = @(@("顧客", "担当", "単価", "数量"), @("A社", "佐藤", "100", "2"), @("B社", "鈴木", "250.5", "4"), @("C社", "佐藤", "30", "3"))
    for ($r = 0; $r -lt 4; $r++) { for ($c = 0; $c -lt 4; $c++) { $ws.Cells.Item($r + 1, $c + 1).Formula = $rows[$r][$c] } }
    $ws.Range("C2:C4").NumberFormat = "#,##0.00"
    $wb.SaveAs($seed, 51)
    $wb.Close($false)

    foreach ($name in ($scenarios.project | Sort-Object -Unique)) {
        Copy-Item (Join-Path $here $name) (Join-Path $work $name) -Recurse
        Copy-Item $seed (Join-Path $work "$name\Book.xlsx")
        & $Ras build --target js --project (Join-Path $work $name)
        if ($LASTEXITCODE -ne 0) { throw "ras build --target js ($name) が失敗しました" }
    }

    foreach ($s in $scenarios) {
        $dir = Join-Path $work $s.project
        $book = Join-Path $dir "Book.xlsx"
        Write-Host "== $($s.name)" -ForegroundColor Cyan

        Copy-Item $seed $book -Force
        $saved = Join-Path $Out "standin\$($s.name)"
        $a = Invoke-Node $dir $s.entry @{ RAS_BOOK_ADDON = $bookAddon; RAS_WORKBOOKS = $saved }

        Copy-Item $seed $book -Force
        $excelBook = Join-Path $Out "excel\$($s.name).xlsx"
        if ($s.packed) {
            # Excel raises its events for VBA only: the VBA build, written into a copy of the
            # workbook, is run as a macro, and what it returns is what it would have printed.
            # (A compile error of the VBA stops Excel in the VBA editor, and this script with it.)
            $packed = Join-Path $dir "Packed.xlsm"
            if (-not (Test-Path $packed)) {
                & $Ras pack --project $dir --convert --output $packed | Out-Null
                if ($LASTEXITCODE -ne 0) { throw "ras pack ($($s.name)) が失敗しました" }
            }
            $macroBook = Join-Path $dir "Run.xlsm"
            Copy-Item $packed $macroBook -Force
            $wb = $xl.Workbooks.Open($macroBook)
            # Deleting a sheet asks; the alerts can be on again after the macro before.
            $xl.DisplayAlerts = $false
            $text = "$($xl.Run("Run.xlsm!$($s.entry)"))"
            $b = @{ lines = @($text -split "`r?`n" | Where-Object { $_ -ne "" }); stderr = ""; code = 0 }
            $excelBook = $null
            $xl.EnableEvents = $false
        } else {
            $wb = $xl.Workbooks.Open($book)
            $b = Invoke-Node $dir $s.entry @{ RAS_COM_ADDON = $comAddon; RAS_DIALOGS = "dismiss" }
            $xl.CutCopyMode = 0
            $xl.Calculation = -4105
            $xl.ScreenUpdating = $true
            try { $wb.SaveCopyAs($excelBook) } catch { Write-Host "  SaveCopyAs: $_" -ForegroundColor Yellow }
        }
        foreach ($open in @($xl.Workbooks)) { $open.Close($false) }
        $xl.EnableEvents = $true

        # VBA itself, in a module imported into the same workbook.
        $v = $null
        if ($s.vba) {
            Copy-Item $seed $book -Force
            $wb = $xl.Workbooks.Open($book)
            $bas = Join-Path $work "Probe.bas"
            [IO.File]::WriteAllText($bas, ((Get-Content (Join-Path $here "Probe.bas") -Raw) -replace "`r?`n", "`r`n"), [Text.Encoding]::ASCII)
            [void]$wb.VBProject.VBComponents.Import($bas)
            $v = @("$($xl.Run($s.vba))" -split "`n" | Where-Object { $_ -ne "" })
            foreach ($open in @($xl.Workbooks)) { $open.Close($false) }
        }

        $n = [Math]::Max($a.lines.Count, $b.lines.Count)
        if ($v) { $n = [Math]::Max($n, $v.Count) }
        $diffs = @()
        for ($i = 0; $i -lt $n; $i++) {
            $x = if ($i -lt $a.lines.Count) { $a.lines[$i] } else { "<なし>" }
            $y = if ($i -lt $b.lines.Count) { $b.lines[$i] } else { "<なし>" }
            $z = if ($null -eq $v) { $null } elseif ($i -lt $v.Count) { $v[$i] } else { "<なし>" }
            if ($x -cne $y -or ($null -ne $z -and $x -cne $z)) { $diffs += @{ line = $i + 1; standin = $x; excel = $y; vba = $z } }
        }
        foreach ($d in $diffs) {
            Write-Host ("  {0,2}: 代役  {1}" -f $d.line, $d.standin) -ForegroundColor Yellow
            Write-Host ("      Excel {0}" -f $d.excel) -ForegroundColor Yellow
            if ($null -ne $d.vba) { Write-Host ("      VBA   {0}" -f $d.vba) -ForegroundColor Yellow }
        }
        foreach ($side in @(@("代役", $a), @("Excel", $b))) {
            if ($side[1].code -ne 0 -or $side[1].stderr) { Write-Host "  $($side[0]) の終了コード $($side[1].code): $($side[1].stderr)" -ForegroundColor Yellow }
        }

        $bookDiff = $null
        $standinBook = Join-Path $saved "Book.xlsx"
        if ($null -eq $excelBook) {
        } elseif ((Test-Path $standinBook) -and (Test-Path $excelBook)) {
            $bookDiff = @(& $Ras book diff $standinBook $excelBook 2>&1 | ForEach-Object { "$_" })
            if ($LASTEXITCODE -ne 0) { $bookDiff | ForEach-Object { Write-Host "  book: $_" -ForegroundColor Yellow } } else { $bookDiff = @() }
        } elseif (Test-Path $excelBook) {
            Write-Host "  代役はブックを保存していません（変更なし）"
        }
        if ($diffs.Count -eq 0 -and -not $bookDiff) { Write-Host "  同じ" -ForegroundColor Green }
        $results += @{ name = $s.name; entry = $s.entry; standin = $a; excel = $b; printed = $diffs; book = $bookDiff }
    }
} finally {
    if (-not $KeepExcel) {
        # Whatever went wrong, Excel is not left running.
        try { $xl.EnableEvents = $false; foreach ($open in @($xl.Workbooks)) { $open.Close($false) } } catch { }
        try { $xl.Quit() } catch { }
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
$results | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $Out "results.json") -Encoding utf8NoBOM
$bad = @($results | Where-Object { $_.printed.Count -gt 0 -or $_.book })
Write-Host ("{0} 本中 {1} 本に違い（{2}）" -f $results.Count, $bad.Count, (Join-Path $Out "results.json"))
exit ([int]($bad.Count -gt 0))
