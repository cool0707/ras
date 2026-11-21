<#
.SYNOPSIS
    RAS が書き込むユーザーフォームを、実物の Excel で確かめる。

.DESCRIPTION
    forms\*.form.json（全プロパティを既定と違う値にしたテスト用フォーム）について、次の 3 つを比べる。

      1. RAS → Excel : ras pack で書き込んだブックを Excel で開き、VBA エディターのデザイナーから
                       各プロパティを読む。FormSpec に書いた値と同じか。
      2. Excel → RAS : 同じフォームを Excel のデザイナーで作って保存し（参照ブック）、ras form pull で
                       取り込む。Excel が持っている値と同じになるか。
      3. 見た目      : 両方のブックのフォームを表示して画面を撮り、ピクセル単位で比べる。
                       ComboBox / ListBox の項目（実行時に入るもの）もここで比べる。

    結果は -Out のフォルダー（既定は tools\excel-check\out）の report.html と results.json。
    違いがあれば終了コード 1。

    必要なもの: Windows の Excel、Windows PowerShell 5.1（powershell.exe）、ras（PATH にあるか -Ras で指定）、
    Excel のオプション［トラスト センター］［マクロの設定］の
    「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\excel-check\Check-Forms.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\excel-check\Check-Forms.ps1 -Ras C:\tools\ras.exe -Forms CheckLabels,CheckButtons
#>
[CmdletBinding()]
param(
    # ras の実行ファイル。
    [string]$Ras = "ras",
    # 結果を書き出すフォルダー（実行のたびに作り直す）。
    [string]$Out = (Join-Path $PSScriptRoot "out"),
    # 確かめるフォーム（省略時は forms\ のすべて）。
    [string[]]$Forms = @(),
    # 見た目の比較で、違っていてもよいピクセル数（TextBox のカーソルの点滅など）。
    [int]$Tolerance = 40,
    # 見た目の比較をしない。
    [switch]$SkipVisual,
    # 終わった後も Excel を開いたままにする。
    [switch]$KeepExcel
)

$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$previousEncoding = $null
try {
    # ras の出力（UTF-8）をそのまま読む。
    $previousEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch { }
Add-Type -AssemblyName System.Drawing

function Write-Step([string]$text) { Write-Host "== $text" -ForegroundColor Cyan }

# ============================================================================ ras

$rasCommand = Get-Command $Ras -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $rasCommand) { throw "ras が見つかりません。PATH に置くか、-Ras で ras.exe のパスを指定してください。" }
$Ras = $rasCommand.Definition

function Invoke-Ras([string[]]$arguments) {
    # Windows PowerShell は、Stop のままだと標準エラーへの出力（警告）で止まってしまう。
    $ErrorActionPreference = "Continue"
    $output = & $Ras @arguments 2>&1 | ForEach-Object { "$_" }
    $text = ($output -join "`n")
    if ($LASTEXITCODE -ne 0) { throw "ras $($arguments -join ' ') が失敗しました:`n$text" }
    return $text
}

# ============================================================================ the forms

function Read-Json([string]$path) { Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json }

# A field of a JSON object, or $null.
function Get-Field($object, [string]$name) {
    if ($null -eq $object) { return $null }
    $p = $object.PSObject.Properties[$name]
    if ($p) { return $p.Value }
    return $null
}

$cases = @()
foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $here "forms") -Filter "*.form.json" | Sort-Object Name)) {
    $spec = Read-Json $file.FullName
    if ($Forms.Count -gt 0 -and -not ($Forms -contains $spec.form.name)) { continue }
    $cases += , $spec
}
if ($cases.Count -eq 0) { throw "確かめるフォームがありません。" }

# ============================================================================ values

$systemColors = @("scrollBar", "desktop", "activeCaption", "inactiveCaption", "menu", "window", "windowFrame", "menuText", "windowText",
    "captionText", "activeBorder", "inactiveBorder", "appWorkspace", "highlight", "highlightText", "buttonFace", "buttonShadow", "grayText",
    "buttonText", "inactiveCaptionText", "buttonHighlight", "buttonDarkShadow", "buttonLight", "infoText", "infoBackground", "", "hotLight",
    "gradientActiveCaption", "gradientInactiveCaption", "menuHighlight", "menuBar")
# 0x80000000 と書くと、Windows PowerShell では負の Int32 になる。
$highBit = [int64]2147483648
$twoTo32 = [int64]4294967296

# A FormSpec color ({ kind, value }) as an OLE color (unsigned).
function ConvertTo-Ole($color) {
    $kind = Get-Field $color "kind"; $value = [string](Get-Field $color "value")
    switch ($kind) {
        "rgb" {
            $n = [int64][Convert]::ToUInt32($value.Substring(1), 16)
            return [uint32]((($n -band 255) * 65536) + ($n -band 65280) + [Math]::Floor($n / 65536))
        }
        "system" { return [uint32]($highBit + [array]::IndexOf($systemColors, $value)) }
        default { return [Convert]::ToUInt32($value.Substring(2), 16) }
    }
}

# A Long of VBA as an unsigned OLE color.
function ConvertFrom-Long($n) {
    $v = [int64]$n
    if ($v -lt 0) { $v += $twoTo32 }
    return [uint32]$v
}

# ColumnWidths ("30 pt;;1 in") as points; -1 for an empty entry.
function ConvertFrom-Widths([string]$text) {
    $out = @()
    if ($null -eq $text) { return , $out }
    foreach ($entry in $text.Split(";")) {
        $e = $entry.Trim().ToLowerInvariant().Replace(",", ".")
        if ($e -eq "" -or $e -eq "-1") { $out += -1.0; continue }
        $unit = 1.0
        if ($e.EndsWith("pt")) { $e = $e.Substring(0, $e.Length - 2).Trim() }
        elseif ($e.EndsWith("in")) { $e = $e.Substring(0, $e.Length - 2).Trim(); $unit = 72.0 }
        elseif ($e.EndsWith("cm")) { $e = $e.Substring(0, $e.Length - 2).Trim(); $unit = 72.0 / 2.54 }
        $out += [double]::Parse($e, [Globalization.CultureInfo]::InvariantCulture) * $unit
    }
    return , $out
}

$toggles = @("CheckBox", "OptionButton", "ToggleButton")
$geometry = @("left", "top", "width", "height")

# A FormSpec value in the form the comparison uses.
function ConvertFrom-Spec([string]$field, $value, [string]$type) {
    if ($null -eq $value) { return $null }
    if ($field -in @("foreColor", "backColor", "borderColor")) { return ConvertTo-Ole $value }
    if ($field -eq "columnWidths") { return , (ConvertFrom-Widths $value) }
    if ($field -eq "tabs") { return ConvertFrom-Tabs $value }
    if ($field -eq "value" -and $toggles -contains $type) {
        if ($value -is [bool]) { return $value }
        if ($value -is [string]) {
            $t = $value.Trim().ToLowerInvariant()
            if ($t -eq "null") { return "Null" }
            return ($t -in @("true", "1", "-1"))
        }
        return ([double]$value -ne 0)
    }
    if ($field -in @("picture", "mouseIcon")) {
        if ((Get-Field $value "path") -and -not (Get-Field $value "remove")) { return "present" }
        return "none"
    }
    if ($field -eq "font") {
        $font = [ordered]@{}
        foreach ($p in $value.PSObject.Properties) {
            if ($p.Name -eq "pitchAndFamily") { continue }
            $v = $p.Value
            if ($v -isnot [bool] -and $v -isnot [string]) { $v = [double]$v }
            $font[$p.Name] = $v
        }
        return $font
    }
    if ($value -is [bool] -or $value -is [string]) { return $value }
    return [double]$value
}

# A value read through Application.Run in the form the comparison uses.
function ConvertFrom-Vba([string]$field, $value) {
    if ($value -is [DBNull]) { return "Null" }
    if ($null -eq $value) { return $null }
    if ($value -is [string] -and $value.StartsWith("#ERROR")) { return $value }
    if ($field -in @("foreColor", "backColor", "borderColor")) { return ConvertFrom-Long $value }
    if ($field -eq "columnWidths") { return , (ConvertFrom-Widths ([string]$value)) }
    if ($value -is [bool] -or $value -is [string]) { return $value }
    return [double]$value
}

# The font as RasCheck_FontInfo gives it.
function ConvertFrom-FontInfo([string]$text) {
    if ($text.StartsWith("#ERROR")) { return $text }
    $p = $text.Split("|")
    $culture = [Globalization.CultureInfo]::InvariantCulture
    return [ordered]@{
        name = $p[0]; size = [double]::Parse($p[1], $culture); bold = ($p[2] -eq "True"); italic = ($p[3] -eq "True")
        underline = ($p[4] -eq "True"); strikeThrough = ($p[5] -eq "True"); charset = [double]::Parse($p[6], $culture)
    }
}

# The tabs of a TabStrip of the FormSpec as RasCheck_TabsInfo writes them.
function ConvertFrom-Tabs($tabs) {
    $lines = @()
    foreach ($t in @($tabs)) {
        $name = [string](Get-Field $t "name")
        $caption = Get-Field $t "caption"; if ($null -eq $caption) { $caption = $name }
        $enabled = Get-Field $t "enabled"; if ($null -eq $enabled) { $enabled = $true }
        $visible = Get-Field $t "visible"; if ($null -eq $visible) { $visible = $true }
        $lines += (@($name, $caption, [string](Get-Field $t "controlTipText"), [string](Get-Field $t "tag"),
                [string](Get-Field $t "accelerator"), [string][bool]$enabled, [string][bool]$visible) -join "`t")
    }
    return ($lines -join "`n")
}

# The items of a list of the FormSpec as RasCheck_ListInfo writes them (without the ListIndex line).
function ConvertFrom-List($list, $columns) {
    $text = ""
    $n = [Math]::Max(1, [int]$columns)
    foreach ($item in @($list)) {
        $cells = @($item)
        $row = @()
        for ($i = 0; $i -lt $n; $i++) { if ($i -lt $cells.Count) { $row += [string]$cells[$i] } else { $row += "" } }
        $text += ($row -join "`t") + "`n"
    }
    return $text
}

# Whether two compared values are the same.
function Test-Same($expected, $actual, [string]$field) {
    if ($null -eq $expected) { return ($null -eq $actual) }
    if ($null -eq $actual) { return $false }
    if ($expected -is [System.Collections.IDictionary]) {
        if ($actual -isnot [System.Collections.IDictionary]) { return $false }
        foreach ($key in $expected.Keys) {
            if (-not $actual.Contains($key)) { return $false }
            if (-not (Test-Same $expected[$key] $actual[$key] $key)) { return $false }
        }
        return $true
    }
    # PowerShell unrolls an array of one item that a function returns.
    if ($expected -is [array] -or $actual -is [array]) {
        $expected = @($expected); $actual = @($actual)
        if ($actual.Count -ne $expected.Count) { return $false }
        for ($i = 0; $i -lt $actual.Count; $i++) { if (-not (Test-Same $expected[$i] $actual[$i] $field)) { return $false } }
        return $true
    }
    if ($expected -is [bool]) { return ($actual -is [bool] -and $expected -eq $actual) }
    if ($expected -is [string]) {
        if ($actual -isnot [string]) { return $false }
        if ($field -eq "name") { return ($expected -ieq $actual) }
        return ($expected -ceq $actual)
    }
    if ($actual -is [bool] -or $actual -is [string]) { return $false }
    # One pixel for positions and sizes, which Excel keeps on its own grid.
    $limit = 0.1; if ($geometry -contains $field) { $limit = 0.76 }
    return ([Math]::Abs([double]$expected - [double]$actual) -le $limit)
}

# Only the keys of a dictionary another one has (the font fields a FormSpec names).
function Select-Keys($value, $like) {
    if ($value -isnot [System.Collections.IDictionary] -or $like -isnot [System.Collections.IDictionary]) { return $value }
    $out = [ordered]@{}
    foreach ($key in $like.Keys) { if ($value.Contains($key)) { $out[$key] = $value[$key] } }
    return $out
}

function Format-Value($v) {
    if ($null -eq $v) { return "（なし）" }
    if ($v -is [System.Collections.IDictionary]) { return (($v.Keys | ForEach-Object { "$_=$($v[$_])" }) -join ", ") }
    if ($v -is [array]) { return "[" + (($v | ForEach-Object { "$_" }) -join "; ") + "]" }
    if ($v -is [uint32]) { return ("0x{0:X8}" -f $v) }
    return [string]$v
}

# ============================================================================ fields

$skipped = @("id", "name", "type", "progId", "parentId", "zIndex", "observed", "unsupported", "properties", "controls", "warnings")

# The VBA name of a FormSpec field (the FormSpec uses the VBA names in lower camel case).
function Get-VbaName([string]$field, [string]$type) {
    switch ($field) {
        "imeMode" { return "IMEMode" }
        "helpContextId" { return "HelpContextID" }
        "strikeThrough" { return "Strikethrough" }
        "value" { if ($type -in @("TextBox", "ComboBox")) { return "Text" } return "Value" }
        "selectedIndex" { return "Value" }
    }
    return $field.Substring(0, 1).ToUpperInvariant() + $field.Substring(1)
}

# The fields of a control (or of the form) of the FormSpec that are compared at design time.
function Get-DesignFields($object, [string]$type) {
    $fields = @()
    foreach ($p in $object.PSObject.Properties) {
        $f = $p.Name
        if ($skipped -contains $f) { continue }
        # The items of lists and their selection are put in by code when the form loads: compared at run time.
        if ($type -in @("ComboBox", "ListBox") -and $f -in @("list", "selectedIndex")) { continue }
        $fields += $f
    }
    return $fields
}

# Order in which the fields of a control are set: what others depend on first.
$firstFields = @("left", "top", "width", "height", "min", "max", "columnCount", "multiSelect", "style", "tripleState", "multiLine",
    "listStyle", "rowSource", "whatsThisHelp")
$lastFields = @("value", "text", "selectedIndex", "autoSize")

function Get-SetOrder($fields) {
    $fields = @($fields)
    $ordered = @()
    foreach ($f in $firstFields) { if ($fields -contains $f) { $ordered += $f } }
    foreach ($f in $fields) { if (-not ($firstFields -contains $f) -and -not ($lastFields -contains $f)) { $ordered += $f } }
    foreach ($f in $lastFields) { if ($fields -contains $f) { $ordered += $f } }
    return $ordered
}

# ============================================================================ the run

$work = Join-Path $Out "work"
if (Test-Path $Out) { Remove-Item -LiteralPath $Out -Recurse -Force }
New-Item -ItemType Directory -Path $work -Force | Out-Null
$Out = (Resolve-Path $Out).Path
$work = (Resolve-Path $work).Path
New-Item -ItemType Directory -Path (Join-Path $Out "images") -Force | Out-Null

Write-Step "RAS でフォームを書き込む（ras pack）"
$rasProject = Join-Path $work "ras"
Invoke-Ras @("init", $rasProject, "--workbook", "RasForms.xlsm") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $rasProject "forms") -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $here "forms\assets") -Destination (Join-Path $rasProject "forms") -Recurse
foreach ($spec in $cases) { Copy-Item -LiteralPath (Join-Path $here "forms\$($spec.form.name).form.json") -Destination (Join-Path $rasProject "forms") }
Write-Host (Invoke-Ras @("--project", $rasProject, "pack"))
$rasBook = Join-Path $rasProject "RasForms.xlsm"
if (-not (Test-Path $rasBook)) { throw "$rasBook がありません。" }

Write-Step "Excel を起動する"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
# 開くブックのマクロを有効にする（この Excel のインスタンスだけ）。
$xl.AutomationSecurity = 1

$rows = New-Object System.Collections.ArrayList
$visuals = New-Object System.Collections.ArrayList
$setupErrors = New-Object System.Collections.ArrayList
$pullLog = ""
$rasWb = $null
$ref = $null

# Runs a function of the RasCheck module of a workbook.
function Invoke-Check($book, [string]$name) {
    $macro = "'" + $book.Name + "'!RasCheck." + $name
    $a = $args
    switch ($a.Count) {
        0 { $r = $xl.Run($macro) }
        1 { $r = $xl.Run($macro, $a[0]) }
        2 { $r = $xl.Run($macro, $a[0], $a[1]) }
        3 { $r = $xl.Run($macro, $a[0], $a[1], $a[2]) }
        4 { $r = $xl.Run($macro, $a[0], $a[1], $a[2], $a[3]) }
        default { throw "引数が多すぎます" }
    }
    # As one object: PowerShell would enumerate a COM object that can be enumerated.
    return , $r
}

function Import-Helpers($book) {
    try { $null = $book.VBProject.VBComponents.Count }
    catch { throw "VBA プロジェクトにアクセスできません。Excel のオプション［トラスト センター］［マクロの設定］で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにしてください。" }
    $null = $book.VBProject.VBComponents.Import((Join-Path $here "RasCheck.bas"))
}

# Data the controls bound to cells (ControlSource, RowSource) read: the same in both workbooks.
function Set-SheetData($book) {
    $sheet = $book.Worksheets.Item(1)
    if ($sheet.Name -ne "Sheet1") { $sheet.Name = "Sheet1" }
    $cells = @(@("東京", "Tokyo"), @("大阪", "Osaka"), @("名古屋", "Nagoya"))
    for ($r = 0; $r -lt 3; $r++) { for ($c = 0; $c -lt 2; $c++) { $sheet.Cells.Item($r + 1, $c + 1).Value2 = $cells[$r][$c] } }
    $values = @("セルの値", "大阪", "名古屋", $true, 3, 7, $true, $false)
    for ($i = 0; $i -lt $values.Count; $i++) { $sheet.Cells.Item($i + 1, 4).Value2 = $values[$i] }
}

function Add-SetupError([string]$where, $message) {
    if ($message) { [void]$setupErrors.Add("${where}: $message") }
}

# A value of the FormSpec as it is set through Application.Run.
function ConvertTo-VbaValue([string]$field, $value, [string]$type) {
    if ($field -in @("foreColor", "backColor", "borderColor")) {
        $ole = [int64](ConvertTo-Ole $value)
        if ($ole -ge $highBit) { $ole -= $twoTo32 }
        return [int]$ole
    }
    if ($field -eq "value" -and $toggles -contains $type) {
        $c = ConvertFrom-Spec $field $value $type
        if ($c -is [string]) { return "<Null>" }
        return [bool]$c
    }
    if ($value -is [bool] -or $value -is [string]) { return $value }
    $d = [double]$value
    if ($d -eq [Math]::Floor($d) -and [Math]::Abs($d) -lt 2147483647) { return [int]$d }
    return $d
}

function Set-Font($book, $target, $font, [string]$where) {
    $f = Invoke-Check $book "RasCheck_GetObject" $target "Font"
    if ($null -eq $f) { Add-SetupError "$where.font" "Font を取れません"; return }
    foreach ($p in $font.PSObject.Properties) {
        if ($p.Name -eq "pitchAndFamily") { continue }
        Add-SetupError "$where.font.$($p.Name)" (Invoke-Check $book "RasCheck_Let" $f (Get-VbaName $p.Name "") (ConvertTo-VbaValue $p.Name $p.Value ""))
    }
}

function Get-AssetPath($value) { return Join-Path $here ([string](Get-Field $value "path")).Replace("/", "\") }

# The outer size of a form minus its client size (title bar and borders), measured on a new form.
$frameDx = 0.0
$frameDy = 0.0
$frameMeasured = $false

# Makes the form of a case in the Excel reference through the designer, as a person would.
function Build-Reference($book, $spec) {
    $formName = $spec.form.name
    Add-SetupError $formName (Invoke-Check $book "RasCheck_NewForm" $formName)
    if (-not $script:frameMeasured) {
        $script:frameDx = [double](Invoke-Check $book "RasCheck_FormGet" $formName "Width") - [double](Invoke-Check $book "RasCheck_FormGet" $formName "InsideWidth")
        $script:frameDy = [double](Invoke-Check $book "RasCheck_FormGet" $formName "Height") - [double](Invoke-Check $book "RasCheck_FormGet" $formName "InsideHeight")
        $script:frameMeasured = $true
    }
    $names = @{}
    foreach ($c in $spec.controls) { $names[$c.id] = $c.name }
    $deferred = @()
    $tabIndexes = @()
    foreach ($c in $spec.controls) {
        $type = $c.type
        $where = "$formName.$($c.name)"
        $parent = ""; if (Get-Field $c "parentId") { $parent = $names[$c.parentId] }
        try { $ctl = Invoke-Check $book "RasCheck_AddControl" $formName $parent "Forms.$type.1" $c.name }
        catch { Add-SetupError $where "コントロールを追加できません: $($_.Exception.Message)"; continue }
        if ($null -eq $ctl) { Add-SetupError $where "コントロールを追加できません"; continue }
        if ($type -eq "TabStrip") {
            foreach ($t in @($c.tabs)) {
                if ($null -eq $t) { continue }
                $caption = Get-Field $t "caption"; if ($null -eq $caption) { $caption = $t.name }
                $tab = Invoke-Check $book "RasCheck_AddTab" $ctl $t.name $caption
                foreach ($k in @("controlTipText", "tag", "accelerator", "enabled", "visible")) {
                    $v = Get-Field $t $k
                    if ($null -ne $v) { Add-SetupError "$where.tabs.$($t.name).$k" (Invoke-Check $book "RasCheck_Let" $tab (Get-VbaName $k "") $v) }
                }
            }
        }
        foreach ($f in (Get-SetOrder (Get-DesignFields $c $type))) {
            if ($f -eq "tabs") { continue }
            $value = Get-Field $c $f
            if ($f -eq "font") { Set-Font $book $ctl $value $where; continue }
            if ($f -in @("picture", "mouseIcon")) {
                Add-SetupError "$where.$f" (Invoke-Check $book "RasCheck_SetPicture" $ctl (Get-VbaName $f $type) (Get-AssetPath $value))
                continue
            }
            # The page a MultiPage shows can only be chosen once its pages are there.
            if ($type -eq "MultiPage" -and $f -eq "selectedIndex") { $deferred += , @($ctl, [int]$value, "$where.selectedIndex"); continue }
            # Setting a TabIndex moves the others: set them all at the end, smallest first.
            if ($f -eq "tabIndex") { $tabIndexes += , @($ctl, [int]$value, "$where.tabIndex"); continue }
            Add-SetupError "$where.$f" (Invoke-Check $book "RasCheck_Let" $ctl (Get-VbaName $f $type) (ConvertTo-VbaValue $f $value $type))
        }
    }
    foreach ($d in $deferred) { Add-SetupError $d[2] (Invoke-Check $book "RasCheck_Let" $d[0] "Value" $d[1]) }
    foreach ($d in @($tabIndexes | Sort-Object { $_[1] })) { Add-SetupError $d[2] (Invoke-Check $book "RasCheck_Let" $d[0] "TabIndex" $d[1]) }

    # The form itself, after its controls: a control made later would take the form's font and colors.
    $form = $spec.form
    $where = "$formName.(form)"
    foreach ($f in (Get-SetOrder (Get-DesignFields $form "UserForm"))) {
        $value = Get-Field $form $f
        if ($f -eq "font") { Set-Font $book (Invoke-Check $book "RasCheck_Design" $formName) $value $where; continue }
        if ($f -in @("picture", "mouseIcon")) {
            Add-SetupError "$where.$f" (Invoke-Check $book "RasCheck_FormSetPicture" $formName (Get-VbaName $f "UserForm") (Get-AssetPath $value))
            continue
        }
        # The FormSpec gives the client size; the VBA editor's Width / Height are the outer size.
        if ($f -eq "width") { $value = [double]$value + $script:frameDx }
        if ($f -eq "height") { $value = [double]$value + $script:frameDy }
        Add-SetupError "$where.$f" (Invoke-Check $book "RasCheck_FormLet" $formName (Get-VbaName $f "UserForm") (ConvertTo-VbaValue $f $value "UserForm"))
    }
}

# Reads a field of a control (or of the form when $control is $null) at design time.
function Read-DesignField($book, [string]$formName, $control, [string]$field, [string]$type) {
    $target = $control
    if ($null -eq $control) { $target = Invoke-Check $book "RasCheck_Design" $formName }
    if ($field -eq "font") { return ConvertFrom-FontInfo (Invoke-Check $book "RasCheck_FontInfo" $target) }
    if ($field -eq "tabs") { return [string](Invoke-Check $book "RasCheck_TabsInfo" $target) }
    if ($field -in @("picture", "mouseIcon")) {
        $info = [string](Invoke-Check $book "RasCheck_PictureInfo" $target (Get-VbaName $field $type))
        if ($info.StartsWith("#ERROR")) { return $info }
        if ($info) { return "present" }
        return "none"
    }
    if ($null -eq $control) {
        $value = ConvertFrom-Vba $field (Invoke-Check $book "RasCheck_FormGet" $formName (Get-VbaName $field $type))
        if ($field -eq "width" -and $value -is [double]) { $value -= $script:frameDx }
        if ($field -eq "height" -and $value -is [double]) { $value -= $script:frameDy }
        return $value
    }
    return ConvertFrom-Vba $field (Invoke-Check $book "RasCheck_Get" $control (Get-VbaName $field $type))
}

# A field of a pulled control (or form), in the form the comparison uses.
function Read-PulledField($object, [string]$field, [string]$type) {
    if ($null -eq $object) { return $null }
    $value = Get-Field $object $field
    if ($null -eq $value -and $field -in @("text", "value") -and $type -in @("TextBox", "ComboBox")) {
        $other = "text"; if ($field -eq "text") { $other = "value" }
        $value = Get-Field $object $other
    }
    return ConvertFrom-Spec $field $value $type
}

# One compared property. $pulledValue "―" when there is nothing to pull (run-time lists).
function Add-Row([string]$form, [string]$control, [string]$field, $expected, $ras, $reference, $pulledValue, [string]$note, [bool]$counts) {
    $rasOk = Test-Same $expected $ras $field
    $refOk = Test-Same $expected $reference $field
    $refBroken = ($reference -is [string] -and $reference.StartsWith("#ERROR"))
    if ($pulledValue -is [string] -and $pulledValue -eq "―") { $pullOk = $true }
    elseif ($refBroken) { $pullOk = Test-Same $expected $pulledValue $field }
    elseif (-not $refOk -and $null -eq $pulledValue) { $pullOk = $true }
    else { $pullOk = Test-Same (Select-Keys $reference $expected) $pulledValue $field }
    $problems = @()
    $serious = $false
    if (-not $rasOk) {
        if (-not $refOk -and (Test-Same (Select-Keys $reference $expected) $ras $field)) { $problems += "Excel が値を変える（参照も同じ値）" }
        else { $problems += "RAS→Excel"; $serious = $true }
    }
    if (-not $pullOk) { $problems += "Excel→RAS"; $serious = $true }
    if (-not $refOk -and $rasOk) { $problems += "参照の作成" }
    $status = "OK"
    if ($problems.Count -gt 0) { if ($serious -and $counts) { $status = "NG" } else { $status = "注意" } }
    [void]$rows.Add([pscustomobject][ordered]@{
            form = $form; control = $control; field = $field; status = $status; problems = ($problems -join "、"); note = $note
            expected = (Format-Value $expected); ras = (Format-Value $ras); reference = (Format-Value $reference); pulled = (Format-Value $pulledValue)
        })
}

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public static class RasCapture {
    delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT point);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);

    public static uint ProcessOf(IntPtr hwnd) { uint pid; GetWindowThreadProcessId(hwnd, out pid); return pid; }

    // A visible user form window (class ThunderDFrame) of the process: the one with this caption
    // when there is one.
    public static IntPtr Find(string caption, uint pid) {
        IntPtr exact = IntPtr.Zero, any = IntPtr.Zero;
        EnumWindows(delegate (IntPtr hwnd, IntPtr lParam) {
            StringBuilder name = new StringBuilder(256);
            GetClassName(hwnd, name, name.Capacity);
            if (name.ToString() != "ThunderDFrame" || !IsWindowVisible(hwnd) || ProcessOf(hwnd) != pid) return true;
            if (any == IntPtr.Zero) any = hwnd;
            StringBuilder text = new StringBuilder(512);
            GetWindowText(hwnd, text, text.Capacity);
            if (text.ToString() != caption) return true;
            exact = hwnd;
            return false;
        }, IntPtr.Zero);
        return exact != IntPtr.Zero ? exact : any;
    }

    // The client area of a window, drawn by the window itself (covered parts included).
    public static Bitmap CaptureClient(IntPtr hwnd) {
        RECT window, client;
        GetWindowRect(hwnd, out window);
        GetClientRect(hwnd, out client);
        POINT origin = new POINT();
        ClientToScreen(hwnd, ref origin);
        Bitmap whole = new Bitmap(Math.Max(1, window.Right - window.Left), Math.Max(1, window.Bottom - window.Top), PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(whole)) {
            IntPtr hdc = g.GetHdc();
            // PW_RENDERFULLCONTENT (Windows 8.1 and later), else the classic way.
            if (!PrintWindow(hwnd, hdc, 2)) PrintWindow(hwnd, hdc, 0);
            g.ReleaseHdc(hdc);
        }
        Rectangle area = new Rectangle(origin.X - window.Left, origin.Y - window.Top, Math.Max(1, client.Right), Math.Max(1, client.Bottom));
        area.Intersect(new Rectangle(0, 0, whole.Width, whole.Height));
        Bitmap part = whole.Clone(area, PixelFormat.Format32bppArgb);
        whole.Dispose();
        return part;
    }

    // Pixels that differ (-1 when the sizes differ); writes an image with them in red.
    public static int Compare(Bitmap a, Bitmap b, string diffPath) {
        if (a.Width != b.Width || a.Height != b.Height) return -1;
        int n = 0;
        using (Bitmap d = new Bitmap(a.Width, a.Height, PixelFormat.Format32bppArgb)) {
            for (int y = 0; y < a.Height; y++) {
                for (int x = 0; x < a.Width; x++) {
                    Color ca = a.GetPixel(x, y), cb = b.GetPixel(x, y);
                    if (ca.ToArgb() != cb.ToArgb()) { n++; d.SetPixel(x, y, Color.Red); }
                    else { int g = 190 + (ca.R + ca.G + ca.B) / 12; d.SetPixel(x, y, Color.FromArgb(g, g, g)); }
                }
            }
            d.Save(diffPath, ImageFormat.Png);
        }
        return n;
    }
}
"@

# The items of a list of the case as RasCheck_AddItems takes them.
function Get-ItemsText($control) {
    $lines = @()
    foreach ($item in @($control.list)) { $lines += ((@($item) | ForEach-Object { [string]$_ }) -join "`t") }
    return ($lines -join "`n")
}

# Loads a form of a workbook (in the Excel reference, puts in the items RAS puts in by code),
# reads its lists and captures it.
function Show-Form($book, $spec, [string]$label, [bool]$fill, [uint32]$excelPid) {
    $formName = $spec.form.name
    $result = @{ lists = @{}; image = $null; file = $null; error = $null }
    $instance = Invoke-Check $book "RasCheck_Load" $formName
    try {
        foreach ($c in $spec.controls) {
            if ($c.type -notin @("ComboBox", "ListBox")) { continue }
            $ctl = Invoke-Check $book "RasCheck_Find" $instance $c.name
            if ($null -eq $ctl) { $result.lists[$c.name] = "#ERROR コントロールがありません"; continue }
            if ($fill -and $null -ne $c.list) {
                $index = Get-Field $c "selectedIndex"; if ($null -eq $index) { $index = -1 }
                Add-SetupError "$formName.$($c.name).list" (Invoke-Check $book "RasCheck_AddItems" $ctl (Get-ItemsText $c) ([int]$index))
            }
            $result.lists[$c.name] = [string](Invoke-Check $book "RasCheck_ListInfo" $ctl)
        }
        if ($SkipVisual) { return $result }
        $book.Activate()
        Invoke-Check $book "RasCheck_Show" $instance | Out-Null
        Start-Sleep -Milliseconds 700
        $caption = [string](Get-Field $spec.form "caption")
        $hwnd = [RasCapture]::Find($caption, $excelPid)
        if ($hwnd -eq [IntPtr]::Zero) { $result.error = "フォームのウィンドウが見つかりません"; return $result }
        $result.image = [RasCapture]::CaptureClient($hwnd)
        $result.file = Join-Path $Out "images\$formName.$label.png"
        $result.image.Save($result.file, [System.Drawing.Imaging.ImageFormat]::Png)
        return $result
    }
    finally {
        try { Invoke-Check $book "RasCheck_Unload" $instance | Out-Null } catch { }
    }
}

function Close-Excel {
    if ($KeepExcel) { return }
    foreach ($b in @($rasWb, $ref)) { if ($null -ne $b) { try { $b.Close($false) } catch { } } }
    try { $xl.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

try {
    Write-Step "Excel のデザイナーで同じフォームを作る（参照ブック）"
    $refBook = Join-Path $work "ExcelForms.xlsm"
    $ref = $xl.Workbooks.Add()
    $ref.SaveAs($refBook, 52)
    Import-Helpers $ref
    Set-SheetData $ref
    foreach ($spec in $cases) {
        Write-Host "  $($spec.form.name)"
        try { Build-Reference $ref $spec } catch { Add-SetupError $spec.form.name $_.Exception.Message }
    }
    $ref.Save()
    $refCopy = Join-Path $work "ExcelForms-copy.xlsm"
    $ref.SaveCopyAs($refCopy)

    Write-Step "参照ブックを RAS で取り込む（ras form pull）"
    $pullProject = Join-Path $work "pull"
    Invoke-Ras @("init", $pullProject) | Out-Null
    try { $pullLog = Invoke-Ras @("--project", $pullProject, "form", "pull", "--book", $refCopy, "--force") } catch { $pullLog = $_.Exception.Message }
    Write-Host $pullLog
    $pulled = @{}
    foreach ($spec in $cases) {
        $path = Join-Path $pullProject "forms\$($spec.form.name).form.json"
        if (Test-Path $path) { $pulled[$spec.form.name] = Read-Json $path }
    }

    Write-Step "RAS のブックを開く"
    $rasWb = $xl.Workbooks.Open($rasBook)
    Import-Helpers $rasWb
    Set-SheetData $rasWb

    Write-Step "プロパティを比べる"
    foreach ($spec in $cases) {
        $formName = $spec.form.name
        Write-Host "  $formName"
        try {
            $pulledForm = $pulled[$formName]
            $pulledFormSpec = $null; if ($pulledForm) { $pulledFormSpec = $pulledForm.form }
            foreach ($f in (Get-DesignFields $spec.form "UserForm")) {
                $expected = ConvertFrom-Spec $f (Get-Field $spec.form $f) "UserForm"
                $ras = Read-DesignField $rasWb $formName $null $f "UserForm"
                $reference = Read-DesignField $ref $formName $null $f "UserForm"
                $p = Read-PulledField $pulledFormSpec $f "UserForm"
                Add-Row $formName "(form)" $f $expected $ras $reference $p "" $true
            }
            foreach ($c in $spec.controls) {
                $type = $c.type
                $rasControl = Invoke-Check $rasWb "RasCheck_Control" $formName $c.name
                $refControl = Invoke-Check $ref "RasCheck_Control" $formName $c.name
                $pulledControl = $null
                if ($pulledForm) { $pulledControl = @($pulledForm.controls | Where-Object { $_.name -eq $c.name }) | Select-Object -First 1 }
                $auto = (Get-Field $c "autoSize") -eq $true
                foreach ($f in (Get-DesignFields $c $type)) {
                    $expected = ConvertFrom-Spec $f (Get-Field $c $f) $type
                    if ($null -ne $rasControl) { $ras = Read-DesignField $rasWb $formName $rasControl $f $type } else { $ras = "#ERROR コントロールがありません" }
                    if ($null -ne $refControl) { $reference = Read-DesignField $ref $formName $refControl $f $type } else { $reference = "#ERROR コントロールがありません" }
                    $p = Read-PulledField $pulledControl $f $type
                    $note = ""; $counts = $true
                    if ($auto -and $geometry -contains $f) { $note = "AutoSize: 大きさは Excel が決める"; $counts = $false }
                    Add-Row $formName $c.name $f $expected $ras $reference $p $note $counts
                }
            }
        }
        catch {
            [void]$rows.Add([pscustomobject][ordered]@{
                    form = $formName; control = "(form)"; field = "（読めない）"; status = "NG"; problems = $_.Exception.Message; note = ""
                    expected = ""; ras = ""; reference = ""; pulled = ""
                })
        }
    }

    Write-Step "フォームを表示して比べる"
    $excelPid = [RasCapture]::ProcessOf([IntPtr][int64]$xl.Hwnd)
    foreach ($spec in $cases) {
        $formName = $spec.form.name
        Write-Host "  $formName"
        try { $rasShown = Show-Form $rasWb $spec "ras" $false $excelPid } catch { $rasShown = @{ lists = @{}; image = $null; error = $_.Exception.Message } }
        try { $refShown = Show-Form $ref $spec "excel" $true $excelPid } catch { $refShown = @{ lists = @{}; image = $null; error = $_.Exception.Message } }
        # The items RAS puts in by code, and those read from cells (RowSource).
        foreach ($c in $spec.controls) {
            if ($c.type -notin @("ComboBox", "ListBox")) { continue }
            $columns = Get-Field $c "columnCount"; if ($null -eq $columns -or $columns -lt 1) { $columns = 1 }
            $index = Get-Field $c "selectedIndex"; if ($null -eq $index) { $index = -1 }
            $rasList = $rasShown.lists[$c.name]; $refList = $refShown.lists[$c.name]
            if ($null -ne $c.list) { $expected = (ConvertFrom-List $c.list $columns) + "ListIndex=$index" } else { $expected = $refList }
            Add-Row $formName $c.name "list（実行時）" $expected $rasList $refList "―" "" $true
        }
        $entry = [ordered]@{ form = $formName; ras = $null; excel = $null; diff = $null; pixels = $null; status = "未比較"; error = "" }
        if ($rasShown.error) { $entry.error += "RAS: $($rasShown.error) " }
        if ($refShown.error) { $entry.error += "Excel: $($refShown.error)" }
        if ($rasShown.file) { $entry.ras = "images/$formName.ras.png" }
        if ($refShown.file) { $entry.excel = "images/$formName.excel.png" }
        if ($rasShown.image -and $refShown.image) {
            $n = [RasCapture]::Compare($rasShown.image, $refShown.image, (Join-Path $Out "images\$formName.diff.png"))
            if ($n -ge 0) { $entry.diff = "images/$formName.diff.png" }
            $entry.pixels = $n
            if ($n -lt 0) { $entry.status = "NG（大きさが違う）" } elseif ($n -le $Tolerance) { $entry.status = "OK" } else { $entry.status = "NG" }
        }
        if ($rasShown.image) { $rasShown.image.Dispose() }
        if ($refShown.image) { $refShown.image.Dispose() }
        [void]$visuals.Add([pscustomobject]$entry)
    }
}
finally {
    Close-Excel
}

# ============================================================================ report

function Format-Html([string]$text) { return [System.Net.WebUtility]::HtmlEncode($text).Replace("`n", "<br>").Replace("`t", " &#x21E5; ") }

$ng = @($rows | Where-Object { $_.status -eq "NG" }).Count
$warn = @($rows | Where-Object { $_.status -eq "注意" }).Count
$visualNg = @($visuals | Where-Object { $_.status -like "NG*" }).Count
$html = New-Object System.Text.StringBuilder
[void]$html.Append(@"
<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><title>RAS フォームの確認</title>
<style>
body { font-family: "Yu Gothic UI", "Meiryo UI", sans-serif; margin: 16px; color: #222; background: #fff; }
table { border-collapse: collapse; margin: 8px 0 24px; font-size: 13px; }
th, td { border: 1px solid #ccc; padding: 3px 6px; vertical-align: top; }
th { background: #f0f0f0; position: sticky; top: 0; }
tr.NG td { background: #ffe4e4; } tr.注意 td { background: #fff6d6; }
.images { display: flex; gap: 12px; flex-wrap: wrap; align-items: flex-start; }
.images figure { margin: 0; } .images img { border: 1px solid #999; }
.summary { font-size: 15px; }
label { font-size: 13px; } body.only-problems tr.OK { display: none; }
</style></head><body>
<h1>RAS フォームの確認</h1>
<p class="summary">プロパティ: NG <b>$ng</b> 件、注意 $warn 件（全 $($rows.Count) 件）。見た目: NG <b>$visualNg</b> / $($visuals.Count) フォーム。</p>
<p>列: <b>期待値</b> = forms\*.form.json の値、<b>RAS→Excel</b> = ras pack で書いたフォームを Excel が読んだ値、
<b>Excel（参照）</b> = 同じ値を Excel のデザイナーで設定したフォームの値、<b>Excel→RAS</b> = 参照のフォームを ras form pull で取り込んだ値。<br>
NG = RAS の読み書きの違い。注意 = Excel 自身が値を変える（参照も同じ値になる）、参照のフォームを作れなかったなど、RAS の違いとは言えないもの。</p>
<p><label><input type="checkbox" onchange="document.body.classList.toggle('only-problems', this.checked)"> NG と注意だけ表示</label></p>
"@)
if ($setupErrors.Count -gt 0) {
    [void]$html.Append("<h2>参照のフォームを作るときに出たエラー</h2><ul>")
    foreach ($e in $setupErrors) { [void]$html.Append("<li>" + (Format-Html $e) + "</li>") }
    [void]$html.Append("</ul>")
}
[void]$html.Append("<h2>ras form pull の出力</h2><pre>" + (Format-Html $pullLog) + "</pre>")
$imageCaptions = @{ ras = "RAS"; excel = "Excel（参照）"; diff = "違い（赤）" }
foreach ($spec in $cases) {
    $formName = $spec.form.name
    [void]$html.Append("<h2>" + (Format-Html $formName) + "</h2>")
    $v = @($visuals | Where-Object { $_.form -eq $formName }) | Select-Object -First 1
    if ($v) {
        [void]$html.Append("<p>見た目: <b>" + (Format-Html $v.status) + "</b>")
        if ($null -ne $v.pixels -and $v.pixels -ge 0) { [void]$html.Append("（違うピクセル $($v.pixels)）") }
        if ($v.error) { [void]$html.Append(" " + (Format-Html $v.error)) }
        [void]$html.Append("</p><div class='images'>")
        foreach ($k in @("ras", "excel", "diff")) {
            if ($v.$k) { [void]$html.Append("<figure><img src='$($v.$k)'><figcaption>$($imageCaptions[$k])</figcaption></figure>") }
        }
        [void]$html.Append("</div>")
    }
    [void]$html.Append("<table><tr><th>コントロール</th><th>プロパティ</th><th>結果</th><th>期待値</th><th>RAS→Excel</th><th>Excel（参照）</th><th>Excel→RAS</th><th>備考</th></tr>")
    foreach ($r in @($rows | Where-Object { $_.form -eq $formName })) {
        $result = $r.status; if ($r.problems) { $result += "：" + $r.problems }
        [void]$html.Append("<tr class='$($r.status)'><td>$(Format-Html $r.control)</td><td>$(Format-Html $r.field)</td><td>$(Format-Html $result)</td><td>$(Format-Html $r.expected)</td><td>$(Format-Html $r.ras)</td><td>$(Format-Html $r.reference)</td><td>$(Format-Html $r.pulled)</td><td>$(Format-Html $r.note)</td></tr>")
    }
    [void]$html.Append("</table>")
}
[void]$html.Append("</body></html>")
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $Out "report.html"), $html.ToString(), $utf8)
$json = [ordered]@{ properties = @($rows); visuals = @($visuals); setupErrors = @($setupErrors); pull = $pullLog } | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $Out "results.json"), $json, $utf8)

if ($null -ne $previousEncoding) { try { [Console]::OutputEncoding = $previousEncoding } catch { } }

Write-Step "結果"
foreach ($r in @($rows | Where-Object { $_.status -eq "NG" })) {
    Write-Host ("  NG {0}.{1}.{2}: {3}（期待値 {4} / RAS→Excel {5} / Excel {6} / Excel→RAS {7}）" -f $r.form, $r.control, $r.field, $r.problems, $r.expected, $r.ras, $r.reference, $r.pulled) -ForegroundColor Red
}
foreach ($v in @($visuals | Where-Object { $_.status -like "NG*" })) {
    Write-Host ("  見た目 NG {0}: {1} {2}" -f $v.form, $v.pixels, $v.error) -ForegroundColor Red
}
Write-Host "プロパティ: NG $ng 件、注意 $warn 件（全 $($rows.Count) 件）"
Write-Host "見た目: NG $visualNg / $($visuals.Count) フォーム"
Write-Host "報告: $(Join-Path $Out 'report.html')"
if ($ng -gt 0 -or $visualNg -gt 0) { exit 1 }
exit 0
