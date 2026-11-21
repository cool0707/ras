Attribute VB_Name = "Probe"
Option Explicit
' What Excel and VBA themselves do, where the stand-in and the debug run over COM differ.
' Run by Check-StandIn.ps1 -Probe; ASCII only (VBComponents.Import reads the system code page).

Private out As String

Private Sub Say(ByVal s As String)
    out = out & s & vbLf
End Sub

Private Function Mode() As String
    Mode = CStr(CLng(Application.CutCopyMode))
End Function

' Whether the copy is still there after an operation: "<operation>=<CutCopyMode>".
Private Sub AfterCopy(ByVal ws As Worksheet, ByVal what As String)
    Dim m As String
    ws.Range("A2").Copy
    On Error Resume Next
    Select Case what
        Case "value": ws.Range("J9").Value = 1
        Case "formula": ws.Range("J10").Formula = "=1+1"
        Case "clearContents": ws.Range("J9").ClearContents
        Case "clear": ws.Range("J9").Clear
        Case "font": ws.Range("J9").Font.Bold = True
        Case "numberFormat": ws.Range("J9").NumberFormat = "0.0"
        Case "interior": ws.Range("J9").Interior.Color = 255
        Case "columnWidth": ws.Columns("J").ColumnWidth = 12
        Case "rowHeight": ws.Rows(9).RowHeight = 20
        Case "hidden": ws.Columns("M").Hidden = True
        Case "insert": ws.Rows(20).Insert
        Case "delete": ws.Rows(20).Delete
        Case "sort": ws.Range("A1:D4").Sort ws.Range("C1"), xlDescending, Header:=xlYes
        Case "autoFilter": ws.Range("A1").AutoFilter 2, "x"
        Case "showAll": ws.AutoFilterMode = False
        Case "merge": ws.Range("L1:M1").Merge
        Case "fillDown": ws.Range("J9:J11").FillDown
        Case "calculate": Application.Calculate
        Case "select": ws.Range("B3").Select
        Case "sheetAdd": ws.Parent.Worksheets.Add After:=ws
        Case "activate": ws.Activate
        Case "rename": ws.Name = "S1"
        Case "nameAdd": ws.Parent.Names.Add "Probe1", "=1"
        Case "pasteSpecial": ws.Range("N2").PasteSpecial xlPasteValues
        Case "paste": ws.Paste ws.Range("N3")
        Case "copyDest": ws.Range("A3").Copy ws.Range("N4")
        Case "borders": ws.Range("J9").Borders.LineStyle = xlContinuous
        Case "calcManual": Application.Calculation = xlCalculationManual: Application.Calculation = xlCalculationAutomatic
        Case "screenUpdating": Application.ScreenUpdating = False: Application.ScreenUpdating = True
        Case "enableEvents": Application.EnableEvents = False: Application.EnableEvents = True
        Case "save": ws.Parent.Save
    End Select
    If Err.Number <> 0 Then m = " err" & Err.Number
    On Error GoTo 0
    Say "copy." & what & "=" & Mode() & m
    Application.CutCopyMode = False
End Sub

Public Function RunProbe() As String
    Dim ws As Worksheet, wb As Workbook, op As Variant
    out = ""
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)
    ws.Activate

    For Each op In Array("value", "formula", "clearContents", "clear", "font", "numberFormat", "interior", "borders", _
            "columnWidth", "rowHeight", "hidden", "insert", "delete", "sort", "autoFilter", "showAll", "merge", "fillDown", _
            "calculate", "select", "activate", "rename", "nameAdd", "pasteSpecial", "paste", "copyDest", _
            "calcManual", "screenUpdating", "enableEvents", "save", "sheetAdd")
        AfterCopy ws, CStr(op)
    Next
    ws.Activate

    ' The selection after a paste.
    ws.Range("A1").Select
    ws.Range("A2:A4").Copy
    ws.Range("P6").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Say "sel.pasteSpecial=" & Selection.Address(False, False)
    ws.Range("A1").Select
    ws.Range("A2:A4").Copy
    ws.Paste ws.Range("P2")
    Say "sel.pasteDest=" & Selection.Address(False, False) & " mode=" & Mode()
    ws.Range("A1").Select
    ws.Range("A2:A4").Cut
    ws.Paste ws.Range("R2")
    Say "sel.cutPasteDest=" & Selection.Address(False, False) & " mode=" & Mode()
    ws.Range("R2:R4").Cut ws.Range("A2")
    ws.Range("B1").Select
    ws.Range("A2:A4").Copy
    ws.Range("T2").Select
    ws.Paste
    Say "sel.paste=" & Selection.Address(False, False) & " mode=" & Mode()
    Application.CutCopyMode = False

    ' Union and Intersect.
    Dim u As Range
    Set u = Union(ws.Range("A1"), ws.Range("C1:C2"), ws.Range("A2"))
    Say "union=" & u.Address(False, False) & " areas=" & u.Areas.Count & " count=" & u.Count & " intersect=" & Intersect(u, ws.Range("A1:B9")).Address(False, False)
    Set u = Union(ws.Range("A1"), ws.Range("C1"), ws.Range("B1"))
    Say "union2=" & u.Address(False, False) & " areas=" & u.Areas.Count
    Set u = Union(ws.Range("A1:B2"), ws.Range("B2:C3"))
    Say "union3=" & u.Address(False, False) & " areas=" & u.Areas.Count & " count=" & u.Count

    ' The theme colour of a border that has none.
    ws.Range("V1:W2").BorderAround xlContinuous, xlThick
    On Error Resume Next
    Dim t As Variant
    t = "unset"
    t = ws.Range("V1").Borders(xlEdgeLeft).ThemeColor
    Say "border.themeColor=" & TypeName(t) & " err=" & Err.Number
    Err.Clear
    t = "unset"
    t = ws.Range("V1").Font.ThemeColor
    Say "font.themeColor=" & TypeName(t) & " " & CStr(t) & " err=" & Err.Number
    Err.Clear
    t = "unset"
    t = ws.Range("V1").Interior.ThemeColor
    Say "interior.themeColor=" & TypeName(t) & " err=" & Err.Number
    Err.Clear
    On Error GoTo 0

    ' Dates and the formats formulas take.
    ws.Range("X1").Value = "2026/3/4"
    Say "date=" & TypeName(ws.Range("X1").Value) & " " & ws.Range("X1").NumberFormat & " local=" & ws.Range("X1").NumberFormatLocal & " text=" & ws.Range("X1").Text
    ws.Range("X2").Value = DateSerial(2026, 1, 31)
    Say "date2=" & ws.Range("X2").NumberFormat & " text=" & ws.Range("X2").Text
    ws.Range("X3").Formula = "=X2+1"
    Say "date3=" & ws.Range("X3").NumberFormat & " text=" & ws.Range("X3").Text
    ws.Range("X4").Formula = "=SUM(C2:C4)"
    Say "inherit.sum=" & ws.Range("X4").NumberFormat & " text=" & ws.Range("X4").Text
    ws.Range("X5").Formula = "=C2*D2"
    Say "inherit.mul=" & ws.Range("X5").NumberFormat & " text=" & ws.Range("X5").Text
    ws.Range("X6").Formula = "=C2+1"
    Say "inherit.add=" & ws.Range("X6").NumberFormat & " text=" & ws.Range("X6").Text
    ws.Range("X7").Formula = "=ROUND(C3*D3*1.1,0)"
    Say "inherit.round=" & ws.Range("X7").NumberFormat & " text=" & ws.Range("X7").Text
    ws.Range("X8").Formula = "=D2+C2"
    Say "inherit.first=" & ws.Range("X8").NumberFormat & " text=" & ws.Range("X8").Text
    ws.Range("X9").Value = "50%"
    Say "percent=" & ws.Range("X9").NumberFormat & " " & CStr(ws.Range("X9").Value2)
    ws.Range("X10").Value = "1,234"
    Say "thousands=" & ws.Range("X10").NumberFormat & " " & CStr(ws.Range("X10").Value2)
    ws.Range("X11").Value = "12:30"
    Say "time=" & ws.Range("X11").NumberFormat & " text=" & ws.Range("X11").Text
    ws.Range("X12").Value = "1e3"
    Say "sci=" & ws.Range("X12").NumberFormat & " " & CStr(ws.Range("X12").Value2)
    ws.Range("X13").Value = "TRUE"
    Say "bool=" & TypeName(ws.Range("X13").Value)
    ws.Range("X14").Value = "001"
    Say "leadingZero=" & TypeName(ws.Range("X14").Value) & " " & ws.Range("X14").Text
    ws.Range("X15").Value = "3/4"
    Say "monthDay=" & TypeName(ws.Range("X15").Value) & " " & ws.Range("X15").NumberFormat & " local=" & ws.Range("X15").NumberFormatLocal
    ws.Range("X16").Value = "\100"
    Say "yen=" & TypeName(ws.Range("X16").Value) & " " & ws.Range("X16").NumberFormat

    ' A sheet added while the VBA editor may not have been opened.
    Dim added As Worksheet
    Set added = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    Say "added.codeName=[" & added.CodeName & "] name=" & added.Name

    ' Scripting objects.
    Dim fso As Object, d As Object, re As Object, f As Object, v As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = CreateObject("Scripting.Dictionary")
    Set f = fso.GetFile(wb.FullName)
    Say "typeName=" & TypeName(f) & " " & TypeName(d) & " " & TypeName(fso) & " " & TypeName(f.ParentFolder)
    v = d("zz")
    Say "missing=" & TypeName(v) & " count=" & d.Count
    Set re = CreateObject("VBScript.RegExp")
    Say "typeName.re=" & TypeName(re)
    Dim p As Variant
    For Each p In Array("(", ")", "[", "a{2", "*a", "a**", "\")
        re.Pattern = p
        On Error Resume Next
        re.Test "x"
        Say "regex " & p & " err=" & Err.Number
        Err.Clear
        On Error GoTo 0
    Next
    RunProbe = out
End Function
