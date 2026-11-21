Attribute VB_Name = "RasCheck"
' Helpers for Check-Forms.ps1 (tools/excel-check): read and set the properties of user forms
' at design time (through the VBA editor's designer) and show them. The script calls them
' with Application.Run; nothing here is part of RAS.
'
' Needs "Trust access to the VBA project object model" (Excel Options > Trust Center).
' ASCII only: VBComponents.Import reads the file in the system code page.
Option Explicit

Private Const NULL_MARKER As String = "<Null>"

' ------------------------------------------------------------------ finding things

' The VBComponent of a form of this workbook.
Public Function RasCheck_Component(ByVal formName As String) As Object
    Set RasCheck_Component = ThisWorkbook.VBProject.VBComponents(formName)
End Function

' The design-time form (MSForms.UserForm) of a form of this workbook.
Public Function RasCheck_Design(ByVal formName As String) As Object
    Set RasCheck_Design = ThisWorkbook.VBProject.VBComponents(formName).Designer
End Function

' A control by name under root (a form, Frame or Page), including the pages of MultiPages.
Public Function RasCheck_Find(ByVal root As Object, ByVal name As String) As Object
    Dim c As Object, p As Object
    On Error Resume Next
    Set RasCheck_Find = root.Controls(name)
    On Error GoTo 0
    If Not RasCheck_Find Is Nothing Then Exit Function
    For Each c In root.Controls
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                If StrComp(p.Name, name, vbTextCompare) = 0 Then
                    Set RasCheck_Find = p
                    Exit Function
                End If
            Next
        End If
    Next
End Function

' A control of a form at design time.
Public Function RasCheck_Control(ByVal formName As String, ByVal controlName As String) As Object
    Set RasCheck_Control = RasCheck_Find(RasCheck_Design(formName), controlName)
End Function

' ------------------------------------------------------------------ reading

' A property; Null stays Null, an error is returned as "#ERROR <number>: <description>".
Public Function RasCheck_Get(ByVal target As Object, ByVal name As String) As Variant
    On Error GoTo Failed
    RasCheck_Get = CallByName(target, name, VbGet)
    Exit Function
Failed:
    RasCheck_Get = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' An object property (Font, Picture ...).
Public Function RasCheck_GetObject(ByVal target As Object, ByVal name As String) As Object
    On Error Resume Next
    Set RasCheck_GetObject = CallByName(target, name, VbGet)
End Function

' Properties of a form the VBA editor keeps in its designer block (VBComponent.Properties), not on
' the form: Width / Height are the outer size, title bar and borders included.
Private Function IsExtenderProperty(ByVal name As String) As Boolean
    Select Case LCase$(name)
        Case "startupposition", "showmodal", "tag", "helpcontextid", "righttoleft", _
             "whatsthisbutton", "whatsthishelp", "caption", "enabled", "width", "height", "left", "top"
            IsExtenderProperty = True
    End Select
End Function

' True / False whatever the language of Office.
Private Function B(ByVal v As Variant) As String
    If IsNull(v) Then
        B = "Null"
    ElseIf v Then
        B = "True"
    Else
        B = "False"
    End If
End Function

' A number with a decimal point whatever the regional settings.
Private Function N(ByVal v As Variant) As String
    N = Trim$(Str$(v))
End Function

' A property of a form: of its designer block where the VBA editor keeps it, else of the form.
Public Function RasCheck_FormGet(ByVal formName As String, ByVal name As String) As Variant
    On Error GoTo Failed
    If IsExtenderProperty(name) Then
        RasCheck_FormGet = RasCheck_Component(formName).Properties(name).Value
    Else
        RasCheck_FormGet = CallByName(RasCheck_Design(formName), name, VbGet)
    End If
    Exit Function
Failed:
    RasCheck_FormGet = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' "" without a picture, else "<Type>:<Width>x<Height>" (HIMETRIC).
Public Function RasCheck_PictureInfo(ByVal target As Object, ByVal name As String) As String
    Dim p As Object
    On Error GoTo Failed
    Set p = CallByName(target, name, VbGet)
    If p Is Nothing Then Exit Function
    If p.Handle = 0 Then Exit Function
    RasCheck_PictureInfo = p.Type & ":" & N(p.Width) & "x" & N(p.Height)
    Exit Function
Failed:
    If Err.Number <> 0 Then RasCheck_PictureInfo = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' The font as "name|size|bold|italic|underline|strikethrough|charset".
Public Function RasCheck_FontInfo(ByVal target As Object) As String
    Dim f As Object
    On Error GoTo Failed
    Set f = target.Font
    RasCheck_FontInfo = f.Name & "|" & N(f.Size) & "|" & B(f.Bold) & "|" & B(f.Italic) & "|" & _
        B(f.Underline) & "|" & B(f.Strikethrough) & "|" & N(f.Charset)
    Exit Function
Failed:
    RasCheck_FontInfo = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' The tabs of a TabStrip, one per line: name, caption, tip, tag, accelerator, enabled, visible
' separated by tabs.
Public Function RasCheck_TabsInfo(ByVal target As Object) As String
    Dim t As Object, s As String
    On Error GoTo Failed
    For Each t In target.Tabs
        If Len(s) > 0 Then s = s & vbLf
        s = s & t.Name & vbTab & t.Caption & vbTab & t.ControlTipText & vbTab & t.Tag & vbTab & _
            t.Accelerator & vbTab & B(t.Enabled) & vbTab & B(t.Visible)
    Next
    RasCheck_TabsInfo = s
    Exit Function
Failed:
    RasCheck_TabsInfo = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' The items of a ComboBox / ListBox, one row per line with the columns separated by tabs,
' then a line "ListIndex=<n>".
Public Function RasCheck_ListInfo(ByVal target As Object) As String
    Dim r As Long, c As Long, s As String, columns As Long
    On Error GoTo Failed
    columns = target.ColumnCount
    If columns < 1 Then columns = 1
    For r = 0 To target.ListCount - 1
        For c = 0 To columns - 1
            If c > 0 Then s = s & vbTab
            s = s & CStr(Nz(target.List(r, c)))
        Next
        s = s & vbLf
    Next
    RasCheck_ListInfo = s & "ListIndex=" & target.ListIndex
    Exit Function
Failed:
    RasCheck_ListInfo = "#ERROR " & Err.Number & ": " & Err.Description
End Function

Private Function Nz(ByVal v As Variant) As Variant
    If IsNull(v) Then Nz = "" Else Nz = v
End Function

' ------------------------------------------------------------------ setting (the Excel reference)

' Sets a property; "<Null>" sets Null. Returns "" or the error.
Public Function RasCheck_Let(ByVal target As Object, ByVal name As String, ByVal value As Variant) As String
    On Error GoTo Failed
    If VarType(value) = vbString Then
        If value = NULL_MARKER Then value = Null
    End If
    CallByName target, name, VbLet, value
    Exit Function
Failed:
    RasCheck_Let = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' Sets a property of a form: in its designer block where the VBA editor keeps it, else on the form.
Public Function RasCheck_FormLet(ByVal formName As String, ByVal name As String, ByVal value As Variant) As String
    On Error GoTo Failed
    If IsExtenderProperty(name) Then
        RasCheck_Component(formName).Properties(name).Value = value
    Else
        CallByName RasCheck_Design(formName), name, VbLet, value
    End If
    Exit Function
Failed:
    RasCheck_FormLet = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' Adds a form named formName. Returns "" or the error.
Public Function RasCheck_NewForm(ByVal formName As String) As String
    Dim comp As Object
    On Error GoTo Failed
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)
    comp.Name = formName
    Exit Function
Failed:
    RasCheck_NewForm = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' Adds a control (a Page with progId "Forms.Page.1") to the form or to parentName (Frame, Page,
' MultiPage). A new MultiPage or TabStrip loses the two pages / tabs it comes with.
Public Function RasCheck_AddControl(ByVal formName As String, ByVal parentName As String, _
        ByVal progId As String, ByVal controlName As String) As Object
    Dim root As Object, parent As Object, c As Object
    Set root = RasCheck_Design(formName)
    If Len(parentName) = 0 Then
        Set parent = root
    Else
        Set parent = RasCheck_Find(root, parentName)
    End If
    If progId = "Forms.Page.1" Then
        Set c = parent.Pages.Add(controlName)
    Else
        Set c = parent.Controls.Add(progId, controlName, True)
        If TypeName(c) = "MultiPage" Then c.Pages.Clear
        If TypeName(c) = "TabStrip" Then c.Tabs.Clear
    End If
    Set RasCheck_AddControl = c
End Function

' Adds a tab to a TabStrip.
Public Function RasCheck_AddTab(ByVal target As Object, ByVal name As String, ByVal caption As String) As Object
    Set RasCheck_AddTab = target.Tabs.Add(name, caption)
End Function

' Sets a picture property from a file. Returns "" or the error.
Public Function RasCheck_SetPicture(ByVal target As Object, ByVal name As String, ByVal path As String) As String
    Dim p As Object
    On Error GoTo Failed
    Set p = LoadPicture(path)
    On Error Resume Next
    CallByName target, name, VbSet, p
    If Err.Number <> 0 Then
        Err.Clear
        CallByName target, name, VbLet, p
    End If
    If Err.Number <> 0 Then GoTo Failed
    Exit Function
Failed:
    RasCheck_SetPicture = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' Sets a picture property of a form from a file.
Public Function RasCheck_FormSetPicture(ByVal formName As String, ByVal name As String, ByVal path As String) As String
    RasCheck_FormSetPicture = RasCheck_SetPicture(RasCheck_Design(formName), name, path)
End Function

' Fills a ComboBox / ListBox at run time: rows separated by line feeds, columns by tabs.
Public Function RasCheck_AddItems(ByVal target As Object, ByVal text As String, ByVal listIndex As Long) As String
    Dim rows() As String, cols() As String, r As Long, c As Long
    On Error GoTo Failed
    If Len(text) > 0 Then
        rows = Split(text, vbLf)
        For r = 0 To UBound(rows)
            cols = Split(rows(r), vbTab)
            target.AddItem cols(0)
            For c = 1 To UBound(cols)
                target.List(target.ListCount - 1, c) = cols(c)
            Next
        Next
    End If
    If listIndex >= 0 Then target.ListIndex = listIndex
    Exit Function
Failed:
    RasCheck_AddItems = "#ERROR " & Err.Number & ": " & Err.Description
End Function

' ------------------------------------------------------------------ run time

' A loaded instance of a form (its Initialize has run).
Public Function RasCheck_Load(ByVal formName As String) As Object
    Set RasCheck_Load = VBA.UserForms.Add(formName)
End Function

Public Sub RasCheck_Show(ByVal form As Object)
    form.Show vbModeless
End Sub

Public Sub RasCheck_Unload(ByVal form As Object)
    Unload form
End Sub
