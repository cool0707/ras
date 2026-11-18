'@@ decl GetTickCount
#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
'@@ decl GetInputState
#If VBA7 Then
Private Declare PtrSafe Function GetInputState Lib "user32" () As Long
#Else
Private Declare Function GetInputState Lib "user32" () As Long
#End If
'@@ decl RAS_SleepApi
#If VBA7 Then
Private Declare PtrSafe Sub RAS_SleepApi Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub RAS_SleepApi Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If
'@@ decl RAS_CopyMemory
#If VBA7 Then
Private Declare PtrSafe Sub RAS_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub RAS_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If
'@@ decl ras_yieldMask ras_lastYield ras_lastCall ras_inDoEvents RAS_YIELD_MS
Public ras_yieldMask As Long
Public ras_inDoEvents As Boolean
Private ras_lastYield As Double
Private ras_lastCall As Double
Private Const RAS_YIELD_MS As Double = {{YIELD_MS}}#
'@@ decl ras_tickLast ras_tickBase
Private ras_tickLast As Double
Private ras_tickBase As Double
'@@ decl ras_depth ras_silent ras_initDone ras_initBusy
Private ras_depth As Long
Private ras_silent As Boolean
Private ras_initDone As Boolean
Private ras_initBusy As Boolean
'@@ decl ras_errNumber ras_errSource ras_errDescription ras_errLine ras_errFile ras_errPending
Private ras_errNumber As Long
Private ras_errSource As String
Private ras_errDescription As String
Private ras_errLine As Long
Private ras_errFile As String
Private ras_errPending As Boolean
'@@ decl ras_appDepth ras_savScreen ras_savEvents ras_savCalc ras_savAlerts ras_savCursor ras_savInteractive
Private ras_appDepth(0 To 6) As Long
Private ras_savScreen As Boolean
Private ras_savEvents As Boolean
Private ras_savCalc As Long
Private ras_savAlerts As Boolean
Private ras_savCursor As Long
Private ras_savInteractive As Boolean
'@@ proc RAS_Len
' Length of an array; 0 for an unallocated (erased) array instead of an error (spec 13.3).
Public Function RAS_Len(ByRef arr As Variant) As Long
    On Error Resume Next
    RAS_Len = UBound(arr) - LBound(arr) + 1&
End Function
'@@ proc RAS_Len2
' Number of columns of a 2-D array.
Public Function RAS_Len2(ByRef arr As Variant) As Long
    On Error Resume Next
    RAS_Len2 = UBound(arr, 2) - LBound(arr, 2) + 1&
End Function
'@@ proc RAS_IsNullish
Public Function RAS_IsNullish(ByRef v As Variant) As Boolean
    If IsObject(v) Then
        RAS_IsNullish = (v Is Nothing)
    Else
        RAS_IsNullish = IsEmpty(v) Or IsNull(v)
    End If
End Function
'@@ proc RAS_Assign
' Assignment when it is only known at run time whether the value is an object (spec 2.3).
Public Sub RAS_Assign(ByRef target As Variant, ByRef source As Variant)
    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If
End Sub
'@@ proc RAS_FMod
' Remainder for real numbers; the sign follows the dividend (VBA's Mod rounds its operands first).
Public Function RAS_FMod(ByVal a As Double, ByVal b As Double) As Double
    RAS_FMod = a - b * Fix(a / b)
End Function
'@@ proc RAS_Round
' Rounds half away from zero (VBA's Round is banker's rounding, spec 13.2).
Public Function RAS_Round(ByVal x As Double, Optional ByVal digits As Long = 0&) As Double
    Dim f As Variant
    On Error GoTo Fallback
    f = CDec(10 ^ digits)
    RAS_Round = CDbl(Fix(CDec(x) * f + CDec(0.5) * Sgn(x)) / f)
    Exit Function
Fallback:
    RAS_Round = x
End Function
'@@ proc RAS_Shl
Public Function RAS_Shl(ByVal v As Long, ByVal n As Long) As Long
    Dim i As Long
    n = n And 31&
    For i = 1& To n
        ' drop the top bit before doubling so that the multiplication cannot overflow
        If (v And &H40000000) <> 0& Then
            v = ((v And &H3FFFFFFF) * 2&) Or &H80000000
        Else
            v = (v And &H3FFFFFFF) * 2&
        End If
    Next i
    RAS_Shl = v
End Function
'@@ proc RAS_Shr
' Arithmetic shift right.
Public Function RAS_Shr(ByVal v As Long, ByVal n As Long) As Long
    n = n And 31&
    RAS_Shr = CLng(Int(CDbl(v) / (2# ^ n)))
End Function
'@@ proc RAS_NumStr
' Number to string, independent of the locale's decimal separator.
Public Function RAS_NumStr(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If
    RAS_NumStr = s
End Function
'@@ proc RAS_BoolStr
Public Function RAS_BoolStr(ByVal b As Boolean) As String
    If b Then
        RAS_BoolStr = "true"
    Else
        RAS_BoolStr = "false"
    End If
End Function
'@@ proc RAS_Str
' String form of a value whose type is only known at run time.
Public Function RAS_Str(ByRef v As Variant) As String
    If IsObject(v) Then
        RAS_Str = TypeName(v)
        Exit Function
    End If
    Select Case VarType(v)
    Case vbBoolean
        RAS_Str = RAS_BoolStr(v)
    Case vbSingle, vbDouble, vbCurrency, vbDecimal
        RAS_Str = RAS_NumStr(v)
    Case vbEmpty, vbNull
        RAS_Str = vbNullString
    Case Else
        RAS_Str = CStr(v)
    End Select
End Function
'@@ proc RAS_ToFixed
Public Function RAS_ToFixed(ByVal x As Double, ByVal digits As Long) As String
    Dim s As String
    Dim p As Long
    Dim missing As Long
    s = RAS_NumStr(RAS_Round(x, digits))
    If InStr(1&, s, "E", vbBinaryCompare) > 0& Then
        RAS_ToFixed = s
        Exit Function
    End If
    p = InStr(1&, s, ".", vbBinaryCompare)
    If digits <= 0& Then
        If p > 0& Then s = Left$(s, p - 1&)
    Else
        If p = 0& Then
            s = s & "."
            p = Len(s)
        End If
        missing = digits - (Len(s) - p)
        If missing > 0& Then s = s & String$(missing, "0")
    End If
    RAS_ToFixed = s
End Function
'@@ proc RAS_Repeat
Public Function RAS_Repeat(ByVal s As String, ByVal n As Long) As String
    If n <= 0& Or Len(s) = 0& Then
        RAS_Repeat = vbNullString
    ElseIf Len(s) = 1& Then
        RAS_Repeat = String$(n, s)
    Else
        RAS_Repeat = Replace(Space$(n), " ", s, 1&, -1&, vbBinaryCompare)
    End If
End Function
'@@ proc RAS_PadStart
Public Function RAS_PadStart(ByVal s As String, ByVal width As Long, ByVal ch As String) As String
    Dim n As Long
    n = width - Len(s)
    If n <= 0& Or Len(ch) = 0& Then
        RAS_PadStart = s
    Else
        RAS_PadStart = Left$(RAS_Repeat(ch, n \ Len(ch) + 1&), n) & s
    End If
End Function
'@@ proc RAS_PadEnd
Public Function RAS_PadEnd(ByVal s As String, ByVal width As Long, ByVal ch As String) As String
    Dim n As Long
    n = width - Len(s)
    If n <= 0& Or Len(ch) = 0& Then
        RAS_PadEnd = s
    Else
        RAS_PadEnd = s & Left$(RAS_Repeat(ch, n \ Len(ch) + 1&), n)
    End If
End Function
'@@ proc RAS_ClampIndex
' JavaScript-style index: negative values count from the end; the result is within 0..n.
Public Function RAS_ClampIndex(ByVal i As Long, ByVal n As Long) As Long
    If i < 0& Then i = n + i
    If i < 0& Then i = 0&
    If i > n Then i = n
    RAS_ClampIndex = i
End Function
'@@ proc RAS_StrSlice
Public Function RAS_StrSlice(ByVal s As String, ByVal startIndex As Long, Optional ByVal endIndex As Variant) As String
    Dim a As Long
    Dim b As Long
    a = RAS_ClampIndex(startIndex, Len(s))
    If IsMissing(endIndex) Then
        b = Len(s)
    Else
        b = RAS_ClampIndex(CLng(endIndex), Len(s))
    End If
    If b > a Then RAS_StrSlice = Mid$(s, a + 1&, b - a)
End Function
'@@ proc RAS_IndexOf
Public Function RAS_IndexOf(ByRef arr As Variant, ByRef x As Variant, Optional ByVal fromIndex As Long = 0&) As Long
    Dim i As Long
    RAS_IndexOf = -1&
    If fromIndex < 0& Then fromIndex = 0&
    For i = fromIndex To RAS_Len(arr) - 1&
        If IsObject(arr(i)) Then
            If IsObject(x) Then
                If arr(i) Is x Then
                    RAS_IndexOf = i
                    Exit Function
                End If
            End If
        ElseIf Not IsObject(x) Then
            If arr(i) = x Then
                RAS_IndexOf = i
                Exit Function
            End If
        End If
    Next i
End Function
'@@ proc RAS_Join
Public Function RAS_Join(ByRef arr As Variant, ByVal sep As String) As String
    Dim n As Long
    Dim i As Long
    Dim parts() As String
    n = RAS_Len(arr)
    If n = 0& Then Exit Function
    ReDim parts(0 To n - 1&)
    For i = 0& To n - 1&
        parts(i) = RAS_Str(arr(i))
    Next i
    RAS_Join = Join(parts, sep)
End Function
'@@ proc RAS_Swap
Private Sub RAS_Swap(ByRef arr As Variant, ByVal i As Long, ByVal j As Long)
    Dim t As Variant
    If IsObject(arr(i)) Then
        Set t = arr(i)
        Set arr(i) = arr(j)
        Set arr(j) = t
    Else
        t = arr(i)
        arr(i) = arr(j)
        arr(j) = t
    End If
End Sub
'@@ proc RAS_Sort RAS_QSort
' In-place ascending sort (quicksort).
Public Sub RAS_Sort(ByRef arr As Variant)
    If RAS_Len(arr) > 1& Then RAS_QSort arr, 0&, UBound(arr)
End Sub

Private Sub RAS_QSort(ByRef arr As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long
    Dim j As Long
    Dim p As Variant
    i = lo
    j = hi
    p = arr((lo + hi) \ 2&)
    Do While i <= j
        Do While arr(i) < p
            i = i + 1&
        Loop
        Do While arr(j) > p
            j = j - 1&
        Loop
        If i <= j Then
            RAS_Swap arr, i, j
            i = i + 1&
            j = j - 1&
        End If
    Loop
    If lo < j Then RAS_QSort arr, lo, j
    If i < hi Then RAS_QSort arr, i, hi
End Sub
'@@ proc RAS_SortBy RAS_QSortBy
' In-place sort with a comparison closure (negative / zero / positive).
Public Sub RAS_SortBy(ByRef arr As Variant, ByVal cmp As Object)
    If RAS_Len(arr) > 1& Then RAS_QSortBy arr, 0&, UBound(arr), cmp
End Sub

Private Sub RAS_QSortBy(ByRef arr As Variant, ByVal lo As Long, ByVal hi As Long, ByVal cmp As Object)
    Dim i As Long
    Dim j As Long
    Dim p As Variant
    i = lo
    j = hi
    If IsObject(arr((lo + hi) \ 2&)) Then
        Set p = arr((lo + hi) \ 2&)
    Else
        p = arr((lo + hi) \ 2&)
    End If
    Do While i <= j
        Do While cmp.Invoke(arr(i), p) < 0
            i = i + 1&
        Loop
        Do While cmp.Invoke(arr(j), p) > 0
            j = j - 1&
        Loop
        If i <= j Then
            RAS_Swap arr, i, j
            i = i + 1&
            j = j - 1&
        End If
    Loop
    If lo < j Then RAS_QSortBy arr, lo, j, cmp
    If i < hi Then RAS_QSortBy arr, i, hi, cmp
End Sub
'@@ proc RAS_Reverse
Public Sub RAS_Reverse(ByRef arr As Variant)
    Dim i As Long
    Dim n As Long
    n = RAS_Len(arr)
    For i = 0& To n \ 2& - 1&
        RAS_Swap arr, i, n - 1& - i
    Next i
End Sub
'@@ proc RAS_MapGet
' Dictionary lookup that does not add the key when it is missing.
Public Function RAS_MapGet(ByVal d As Object, ByRef k As Variant) As Variant
    If d.Exists(k) Then RAS_MapGet = d.Item(k)
End Function
'@@ proc RAS_MapGetObj
Public Function RAS_MapGetObj(ByVal d As Object, ByRef k As Variant) As Object
    If d.Exists(k) Then Set RAS_MapGetObj = d.Item(k)
End Function
'@@ proc RAS_TickMs
' Milliseconds since start-up as a Double; GetTickCount wraps every 49.7 days (spec 10.4.2).
Public Function RAS_TickMs() As Double
    Dim t As Double
    t = GetTickCount()
    If t < 0# Then t = t + 4294967296#
    If t < ras_tickLast Then ras_tickBase = ras_tickBase + 4294967296#
    ras_tickLast = t
    RAS_TickMs = ras_tickBase + t
End Function
'@@ proc RAS_Yield
' Called every N-th iteration of long loops. Calls DoEvents only when input is pending or
' enough time has passed, and tunes N so that it is itself called about every 10 ms.
Public Sub RAS_Yield()
    Dim now_ As Double
    Dim wasIn As Boolean
    now_ = RAS_TickMs()
    If now_ - ras_lastCall < 5# Then
        If ras_yieldMask < &HFFFFF Then ras_yieldMask = ras_yieldMask * 2& + 1&
    ElseIf now_ - ras_lastCall > 20# Then
        If ras_yieldMask > 0& Then ras_yieldMask = ras_yieldMask \ 2&
    End If
    ras_lastCall = now_
    If GetInputState() <> 0& Or now_ - ras_lastYield >= RAS_YIELD_MS Then
        ras_lastYield = now_
        wasIn = ras_inDoEvents
        ras_inDoEvents = True
        DoEvents
        ras_inDoEvents = wasIn
    End If
End Sub
'@@ proc RAS_YieldNow
Public Sub RAS_YieldNow()
    Dim wasIn As Boolean
    ras_lastYield = RAS_TickMs()
    wasIn = ras_inDoEvents
    ras_inDoEvents = True
    DoEvents
    ras_inDoEvents = wasIn
End Sub
'@@ proc RAS_Sleep
' Waits in short slices so that Excel stays responsive (spec 10.4.1).
Public Sub RAS_Sleep(ByVal ms As Long)
    Dim deadline As Double
    Dim remaining As Double
    deadline = RAS_TickMs() + ms
    Do
        remaining = deadline - RAS_TickMs()
        If remaining <= 0# Then Exit Do
        If remaining > 20# Then remaining = 20#
        RAS_SleepApi CLng(remaining)
        RAS_Yield
    Loop
End Sub
'@@ proc RAS_WaitUntil
' Polls the condition closure until it returns True; False on time-out (negative = no time-out).
Public Function RAS_WaitUntil(ByVal cond As Object, Optional ByVal timeoutMs As Long = -1&) As Boolean
    Dim deadline As Double
    deadline = RAS_TickMs() + timeoutMs
    Do
        If cond.Invoke() Then
            RAS_WaitUntil = True
            Exit Function
        End If
        If timeoutMs >= 0& Then
            If RAS_TickMs() >= deadline Then Exit Function
        End If
        RAS_SleepApi 20&
        RAS_Yield
    Loop
End Function
'@@ proc RAS_AppSet
Private Sub RAS_AppSet(ByVal item As Long, ByVal restore As Boolean)
    ' Setting application properties can fail (e.g. Calculation without an open workbook).
    On Error Resume Next
    Select Case item
    Case 0&
        If restore Then
            Application.ScreenUpdating = ras_savScreen
        Else
            ras_savScreen = Application.ScreenUpdating
            Application.ScreenUpdating = False
        End If
    Case 1&
        If restore Then
            Application.EnableEvents = ras_savEvents
        Else
            ras_savEvents = Application.EnableEvents
            Application.EnableEvents = False
        End If
    Case 2&
        If Application.Workbooks.Count > 0& Then
            If restore Then
                Application.Calculation = ras_savCalc
            Else
                ras_savCalc = Application.Calculation
                Application.Calculation = xlCalculationManual
            End If
        End If
    Case 3&
        If restore Then
            Application.DisplayAlerts = ras_savAlerts
        Else
            ras_savAlerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
        End If
    Case 4&
        If restore Then Application.StatusBar = False
    Case 5&
        If restore Then
            Application.Cursor = ras_savCursor
        Else
            ras_savCursor = Application.Cursor
            Application.Cursor = xlWait
        End If
    Case 6&
        If restore Then
            Application.Interactive = ras_savInteractive
        Else
            ras_savInteractive = Application.Interactive
            Application.Interactive = False
        End If
    End Select
End Sub
'@@ proc RAS_Suspend
' Nested suspend / resume: the original state is saved at depth 0 -> 1 and restored at 1 -> 0 (spec 10.2).
' mask bits: screen=1 events=2 calc=4 alerts=8 statusBar=16 cursor=32 interactive=64
Public Sub RAS_Suspend(ByVal mask As Long)
    Dim i As Long
    For i = 0& To 6&
        If (mask And CLng(2 ^ i)) <> 0& Then
            ras_appDepth(i) = ras_appDepth(i) + 1&
            If ras_appDepth(i) = 1& Then RAS_AppSet i, False
        End If
    Next i
End Sub
'@@ proc RAS_Resume
Public Sub RAS_Resume(ByVal mask As Long)
    Dim i As Long
    For i = 6& To 0& Step -1&
        If (mask And CLng(2 ^ i)) <> 0& Then
            If ras_appDepth(i) > 0& Then
                ras_appDepth(i) = ras_appDepth(i) - 1&
                If ras_appDepth(i) = 0& Then RAS_AppSet i, True
            End If
        End If
    Next i
End Sub
'@@ proc RAS_SuspendAll
Public Sub RAS_SuspendAll()
    RAS_Suspend 63&
End Sub
'@@ proc RAS_ResumeAll
Public Sub RAS_ResumeAll()
    RAS_Resume 63&
End Sub
'@@ proc RAS_ResetAppState
' Restores everything that is still suspended (used after an unhandled error).
Public Sub RAS_ResetAppState()
    Dim i As Long
    For i = 6& To 0& Step -1&
        If ras_appDepth(i) > 0& Then
            ras_appDepth(i) = 0&
            RAS_AppSet i, True
        End If
    Next i
End Sub
'@@ proc RAS_Status
Public Sub RAS_Status(ByVal message As String)
    On Error Resume Next
    If Len(message) = 0& Then
        Application.StatusBar = False
    Else
        Application.StatusBar = message
    End If
End Sub
'@@ proc RAS_EnsureInit
' Runs the project's initialisation exactly once per VBA state (spec 12.3).
Public Sub RAS_EnsureInit()
    If ras_initDone Or ras_initBusy Then Exit Sub
    ras_initBusy = True
    On Error GoTo Failed
{{FIX_EVENTS}}
{{INIT_CALL}}
    ras_initDone = True
    ras_initBusy = False
    Exit Sub
Failed:
    ras_initBusy = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'@@ proc RAS_Enter
' Prologue of every macro entry point (spec 12.4). Returns False when the call must be ignored
' because another macro is running and this one was started from inside DoEvents.
' flags: 1 = reentrant, 2 = silent, 4 = nocancel
Public Function RAS_Enter(ByVal procName As String, Optional ByVal flags As Long = 0&) As Boolean
    If ras_inDoEvents And ras_depth > 0& And (flags And 1&) = 0& Then
        On Error Resume Next
        Application.StatusBar = "{{MSG_BUSY}}"
        Exit Function
    End If
    RAS_EnsureInit
    If ras_depth = 0& Then
        ras_silent = ((flags And 2&) <> 0&)
        ras_errPending = False
        On Error Resume Next
        If (flags And 4&) = 0& Then Application.EnableCancelKey = xlErrorHandler
        On Error GoTo 0
    End If
    ras_depth = ras_depth + 1&
    RAS_Enter = True
End Function
'@@ proc RAS_Leave
Public Sub RAS_Leave()
    If ras_depth > 0& Then ras_depth = ras_depth - 1&
    If ras_depth = 0& Then
        On Error Resume Next
        Application.EnableCancelKey = xlInterrupt
    End If
End Sub
'@@ proc RAS_ErrCapture
' Remembers the error that reached a handler. While an error propagates upwards the first
' (innermost) location is kept.
Public Sub RAS_ErrCapture(ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, ByVal errLine As Long, ByVal rasFile As String)
    If ras_errPending And errNumber = ras_errNumber And errDescription = ras_errDescription Then Exit Sub
    ras_errNumber = errNumber
    ras_errSource = errSource
    ras_errDescription = errDescription
    ras_errLine = errLine
    ras_errFile = rasFile
    ras_errPending = True
End Sub
'@@ proc RAS_Rethrow
Public Sub RAS_Rethrow()
    Err.Raise ras_errNumber, ras_errSource, ras_errDescription
End Sub
'@@ proc RAS_CaughtError
' The error object bound to `catch (e)`.
Public Function RAS_CaughtError() As RAS_Error
    Dim e As RAS_Error
    Set e = New RAS_Error
    e.code = ras_errNumber
    e.source = ras_errSource
    e.description = ras_errDescription
    e.rasFile = ras_errFile
    e.rasLine = ras_errLine
    ras_errPending = False
    Set RAS_CaughtError = e
End Function
'@@ proc RAS_MakeError
Public Function RAS_MakeError(ByVal message As String, Optional ByVal code As Long = vbObjectError + 513) As RAS_Error
    Dim e As RAS_Error
    Set e = New RAS_Error
    e.code = code
    e.description = message
    Set RAS_MakeError = e
End Function
'@@ proc RAS_Throw
' `throw e;` - raises the error again, keeping its original location.
Public Sub RAS_Throw(ByVal e As Object)
    ras_errNumber = e.code
    ras_errSource = e.source
    ras_errDescription = e.description
    ras_errFile = e.rasFile
    ras_errLine = e.rasLine
    ras_errPending = True
    Err.Raise ras_errNumber, ras_errSource, ras_errDescription
End Sub
'@@ proc RAS_Fail
' Top-level error handling of a macro entry. Returns True when the error was handled here
' (outermost entry); False when it must propagate to the calling RAS procedure.
Public Function RAS_Fail(ByVal procName As String, Optional ByVal handler As String = vbNullString) As Boolean
    Dim msg As String
    If ras_depth > 1& Then Exit Function
    RAS_Fail = True
    ras_inDoEvents = False
    ras_errPending = False
    RAS_ResetAppState
    If ras_errNumber = 18& Then
        msg = "{{MSG_CANCELLED}}"
    Else
        msg = ras_errDescription & vbCrLf & vbCrLf & "{{MSG_WHERE}}: " & procName
        If Len(ras_errFile) > 0& Then
            msg = msg & " (" & ras_errFile
            If ras_errLine > 0& Then msg = msg & ":" & CStr(ras_errLine)
            msg = msg & ")"
        End If
        msg = msg & vbCrLf & "{{MSG_NUMBER}}: " & CStr(ras_errNumber)
    End If
    On Error Resume Next
    If Len(handler) > 0& Then
        Application.Run handler, RAS_CaughtError()
    ElseIf Not ras_silent Then
        MsgBox msg, vbExclamation, "RAS"
    End If
End Function
'@@ proc RAS_EventError
' Error raised by an event handler: restore the application state and report.
Public Sub RAS_EventError(ByVal eventName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    ras_inDoEvents = False
    ras_errPending = False
    RAS_ResetAppState
    On Error Resume Next
    MsgBox errDescription & vbCrLf & vbCrLf & "{{MSG_WHERE}}: " & eventName & vbCrLf & "{{MSG_NUMBER}}: " & CStr(errNumber), vbExclamation, "RAS"
End Sub
'@@ proc RAS_Is64
Public Function RAS_Is64() As Boolean
#If Win64 Then
    RAS_Is64 = True
#End If
End Function
'@@ proc RAS_IsVBA7
Public Function RAS_IsVBA7() As Boolean
#If VBA7 Then
    RAS_IsVBA7 = True
#End If
End Function
'@@ proc RAS_Rebase RAS_RebaseCopy
' Turns the value of a Range property into a 0-based 2-D array without copying the elements:
' the lower bounds in the SAFEARRAY descriptor are patched in place (spec 9.2.1).
' A single cell (scalar) becomes a 1 x 1 array.
Public Function RAS_Rebase(ByRef v As Variant) As Variant
    Dim vt As Integer
    Dim cDims As Integer
    Dim fFeatures As Integer
    Dim cLocks As Long
    Dim zero As Long
    Dim boundsOffset As Long
    Dim one(0 To 0, 0 To 0) As Variant
#If VBA7 Then
    Dim psa As LongPtr
#Else
    Dim psa As Long
#End If
    If Not IsArray(v) Then
        one(0, 0) = v
        RAS_Rebase = one
        Exit Function
    End If
    RAS_CopyMemory vt, v, 2
    If (vt And &H6000) = &H2000 Then
        RAS_CopyMemory psa, ByVal VarPtr(v) + 8, LenB(psa)
        If psa <> 0 Then
            RAS_CopyMemory cDims, ByVal psa, 2
            RAS_CopyMemory fFeatures, ByVal psa + 2, 2
            RAS_CopyMemory cLocks, ByVal psa + 8, 4
            ' only patch plain, unlocked, dynamically allocated 2-D arrays (FADF_STATIC = &H2, FADF_FIXEDSIZE = &H10)
            If cDims = 2 And cLocks = 0& And (fFeatures And &H12) = 0 Then
#If Win64 Then
                boundsOffset = 24&
#Else
                boundsOffset = 16&
#End If
                ' rgsabound(i) = { cElements As Long, lLbound As Long }
                RAS_CopyMemory ByVal psa + boundsOffset + 4, zero, 4
                RAS_CopyMemory ByVal psa + boundsOffset + 12, zero, 4
                RAS_Rebase = v
                Exit Function
            End If
        End If
    End If
    RAS_Rebase = RAS_RebaseCopy(v)
End Function

Private Function RAS_RebaseCopy(ByRef v As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim lr As Long
    Dim lc As Long
    Dim out() As Variant
    lr = LBound(v, 1)
    lc = LBound(v, 2)
    ReDim out(0 To UBound(v, 1) - lr, 0 To UBound(v, 2) - lc)
    For r = lr To UBound(v, 1)
        For c = lc To UBound(v, 2)
            out(r - lr, c - lc) = v(r, c)
        Next c
    Next r
    RAS_RebaseCopy = out
End Function
'@@ proc RAS_Read2D
' Range -> 0-based 2-D array in one transfer. A single cell also yields an array (spec 9.2).
Public Function RAS_Read2D(ByVal rng As Object) As Variant
    RAS_Read2D = RAS_Rebase(rng.Value)
End Function
'@@ proc RAS_ToColumn
' First column of a Range as a 0-based 1-D array.
Public Function RAS_ToColumn(ByVal rng As Object) As Variant
    Dim v As Variant
    Dim out() As Variant
    Dim i As Long
    v = RAS_Rebase(rng.Columns(1).Value)
    ReDim out(0 To UBound(v, 1))
    For i = 0& To UBound(v, 1)
        out(i) = v(i, 0)
    Next i
    RAS_ToColumn = out
End Function
'@@ proc RAS_ToRow
' First row of a Range as a 0-based 1-D array.
Public Function RAS_ToRow(ByVal rng As Object) As Variant
    Dim v As Variant
    Dim out() As Variant
    Dim i As Long
    v = RAS_Rebase(rng.Rows(1).Value)
    ReDim out(0 To UBound(v, 2))
    For i = 0& To UBound(v, 2)
        out(i) = v(0, i)
    Next i
    RAS_ToRow = out
End Function
'@@ proc RAS_FromArray
' Writes an array to a Range in one transfer, resizing the Range to the array (spec 9.4).
' A 1-D array is written as one column.
Public Sub RAS_FromArray(ByVal rng As Object, ByRef arr As Variant)
    Dim n As Long
    Dim cols As Long
    Dim i As Long
    Dim col() As Variant
    n = RAS_Len(arr)
    If n = 0& Then Exit Sub
    cols = RAS_Len2(arr)
    If cols > 0& Then
        rng.Resize(n, cols).Value = arr
    Else
        ReDim col(0 To n - 1&, 0 To 0)
        For i = 0& To n - 1&
            col(i, 0) = arr(LBound(arr) + i)
        Next i
        rng.Resize(n, 1).Value = col
    End If
End Sub
'@@ proc RAS_Match
' Application.Match without the run-time error: Empty (null) when nothing matches (spec 13.3).
Public Function RAS_Match(ByRef lookupValue As Variant, ByRef lookupArray As Variant, Optional ByVal matchType As Long = 0&) As Variant
    Dim r As Variant
    r = Application.Match(lookupValue, lookupArray, matchType)
    If Not IsError(r) Then RAS_Match = r
End Function
'@@ class RAS_Error
Public code As Long
Public description As String
Public source As String
Public rasFile As String
Public rasLine As Long

Public Property Get message() As String
    message = description
End Property
