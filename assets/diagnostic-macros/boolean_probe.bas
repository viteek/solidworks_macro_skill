Attribute VB_Name = "BooleanProbe"
Option Explicit
Const RESULT_PATH As String = ""

' Deliberately dirties only the disposable active copy; never saves the model.
Sub main()
    Dim swApp As SldWorks.SldWorks, model As SldWorks.ModelDoc2
    Dim values As Variant, expected As Variant, index As Long, raw As Boolean
    Dim handle As Integer, result As String
    On Error GoTo Failed
    If RESULT_PATH = "" Then Err.Raise vbObjectError + 100, , "Specify a fresh result log"
    If Dir(RESULT_PATH) <> "" Then Err.Raise vbObjectError + 101, , "Result already exists"
    values = Array(0, 1, -1)
    expected = Array(False, True, True)
    For index = 0 To 2
        If (values(index) <> False) <> expected(index) Then Err.Raise vbObjectError + 102, , "Nonzero comparison failed"
        If (values(index) = False) = expected(index) Then Err.Raise vbObjectError + 103, , "Zero comparison failed"
    Next index
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 104, , "No test copy open"
    model.SetSaveFlag
    raw = model.GetSaveFlag
    If raw = False Then Err.Raise vbObjectError + 105, , "Expected dirty test copy"
    result = "RESULT=PASS; numeric_cases=0,1,-1; raw_api=" & CStr(CLng(raw)) & _
             "; normalized_api=" & CStr(CLng(raw <> False))
WriteResult:
    On Error GoTo LogFailed
    Debug.Print result
    If RESULT_PATH <> "" Then
        If Dir(RESULT_PATH) = "" Then
            handle = FreeFile
            Open RESULT_PATH For Output As #handle
            Print #handle, result
            Close #handle
        End If
    End If
    Exit Sub
Failed:
    result = "RESULT=FAIL; " & CStr(Err.Number) & "; " & Err.Description
    Resume WriteResult
LogFailed:
    result = "RESULT_LOG_FAILED; " & CStr(Err.Number) & "; " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    Debug.Print result
    On Error GoTo 0
End Sub
