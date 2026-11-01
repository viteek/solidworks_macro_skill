Attribute VB_Name = "SaveFixtureCopy"
Option Explicit

' Observer for fixture preparation only: publish a NEW native copy of the test model.
Const EXPECTED_INPUT_PATH As String = ""
Const OUTPUT_PATH As String = ""
Const RESULT_PATH As String = ""

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim errors As Long, warnings As Long, handle As Integer
    Dim result As String
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 100, , "No active fixture"
    If EXPECTED_INPUT_PATH = "" Then Err.Raise vbObjectError + 101, , "Specify expected test input"
    If StrComp(model.GetPathName, EXPECTED_INPUT_PATH, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 102, , "Active model is not the reviewed test copy"
    End If
    If OUTPUT_PATH = "" Or RESULT_PATH = "" Then Err.Raise vbObjectError + 103, , "Specify new output and log"
    If Dir(OUTPUT_PATH) <> "" Or Dir(RESULT_PATH) <> "" Then Err.Raise vbObjectError + 104, , "Output already exists"
    model.ClearSelection2 True
    If model.Extension.SaveAs(OUTPUT_PATH, swSaveAsCurrentVersion, _
        swSaveAsOptions_Silent + swSaveAsOptions_Copy, Nothing, errors, warnings) = False Then
        Err.Raise vbObjectError + 105, , "Cannot save fixture copy: " & CStr(errors)
    End If
    If errors <> 0 Or Dir(OUTPUT_PATH) = "" Then Err.Raise vbObjectError + 106, , "Fixture copy missing"
    If FileLen(OUTPUT_PATH) = 0 Then Err.Raise vbObjectError + 107, , "Fixture copy empty"
    result = "RESULT=PASS; fixture_copy=" & OUTPUT_PATH
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
