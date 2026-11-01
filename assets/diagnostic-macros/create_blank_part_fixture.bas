Attribute VB_Name = "CreateBlankPartFixture"
Option Explicit

' Fixture generator: creates and saves a NEW blank part, never the active input.
Const OUTPUT_PATH As String = ""
Const RESULT_PATH As String = ""
Const TEMPLATE_PATH As String = ""  ' empty = installed default part template

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim fixture As SldWorks.ModelDoc2
    Dim template As String, title As String, result As String
    Dim errors As Long, warnings As Long
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    If OUTPUT_PATH = "" Then Err.Raise vbObjectError + 100, , "Specify a new fixture path"
    If Dir(OUTPUT_PATH) <> "" Then Err.Raise vbObjectError + 101, , "Fixture output already exists"
    If RESULT_PATH = "" Then Err.Raise vbObjectError + 102, , "Specify a new result log path"
    If Dir(RESULT_PATH) <> "" Then Err.Raise vbObjectError + 103, , "Result log already exists"
    template = TEMPLATE_PATH
    If template = "" Then template = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
    If template = "" Then Err.Raise vbObjectError + 104, , "No default part template configured"
    If Dir(template) = "" Then Err.Raise vbObjectError + 105, , "Part template is unavailable"
    Set fixture = swApp.NewDocument(template, 0, 0, 0)
    If fixture Is Nothing Then Err.Raise vbObjectError + 106, , "Cannot create blank part"
    title = fixture.GetTitle
    If fixture.GetType <> swDocPART Then Err.Raise vbObjectError + 107, , "Template did not create a part"
    If fixture.Extension.SaveAs(OUTPUT_PATH, swSaveAsCurrentVersion, swSaveAsOptions_Silent, _
                               Nothing, errors, warnings) = False Then
        Err.Raise vbObjectError + 108, , "Cannot save fixture copy: " & CStr(errors)
    End If
    If errors <> 0 Or Dir(OUTPUT_PATH) = "" Then Err.Raise vbObjectError + 109, , "Fixture save failed"
    result = "RESULT=PASS; blank_part_created; template=" & template
CleanUp:
    On Error Resume Next
    Err.Clear
    ' Only the document this macro created is closed; the runner owns the input copy.
    If Not fixture Is Nothing Then swApp.CloseDoc fixture.GetTitle
    If Err.Number <> 0 Then
        result = "RESULT=FAIL; fixture cleanup: " & Err.Description
        Err.Clear
    End If
    On Error GoTo LogFailed
    Dim handle As Integer
    handle = FreeFile
    If RESULT_PATH <> "" Then
        If Dir(RESULT_PATH) = "" Then
            Open RESULT_PATH For Output As #handle
            Print #handle, result
            Close #handle
        End If
    End If
    Debug.Print result
    Exit Sub
Failed:
    result = "RESULT=FAIL; " & CStr(Err.Number) & "; " & Err.Description
    Resume CleanUp
LogFailed:
    result = "RESULT_LOG_FAILED; " & CStr(Err.Number) & "; " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    Debug.Print result
    On Error GoTo 0
End Sub
