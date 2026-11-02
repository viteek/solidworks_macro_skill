Attribute VB_Name = "PlaneSelectionRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. Selects the first of exactly three RefPlanes.
' Requires an active part outside sketch edit mode. Does not save the document.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim candidate As SldWorks.Feature
    Dim firstPlane As SldWorks.Feature
    Dim planeCount As Long
    Dim resultHandle As Integer
    Dim resultOpened As Boolean
    Dim resultText As String
    Dim failureText As String
    Dim cleanupText As String
    On Error GoTo Failed
    If Len(RESULT_PATH) <> 0 Then
        If Len(Dir$(RESULT_PATH)) <> 0 Then Err.Raise vbObjectError + 90, , "Result path already exists; refusing overwrite."
        resultHandle = FreeFile
        Open RESULT_PATH For Output Access Write As #resultHandle
        resultOpened = True
    End If
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a disposable part."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    If Not model.SketchManager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 3, , "Exit sketch edit mode."
    Set candidate = model.FirstFeature
    Do While Not candidate Is Nothing
        If candidate.GetTypeName2 = "RefPlane" Then
            planeCount = planeCount + 1
            If firstPlane Is Nothing Then Set firstPlane = candidate
        End If
        Set candidate = candidate.GetNextFeature
    Loop
    If planeCount <> 3 Then Err.Raise vbObjectError + 4, , "Expected exactly three reference planes; no selection changed."
    model.ClearSelection2 True
    If firstPlane.Select2(False, 0) = False Then Err.Raise vbObjectError + 5, , "Plane selection failed."
    If model.SelectionManager.GetSelectedObjectCount2(-1) <> 1 Then Err.Raise vbObjectError + 6, , "Wrong selection count."
    If model.SelectionManager.GetSelectedObjectType3(1, -1) <> swSelDATUMPLANES Then Err.Raise vbObjectError + 7, , "Wrong selection type."
    resultText = "RESULT plane_selection OK; selected=" & firstPlane.Name & "; mark=0"
    GoTo Cleanup
Failed:
    failureText = "FAIL plane_selection " & Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Resume leaves the active VBA error handler before protected cleanup executes.
    On Error Resume Next
    Err.Clear
    On Error GoTo 0
    If Len(failureText) <> 0 Then resultText = failureText
    If Len(cleanupText) <> 0 Then
        resultText = "FAIL plane_selection cleanup: " & cleanupText & "; prior=" & Replace(resultText, "RESULT ", "")
    End If
    On Error GoTo LogFailed
    If resultOpened <> False Then
        Print #resultHandle, resultText
        Close #resultHandle
        resultOpened = False
    End If
    Debug.Print resultText
    Exit Sub
LogFailed:
    failureText = "FAIL plane_selection result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub
