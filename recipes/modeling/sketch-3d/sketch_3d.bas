Attribute VB_Name = "Sketch3DRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. Creates a named 3D polyline, no save/close.
' Input: a new disposable part. Lengths 10,20,30 mm; absolute model coordinates.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim manager As SldWorks.SketchManager
    Dim sketch As SldWorks.Sketch
    Dim feature As SldWorks.Feature
    Dim segment As SldWorks.SketchSegment
    Dim segments As Variant
    Dim bodies As Variant
    Dim totalLengthM As Double
    Dim index As Long
    Dim oldAddToDB As Boolean
    Dim stateCaptured As Boolean
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
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a NEW disposable part."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    Set manager = model.SketchManager
    If Not manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 4, , "Exit sketch edit mode."
    Set part = model
    Set feature = part.FeatureByName("Recipe_3DPath")
    If Not feature Is Nothing Then Err.Raise vbObjectError + 5, , "Recipe path already exists."
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) = False Then Err.Raise vbObjectError + 6, , "Use a blank part without solids."
    oldAddToDB = manager.AddToDB
    stateCaptured = True
    model.ClearSelection2 True
    manager.Insert3DSketch True
    Set sketch = manager.ActiveSketch
    If sketch Is Nothing Then Err.Raise vbObjectError + 7, , "3D sketch not opened."
    If sketch.Is3D = False Then Err.Raise vbObjectError + 8, , "Expected 3D sketch."
    manager.AddToDB = True
    Set segment = manager.CreateLine(0, 0, 0, 0.01, 0, 0)
    If segment Is Nothing Then Err.Raise vbObjectError + 9, , "First line failed."
    Set feature = segment.GetSketch
    feature.Name = "Recipe_3DPath"
    Set segment = manager.CreateLine(0.01, 0, 0, 0.01, 0.02, 0)
    If segment Is Nothing Then Err.Raise vbObjectError + 10, , "Second line failed."
    Set segment = manager.CreateLine(0.01, 0.02, 0, 0.01, 0.02, 0.03)
    If segment Is Nothing Then Err.Raise vbObjectError + 11, , "Third line failed."
    manager.AddToDB = False
    manager.Insert3DSketch True
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 12, , "Rebuild failed."
    ' Reacquire the feature's specific sketch after leaving edit mode.
    Set sketch = feature.GetSpecificFeature2
    If sketch.Is3D = False Then Err.Raise vbObjectError + 13, , "Readback is not 3D."
    segments = sketch.GetSketchSegments
    If IsEmpty(segments) Then Err.Raise vbObjectError + 14, , "No segments."
    If UBound(segments) - LBound(segments) + 1 <> 3 Then Err.Raise vbObjectError + 15, , "Expected three lines."
    For index = LBound(segments) To UBound(segments)
        Set segment = segments(index)
        If segment.GetType <> swSketchLINE Then Err.Raise vbObjectError + 16, , "Unexpected segment type."
        totalLengthM = totalLengthM + segment.GetLength
    Next index
    If Abs(totalLengthM - 0.06) > 0.000000001 Then Err.Raise vbObjectError + 17, , "Polyline length mismatch."
    manager.AddToDB = oldAddToDB
    model.ClearSelection2 True
    resultText = "RESULT sketch_3d OK; is3d=1; segments=3; length_m=" & totalLengthM
    GoTo Cleanup
Failed:
    failureText = "FAIL sketch_3d " & Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Resume leaves the active VBA error handler before protected cleanup executes.
    On Error Resume Next
    Err.Clear
    If stateCaptured <> False Then
        manager.AddToDB = oldAddToDB
        If Err.Number <> 0 Then cleanupText = "AddToDB restore: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If Len(failureText) <> 0 Then resultText = failureText
    If Len(cleanupText) <> 0 Then
        resultText = "FAIL sketch_3d cleanup: " & cleanupText & "; prior=" & resultText
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
    failureText = "FAIL sketch_3d result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub
