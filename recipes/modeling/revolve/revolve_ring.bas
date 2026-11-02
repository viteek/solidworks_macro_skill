Attribute VB_Name = "RevolveRingRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. New disposable blank part only. No save/close.
' Revolves x=10..20 mm, y=0..10 mm about sketch Y: annular cylinder.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim manager As SldWorks.SketchManager
    Dim plane As SldWorks.Feature
    Dim sketchFeature As SldWorks.Feature
    Dim axisLine As SldWorks.SketchSegment
    Dim selectData As SldWorks.SelectData
    Dim segments As Variant
    Dim firstSegment As SldWorks.SketchSegment
    Dim revolve As SldWorks.Feature
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim properties As Variant
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
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a NEW disposable blank part."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    Set part = model
    Set manager = model.SketchManager
    If Not manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 4, , "Exit sketch edit mode."
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) = False Then Err.Raise vbObjectError + 5, , "Part already has solids."
    Set plane = model.FirstFeature
    Do While Not plane Is Nothing
        If plane.GetTypeName2 = "ProfileFeature" Or plane.GetTypeName2 = "3DProfileFeature" Then Err.Raise vbObjectError + 6, , "Use a blank part."
        If plane.GetTypeName2 = "RefPlane" Then Exit Do
        Set plane = plane.GetNextFeature
    Loop
    If plane Is Nothing Then Err.Raise vbObjectError + 7, , "Reference plane missing."
    oldAddToDB = manager.AddToDB
    stateCaptured = True
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 8, , "Plane selection failed."
    manager.InsertSketch True
    If manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 9, , "Sketch not opened."
    manager.AddToDB = True
    segments = manager.CreateCornerRectangle(0.01, 0, 0, 0.02, 0.01, 0)
    If IsEmpty(segments) Then Err.Raise vbObjectError + 10, , "Profile failed."
    Set firstSegment = segments(LBound(segments))
    Set sketchFeature = firstSegment.GetSketch
    sketchFeature.Name = "Recipe_RevolveProfile"
    Set axisLine = manager.CreateCenterLine(0, -0.005, 0, 0, 0.015, 0)
    If axisLine Is Nothing Then Err.Raise vbObjectError + 11, , "Axis line failed."
    manager.AddToDB = False
    manager.InsertSketch True
    ' Reacquire axis from the inactive sketch before selecting it outside edit mode.
    Set axisLine = FindConstructionLine(sketchFeature)
    If axisLine Is Nothing Then Err.Raise vbObjectError + 12, , "Axis readback failed."
    model.ClearSelection2 True
    If sketchFeature.Select2(False, 0) = False Then Err.Raise vbObjectError + 13, , "Profile selection failed."
    Set selectData = model.SelectionManager.CreateSelectData
    selectData.Mark = 4
    If axisLine.Select4(True, selectData) = False Then Err.Raise vbObjectError + 14, , "Axis selection failed."
    Set revolve = model.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, _
        swEndCondBlind, swEndCondBlind, 8# * Atn(1#), 0, False, False, 0, 0, 0, 0, 0, True, True, True)
    If revolve Is Nothing Then Err.Raise vbObjectError + 15, , "Revolve failed."
    revolve.Name = "Recipe_Ring"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 16, , "Rebuild failed."
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) Then Err.Raise vbObjectError + 17, , "No solid."
    If UBound(bodies) - LBound(bodies) + 1 <> 1 Then Err.Raise vbObjectError + 18, , "Expected one solid."
    Set body = bodies(LBound(bodies))
    properties = body.GetMassProperties(1#)
    If IsEmpty(properties) Then Err.Raise vbObjectError + 19, , "No mass properties."
    If Abs(CDbl(properties(3)) - 4# * Atn(1#) * (0.02 ^ 2 - 0.01 ^ 2) * 0.01) > 0.00000000001 Then Err.Raise vbObjectError + 20, , "Ring volume mismatch."
    manager.AddToDB = oldAddToDB
    model.ClearSelection2 True
    resultText = "RESULT revolve_ring OK; solid_bodies=1; angle_rad=" & 8# * Atn(1#) & "; volume_m3=" & properties(3)
    GoTo Cleanup
Failed:
    failureText = "FAIL revolve_ring " & Err.Number & ": " & Err.Description
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
        resultText = "FAIL revolve_ring cleanup: " & cleanupText & "; prior=" & Replace(resultText, "RESULT ", "")
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
    failureText = "FAIL revolve_ring result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub

Private Function FindConstructionLine(ByVal feature As SldWorks.Feature) As SldWorks.SketchSegment
    Dim sketch As SldWorks.Sketch
    Dim segments As Variant
    Dim segment As SldWorks.SketchSegment
    Dim index As Long
    Set sketch = feature.GetSpecificFeature2
    segments = sketch.GetSketchSegments
    If IsEmpty(segments) Then Exit Function
    For index = LBound(segments) To UBound(segments)
        Set segment = segments(index)
        If segment.ConstructionGeometry <> False Then
            Set FindConstructionLine = segment
            Exit Function
        End If
    Next index
End Function
