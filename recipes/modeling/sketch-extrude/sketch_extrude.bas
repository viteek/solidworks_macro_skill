Attribute VB_Name = "SketchExtrudeRecipe"
Option Explicit

' Original recipe. NOT_LIVE_TESTED. New disposable blank part only. No save/close.
' All geometric numbers are SI: m, rad, m^3. Result: 40 x 20 x 10 mm block.
Private Const WIDTH_M As Double = 0.04
Private Const HEIGHT_M As Double = 0.02
Private Const DEPTH_M As Double = 0.01
Private Const RESULT_PATH As String = ""

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim manager As SldWorks.SketchManager
    Dim plane As SldWorks.Feature
    Dim sketch As SldWorks.Sketch
    Dim sketchFeature As SldWorks.Feature
    Dim boss As SldWorks.Feature
    Dim segments As Variant
    Dim firstSegment As SldWorks.SketchSegment
    Dim firstLine As SldWorks.SketchLine
    Dim anchor As SldWorks.SketchPoint
    Dim defineStatus As Long
    Dim oldAddToDB As Boolean
    Dim stateCaptured As Boolean
    Dim resultHandle As Integer
    Dim resultOpened As Boolean
    Dim failureText As String
    Dim resultText As String
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
    Set plane = RequireBlankPart(model, part)
    If WIDTH_M <= 0 Or HEIGHT_M <= 0 Or DEPTH_M <= 0 Then Err.Raise vbObjectError + 3, , "Dimensions must be positive meters."
    oldAddToDB = manager.AddToDB
    stateCaptured = True
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 4, , "Plane selection failed."
    manager.InsertSketch True
    Set sketch = manager.ActiveSketch
    If sketch Is Nothing Then Err.Raise vbObjectError + 5, , "Sketch was not opened."
    manager.AddToDB = True
    segments = manager.CreateCornerRectangle(0, 0, 0, WIDTH_M, HEIGHT_M, 0)
    manager.AddToDB = False
    If IsEmpty(segments) Then Err.Raise vbObjectError + 6, , "Rectangle was not created."
    If IsArray(segments) = False Then Err.Raise vbObjectError + 7, , "Expected sketch segments."
    If UBound(segments) - LBound(segments) + 1 <> 4 Then Err.Raise vbObjectError + 8, , "Expected four rectangle edges."
    Set firstSegment = segments(LBound(segments))
    Set sketchFeature = firstSegment.GetSketch
    sketchFeature.Name = "Recipe_Profile"
    Set firstLine = firstSegment
    Set anchor = firstLine.GetStartPoint2
    If anchor Is Nothing Then Err.Raise vbObjectError + 9, , "Missing anchor."
    model.ClearSelection2 True
    If anchor.Select4(False, Nothing) = False Then Err.Raise vbObjectError + 10, , "Anchor selection failed."
    ' Fix only one corner. The remaining geometry is defined by relations/dimensions.
    model.SketchAddConstraints "sgFIXED"
    model.ClearSelection2 True
    defineStatus = manager.FullyDefineSketch(True, True, _
        swSketchFullyDefineRelationType_Horizontal Or swSketchFullyDefineRelationType_Vertical, _
        True, 1, anchor, 1, anchor, 1, 1)
    ' FullyDefineSketch's return value is documented as undefined. Inspect the sketch.
    If sketch.GetConstrainedStatus <> swFullyConstrained Then Err.Raise vbObjectError + 11, , "Sketch is not fully constrained."
    manager.InsertSketch True
    model.ClearSelection2 True
    If sketchFeature.Select2(False, 0) = False Then Err.Raise vbObjectError + 12, , "Profile selection failed."
    Set boss = model.FeatureManager.FeatureExtrusion3(True, False, False, _
        swEndCondBlind, swEndCondBlind, DEPTH_M, 0, False, False, False, False, _
        0, 0, False, False, False, False, True, True, True, swStartSketchPlane, 0, False)
    If boss Is Nothing Then Err.Raise vbObjectError + 13, , "Extrusion failed."
    boss.Name = "Recipe_Block"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 14, , "Rebuild failed."
    AssertVolume part, WIDTH_M * HEIGHT_M * DEPTH_M
    manager.AddToDB = oldAddToDB
    model.ClearSelection2 True
    resultText = "RESULT sketch_extrude OK; fully_constrained=3; solid_bodies=1; volume_m3=" & WIDTH_M * HEIGHT_M * DEPTH_M
    GoTo Cleanup
Failed:
    failureText = "FAIL sketch_extrude " & Err.Number & ": " & Err.Description
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
        resultText = "FAIL sketch_extrude cleanup: " & cleanupText & "; prior=" & Replace(resultText, "RESULT ", "")
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
    failureText = "FAIL sketch_extrude result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub

Private Function RequireBlankPart(ByVal model As SldWorks.ModelDoc2, ByVal part As SldWorks.PartDoc) As SldWorks.Feature
    Dim candidate As SldWorks.Feature
    Dim planeCount As Long
    Dim bodies As Variant
    Dim kind As String
    If model.GetConfigurationCount <> 1 Then Err.Raise vbObjectError + 21, , "Exactly one configuration required."
    If Not model.SketchManager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 22, , "Exit sketch edit mode."
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) = False Then Err.Raise vbObjectError + 23, , "Part already has solids."
    bodies = part.GetBodies2(swSheetBody, False)
    If IsEmpty(bodies) = False Then Err.Raise vbObjectError + 24, , "Part already has surfaces."
    Set candidate = model.FirstFeature
    Do While Not candidate Is Nothing
        kind = candidate.GetTypeName2
        If kind = "ProfileFeature" Or kind = "3DProfileFeature" Then Err.Raise vbObjectError + 25, , "Part already has sketches."
        If kind = "RefPlane" Then
            planeCount = planeCount + 1
            If RequireBlankPart Is Nothing Then Set RequireBlankPart = candidate
        End If
        Set candidate = candidate.GetNextFeature
    Loop
    If planeCount <> 3 Then Err.Raise vbObjectError + 26, , "Expected exactly three reference planes."
End Function

Private Sub AssertVolume(ByVal part As SldWorks.PartDoc, ByVal expectedM3 As Double)
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim properties As Variant
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) Then Err.Raise vbObjectError + 30, , "No solid body."
    If IsArray(bodies) = False Then Err.Raise vbObjectError + 31, , "Expected body array."
    If UBound(bodies) - LBound(bodies) + 1 <> 1 Then Err.Raise vbObjectError + 32, , "Expected one solid."
    Set body = bodies(LBound(bodies))
    properties = body.GetMassProperties(1#)
    If IsEmpty(properties) Then Err.Raise vbObjectError + 33, , "Missing mass properties."
    If Abs(CDbl(properties(3)) - expectedM3) > 0.00000000001 Then Err.Raise vbObjectError + 34, , "Volume assertion failed."
End Sub
