Attribute VB_Name = "CircularPatternRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. Blank disposable part, no save/close.
' Creates a diameter 40 x 10 mm disk, four diameter 4 mm axial through holes.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim manager As SldWorks.SketchManager
    Dim plane As SldWorks.Feature
    Dim circleSegment As SldWorks.SketchSegment
    Dim profile As SldWorks.Feature
    Dim disk As SldWorks.Feature
    Dim hole As SldWorks.Feature
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim edges As Variant
    Dim edge As SldWorks.Edge
    Dim curve As SldWorks.Curve
    Dim circleParameters As Variant
    Dim direction As SldWorks.Entity
    Dim selectData As SldWorks.SelectData
    Dim patternData As SldWorks.CircularPatternFeatureData
    Dim pattern As SldWorks.Feature
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
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a blank disposable part."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    Set part = model
    Set manager = model.SketchManager
    Set plane = RequireBlankPart(model, part)
    oldAddToDB = manager.AddToDB
    stateCaptured = True
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 3, , "Plane selection failed."
    manager.InsertSketch True
    If manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 4, , "Disk sketch failed."
    manager.AddToDB = True
    Set circleSegment = manager.CreateCircleByRadius(0, 0, 0, 0.02)
    manager.AddToDB = False
    If circleSegment Is Nothing Then Err.Raise vbObjectError + 5, , "Disk circle failed."
    Set profile = circleSegment.GetSketch
    profile.Name = "Recipe_DiskProfile"
    manager.InsertSketch True
    model.ClearSelection2 True
    If profile.Select2(False, 0) = False Then Err.Raise vbObjectError + 6, , "Disk profile selection failed."
    Set disk = model.FeatureManager.FeatureExtrusion3(True, False, False, _
        swEndCondBlind, swEndCondBlind, 0.01, 0, False, False, False, False, _
        0, 0, False, False, False, False, True, True, True, swStartSketchPlane, 0, False)
    If disk Is Nothing Then Err.Raise vbObjectError + 7, , "Disk extrusion failed."
    disk.Name = "Recipe_Disk"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 8, , "Disk rebuild failed."
    AssertVolume part, 4# * Atn(1#) * 0.02 ^ 2 * 0.01
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 9, , "Hole plane selection failed."
    manager.InsertSketch True
    If manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 10, , "Hole sketch failed."
    manager.AddToDB = True
    Set circleSegment = manager.CreateCircleByRadius(0.01, 0, 0, 0.002)
    manager.AddToDB = False
    If circleSegment Is Nothing Then Err.Raise vbObjectError + 11, , "Hole circle failed."
    Set profile = circleSegment.GetSketch
    profile.Name = "Recipe_DiskHoleProfile"
    manager.InsertSketch True
    model.ClearSelection2 True
    If profile.Select2(False, 0) = False Then Err.Raise vbObjectError + 12, , "Hole profile selection failed."
    Set hole = model.FeatureManager.FeatureCut4(True, False, True, _
        swEndCondThroughAll, swEndCondBlind, 0.01, 0, False, False, False, False, _
        0, 0, False, False, False, False, False, True, True, False, False, False, _
        swStartSketchPlane, 0, False, False)
    If hole Is Nothing Then Err.Raise vbObjectError + 13, , "Disk hole cut failed."
    hole.Name = "Recipe_DiskHole"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 14, , "Hole rebuild failed."
    AssertVolume part, 4# * Atn(1#) * (0.02 ^ 2 - 0.002 ^ 2) * 0.01
    ' A circular outer edge provides the disk center and normal without localized names.
    bodies = part.GetBodies2(swSolidBody, False)
    Set body = bodies(LBound(bodies))
    edges = body.GetEdges
    If IsEmpty(edges) Then Err.Raise vbObjectError + 15, , "Disk edges missing."
    For index = LBound(edges) To UBound(edges)
        Set edge = edges(index)
        Set curve = edge.GetCurve
        If curve.IsCircle <> False Then
            circleParameters = curve.CircleParams
            If Abs(CDbl(circleParameters(6)) - 0.02) < 0.000000001 Then
                Set direction = edge
                Exit For
            End If
        End If
    Next index
    If direction Is Nothing Then Err.Raise vbObjectError + 16, , "Outer circular direction edge not found."
    model.ClearSelection2 True
    Set selectData = model.SelectionManager.CreateSelectData
    selectData.Mark = 1
    If direction.Select4(False, selectData) = False Then Err.Raise vbObjectError + 17, , "Direction selection failed."
    If hole.Select2(True, 4) = False Then Err.Raise vbObjectError + 18, , "Seed selection failed."
    Set patternData = model.FeatureManager.CreateDefinition(swFmCirPattern)
    If patternData Is Nothing Then Err.Raise vbObjectError + 19, , "Circular definition failed."
    patternData.Direction2 = False
    patternData.EqualSpacing = False
    patternData.Spacing = 2# * Atn(1#)
    patternData.TotalInstances = 4
    patternData.GeometryPattern = False
    patternData.VarySketch = False
    Set pattern = model.FeatureManager.CreateFeature(patternData)
    If pattern Is Nothing Then Err.Raise vbObjectError + 20, , "Circular pattern failed."
    pattern.Name = "Recipe_FourDiskHoles"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 21, , "Pattern rebuild failed."
    Set patternData = pattern.GetDefinition
    If patternData.TotalInstances <> 4 Then Err.Raise vbObjectError + 22, , "Instance count mismatch."
    If Abs(patternData.Spacing - 2# * Atn(1#)) > 0.000000001 Then Err.Raise vbObjectError + 23, , "Angle readback mismatch."
    AssertVolume part, 4# * Atn(1#) * (0.02 ^ 2 - 4# * 0.002 ^ 2) * 0.01
    manager.AddToDB = oldAddToDB
    model.ClearSelection2 True
    resultText = "RESULT circular_pattern OK; solid_bodies=1; instances=4; pitch_rad=" & patternData.Spacing
    GoTo Cleanup
Failed:
    failureText = "FAIL circular_pattern " & Err.Number & ": " & Err.Description
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
        resultText = "FAIL circular_pattern cleanup: " & cleanupText & "; prior=" & resultText
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
    failureText = "FAIL circular_pattern result log: " & Err.Number & ": " & Err.Description
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
