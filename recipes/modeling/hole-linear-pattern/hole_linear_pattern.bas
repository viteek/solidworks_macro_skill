Attribute VB_Name = "HoleLinearPatternRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. Run on untouched disposable sketch-extrude result.
' Creates three diameter 4 mm through holes at x=5,15,25; y=10 mm in sketch space.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim manager As SldWorks.SketchManager
    Dim profileFeature As SldWorks.Feature
    Dim baseSketch As SldWorks.Sketch
    Dim plane As SldWorks.Feature
    Dim circleSegment As SldWorks.SketchSegment
    Dim holeSketchFeature As SldWorks.Feature
    Dim hole As SldWorks.Feature
    Dim direction As SldWorks.SketchSegment
    Dim segment As SldWorks.SketchSegment
    Dim sketchLine As SldWorks.SketchLine
    Dim startPoint As SldWorks.SketchPoint
    Dim endPoint As SldWorks.SketchPoint
    Dim segments As Variant
    Dim index As Long
    Dim selectData As SldWorks.SelectData
    Dim patternData As SldWorks.LinearPatternFeatureData
    Dim pattern As SldWorks.Feature
    Dim reverseDirection As Boolean
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
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open the untouched disposable recipe block."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    If model.GetConfigurationCount <> 1 Then Err.Raise vbObjectError + 4, , "One configuration required."
    Set part = model
    Set manager = model.SketchManager
    If Not manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 5, , "Exit sketch edit mode."
    AssertGeometry part, 0.000008, 0
    Set profileFeature = part.FeatureByName("Recipe_Profile")
    If profileFeature Is Nothing Then Err.Raise vbObjectError + 6, , "Recipe_Profile missing."
    Set baseSketch = profileFeature.GetSpecificFeature2
    segments = baseSketch.GetSketchSegments
    If IsEmpty(segments) Then Err.Raise vbObjectError + 7, , "No base sketch segments."
    For index = LBound(segments) To UBound(segments)
        Set segment = segments(index)
        If segment.GetType = swSketchLINE Then
            Set sketchLine = segment
            Set startPoint = sketchLine.GetStartPoint2
            Set endPoint = sketchLine.GetEndPoint2
            If Abs(endPoint.X - startPoint.X) > 0.039 And Abs(endPoint.Y - startPoint.Y) < 0.000000001 Then
                Set direction = segment
                reverseDirection = (endPoint.X < startPoint.X)
                Exit For
            End If
        End If
    Next index
    If direction Is Nothing Then Err.Raise vbObjectError + 8, , "Fixture X direction not found."
    Set plane = model.FirstFeature
    Do While Not plane Is Nothing
        If plane.GetTypeName2 = "RefPlane" Then Exit Do
        Set plane = plane.GetNextFeature
    Loop
    If plane Is Nothing Then Err.Raise vbObjectError + 9, , "Plane missing."
    oldAddToDB = manager.AddToDB
    stateCaptured = True
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 10, , "Plane selection failed."
    manager.InsertSketch True
    If manager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 11, , "Sketch not opened."
    manager.AddToDB = True
    Set circleSegment = manager.CreateCircleByRadius(0.005, 0.01, 0, 0.002)
    manager.AddToDB = False
    If circleSegment Is Nothing Then Err.Raise vbObjectError + 12, , "Circle not created."
    Set holeSketchFeature = circleSegment.GetSketch
    holeSketchFeature.Name = "Recipe_HoleProfile"
    manager.InsertSketch True
    model.ClearSelection2 True
    If holeSketchFeature.Select2(False, 0) = False Then Err.Raise vbObjectError + 13, , "Hole profile selection failed."
    ' Cut default is opposite the normal; reverse it into the positive-normal block.
    Set hole = model.FeatureManager.FeatureCut4(True, False, True, _
        swEndCondThroughAll, swEndCondBlind, 0.01, 0, False, False, False, False, _
        0, 0, False, False, False, False, False, True, True, False, False, False, _
        swStartSketchPlane, 0, False, False)
    If hole Is Nothing Then Err.Raise vbObjectError + 14, , "Hole cut failed."
    hole.Name = "Recipe_Hole"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 15, , "Hole rebuild failed."
    AssertGeometry part, 0.000008 - 4# * Atn(1#) * 0.002 ^ 2 * 0.01, 1
    model.ClearSelection2 True
    Set selectData = model.SelectionManager.CreateSelectData
    selectData.Mark = 1
    If direction.Select4(False, selectData) = False Then Err.Raise vbObjectError + 16, , "Direction selection failed."
    If hole.Select2(True, 4) = False Then Err.Raise vbObjectError + 17, , "Seed selection failed."
    Set patternData = model.FeatureManager.CreateDefinition(swFmLPattern)
    If patternData Is Nothing Then Err.Raise vbObjectError + 18, , "Pattern definition unavailable."
    patternData.D1EndCondition = 0
    patternData.D1ReverseDirection = reverseDirection
    patternData.D1Spacing = 0.01
    patternData.D1TotalInstances = 3
    patternData.D2EndCondition = 0
    patternData.D2TotalInstances = 1
    patternData.GeometryPattern = False
    patternData.VarySketch = False
    Set pattern = model.FeatureManager.CreateFeature(patternData)
    If pattern Is Nothing Then Err.Raise vbObjectError + 19, , "Linear pattern creation failed."
    pattern.Name = "Recipe_ThreeHoles"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 20, , "Pattern rebuild failed."
    Set patternData = pattern.GetDefinition
    If patternData.D1TotalInstances <> 3 Then Err.Raise vbObjectError + 21, , "Instance count readback failed."
    AssertGeometry part, 0.000008 - 3# * 4# * Atn(1#) * 0.002 ^ 2 * 0.01, 3
    manager.AddToDB = oldAddToDB
    model.ClearSelection2 True
    resultText = "RESULT hole_linear_pattern OK; solid_bodies=1; cylindrical_faces=3; instances=3"
    GoTo Cleanup
Failed:
    failureText = "FAIL hole_linear_pattern " & Err.Number & ": " & Err.Description
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
        resultText = "FAIL hole_linear_pattern cleanup: " & cleanupText & "; prior=" & Replace(resultText, "RESULT ", "")
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
    failureText = "FAIL hole_linear_pattern result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub

Private Sub AssertGeometry(ByVal part As SldWorks.PartDoc, ByVal expectedM3 As Double, ByVal expectedCylinders As Long)
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim properties As Variant
    Dim faces As Variant
    Dim face As SldWorks.Face2
    Dim surface As SldWorks.Surface
    Dim index As Long
    Dim cylinderCount As Long
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) Then Err.Raise vbObjectError + 30, , "No solid."
    If IsArray(bodies) = False Then Err.Raise vbObjectError + 31, , "Expected body array."
    If UBound(bodies) - LBound(bodies) + 1 <> 1 Then Err.Raise vbObjectError + 32, , "Expected one solid."
    Set body = bodies(LBound(bodies))
    properties = body.GetMassProperties(1#)
    If IsEmpty(properties) Then Err.Raise vbObjectError + 33, , "No mass properties."
    If Abs(CDbl(properties(3)) - expectedM3) > 0.00000000001 Then Err.Raise vbObjectError + 34, , "Volume mismatch."
    faces = body.GetFaces
    If IsEmpty(faces) Then Err.Raise vbObjectError + 35, , "No faces."
    For index = LBound(faces) To UBound(faces)
        Set face = faces(index)
        Set surface = face.GetSurface
        If surface.IsCylinder <> False Then cylinderCount = cylinderCount + 1
    Next index
    If cylinderCount <> expectedCylinders Then Err.Raise vbObjectError + 36, , "Cylindrical-face count mismatch."
End Sub
