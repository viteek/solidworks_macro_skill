Attribute VB_Name = "SheetMetalFlatBlankRecipe"
Option Explicit

' Original recipe. See evidence.json for execution scope. Disposable blank part only; no save/close.
' All geometric numbers are SI. Result: 100 x 50 x 2 mm unbent sheet-metal blank.
Private Const WIDTH_M As Double = 0.1
Private Const HEIGHT_M As Double = 0.05
Private Const THICKNESS_M As Double = 0.002
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
    Dim flangeData As SldWorks.BaseFlangeFeatureData
    Dim allowance As SldWorks.CustomBendAllowance
    Dim flatCount As Long
    Dim sheetCount As Long
    Dim cutListCount As Long
    Dim seen As Object
    Dim segments As Variant
    Dim firstSegment As SldWorks.SketchSegment
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
    If WIDTH_M <= 0 Or HEIGHT_M <= 0 Or THICKNESS_M <= 0 Then Err.Raise vbObjectError + 3, , "Dimensions must be positive meters."
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
    manager.InsertSketch True
    model.ClearSelection2 True
    If sketchFeature.Select2(False, 0) = False Then Err.Raise vbObjectError + 12, , "Profile selection failed."
    Set flangeData = model.FeatureManager.CreateDefinition(swFmBaseFlange)
    If flangeData Is Nothing Then Err.Raise vbObjectError + 41, , "Base flange definition unavailable."
    flangeData.UseGaugeTable = False
    flangeData.OverrideDefaultSheetMetalParameters = True
    flangeData.Thickness = THICKNESS_M
    flangeData.BendRadius = THICKNESS_M
    flangeData.ReverseThickness = False
    Set allowance = flangeData.GetCustomBendAllowance
    If allowance Is Nothing Then Err.Raise vbObjectError + 42, , "Custom bend allowance unavailable."
    allowance.KFactor = 0.5
    flangeData.Initialize False, True, allowance, True, swSheetMetalReliefRectangular, True, 0.5, 0, 0
    Set boss = model.FeatureManager.CreateFeature(flangeData)
    If boss Is Nothing Then Err.Raise vbObjectError + 13, , "Extrusion failed."
    boss.Name = "Recipe_FlatBlank"
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 14, , "Rebuild failed."
    AssertVolume part, WIDTH_M * HEIGHT_M * THICKNESS_M
    Set flangeData = boss.GetDefinition
    If Abs(flangeData.Thickness - THICKNESS_M) > 0.000000001 Then Err.Raise vbObjectError + 43, , "Thickness readback mismatch."
    Set seen = CreateObject("Scripting.Dictionary")
    Set plane = model.FirstFeature
    Do While Not plane Is Nothing
        InspectTree plane, seen, sheetCount, flatCount, cutListCount
        Set plane = plane.GetNextFeature
    Loop
    If sheetCount < 1 Or flatCount < 1 Then Err.Raise vbObjectError + 44, , "SheetMetal/FlatPattern features missing."
    manager.AddToDB = oldAddToDB
    model.ClearSelection2 True
    resultText = "RESULT create_flat_blank OK; sheet_metal=1; solid_bodies=1; volume_m3=" & WIDTH_M * HEIGHT_M * THICKNESS_M & "; thickness_m=" & flangeData.Thickness & "; flat_patterns=" & flatCount & "; cut_list_folders=" & cutListCount
    GoTo Cleanup
Failed:
    failureText = "FAIL create_flat_blank " & Err.Number & ": " & Err.Description
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
        resultText = "FAIL create_flat_blank cleanup: " & cleanupText & "; prior=" & Replace(resultText, "RESULT ", "")
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
    failureText = "FAIL create_flat_blank result log: " & Err.Number & ": " & Err.Description
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
    If body.IsSheetMetal = False Then Err.Raise vbObjectError + 45, , "Body is not sheet metal."
    properties = body.GetMassProperties(1#)
    If IsEmpty(properties) Then Err.Raise vbObjectError + 33, , "Missing mass properties."
    If Abs(CDbl(properties(3)) - expectedM3) > 0.00000000001 Then Err.Raise vbObjectError + 34, , "Volume assertion failed."
End Sub

Private Sub InspectTree(ByVal feature As SldWorks.Feature, ByVal seen As Object, ByRef sheetCount As Long, ByRef flatCount As Long, ByRef cutListCount As Long)
    Dim key As String
    Dim kind As String
    Dim child As SldWorks.Feature
    Dim sheetData As SldWorks.SheetMetalFeatureData
    Dim folder As SldWorks.BodyFolder
    kind = feature.GetTypeName2
    key = kind & "|" & feature.Name
    If seen.Exists(key) <> False Then Exit Sub
    seen.Add key, True
    If kind = "SheetMetal" Then
        Set sheetData = feature.GetDefinition
        If sheetData Is Nothing Then Err.Raise vbObjectError + 46, , "SheetMetal definition unavailable."
        If Abs(sheetData.Thickness - THICKNESS_M) > 0.000000001 Then Err.Raise vbObjectError + 47, , "Global sheet thickness mismatch."
        sheetCount = sheetCount + 1
    End If
    If kind = "FlatPattern" Then flatCount = flatCount + 1
    If kind = "CutListFolder" Then
        Set folder = feature.GetSpecificFeature2
        If folder Is Nothing Then Err.Raise vbObjectError + 48, , "Cut-list folder definition missing."
        If folder.GetBodyCount <> 1 Then Err.Raise vbObjectError + 49, , "Expected one cut-list member."
        cutListCount = cutListCount + 1
    End If
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        InspectTree child, seen, sheetCount, flatCount, cutListCount
        Set child = child.GetNextSubFeature
    Loop
End Sub
