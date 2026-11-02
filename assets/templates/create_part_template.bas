Attribute VB_Name = "Main"
' ============================================================================
' Macro:    create part (parametric plate) - SOLIDWORKS 2024 SP5 verified (Polish UI)
' Builds:   rectangle LENGTH x WIDTH on the front plane, extruded THICKNESS, HOLE_COUNT_X x HOLE_COUNT_Y
'           through holes at HOLE_EDGE_OFFSET from the edges, corner fillets, material, properties,
'           then saves to OUTPUT_PATH and checks volume and size against the inputs.
' Pattern:  new doc from template -> plane by TYPE (not localized name) -> sketch with AddToDB ->
'           named driving dimensions (no 'Modify' dialog) -> feature -> check Nothing -> next feature.
' All sizes in mm in the Consts; the API gets metres.
' ============================================================================
Option Explicit

' ---- Configuration ----------------------------------------------------------
Const OUTPUT_PATH As String = ""            ' full path of the new .SLDPRT (folder must exist); "" = ask
Const ALLOW_OVERWRITE As Boolean = False    ' never replace an existing file unless True
Const TEMPLATE_PATH As String = ""          ' "" = default part template from Tools > Options
Const LENGTH_MM As Double = 200
Const WIDTH_MM As Double = 120
Const THICKNESS_MM As Double = 10
Const HOLE_DIAMETER_MM As Double = 9        ' 0 = no holes
Const HOLE_EDGE_OFFSET_MM As Double = 15    ' hole centre distance from the outer edges
Const HOLE_COUNT_X As Long = 2              ' holes along the length (>= 2 when holes are used)
Const HOLE_COUNT_Y As Long = 2              ' holes along the width
Const CORNER_RADIUS_MM As Double = 10       ' 0 = sharp corners
Const MATERIAL_DB As String = "SOLIDWORKS Materials" ' material database (file name without .sldmat)
Const MATERIAL_NAME As String = ""          ' exact name as in the database, e.g. "AISI 304"; "" = keep template material
Const PROP_DESCRIPTION As String = ""       ' written to file property "Opis" when not empty
Const SHOW_MESSAGES As Boolean = True       ' False = unattended (result only in RESULT_PATH / Immediate)
Const RESULT_PATH As String = ""            ' optional log file: "CREATED;<path>;..." or "ERROR;..."

Dim swApp As SldWorks.SldWorks

Sub main()
    Dim model As SldWorks.ModelDoc2, part As SldWorks.PartDoc, feat As SldWorks.Feature
    Dim plane As SldWorks.Feature, sketchFeat As SldWorks.Feature, topFace As SldWorks.Entity
    Dim oldInputDim As Boolean, toggleChanged As Boolean, outPath As String, status As String, message As String
    Dim e As Long, w As Long, db As String, i As Long, j As Long, x As Double, y As Double
    Dim mp As SldWorks.MassProperty, mp2 As SldWorks.MassProperty2, expectedVolume As Double
    On Error GoTo Failed
    Set swApp = Application.SldWorks

    ' ---- inputs ----
    If LENGTH_MM <= 0 Or WIDTH_MM <= 0 Or THICKNESS_MM <= 0 Then Err.Raise vbObjectError + 1, , "Sizes must be > 0"
    If HOLE_DIAMETER_MM > 0 And (HOLE_COUNT_X < 2 Or HOLE_COUNT_Y < 1) Then Err.Raise vbObjectError + 2, , "Hole counts invalid"
    If HOLE_DIAMETER_MM > 0 And 2 * HOLE_EDGE_OFFSET_MM + HOLE_DIAMETER_MM >= WIDTH_MM Then Err.Raise vbObjectError + 3, , "Holes don't fit"
    If 2 * CORNER_RADIUS_MM >= WIDTH_MM Then Err.Raise vbObjectError + 4, , "Corner radius too large"
    outPath = OUTPUT_PATH
    If outPath = "" And SHOW_MESSAGES Then outPath = InputBox("Save the new part as (full path .SLDPRT):", "Create part")
    If outPath = "" Then Err.Raise vbObjectError + 5, , "Set OUTPUT_PATH"
    If LCase(Right(outPath, 7)) <> ".sldprt" Then outPath = outPath & ".SLDPRT"
    If Dir(outPath) <> "" And ALLOW_OVERWRITE = False Then Err.Raise vbObjectError + 6, , "File exists (ALLOW_OVERWRITE=False): " & outPath

    ' ---- new document from template ----
    Dim tpl As String
    tpl = TEMPLATE_PATH
    If tpl = "" Then tpl = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
    If Dir(tpl) = "" Then Err.Raise vbObjectError + 10, , "Part template not found (PDM not local?): " & tpl
    Set model = swApp.NewDocument(tpl, 0, 0, 0)
    If model Is Nothing Then Err.Raise vbObjectError + 11, , "NewDocument failed for " & tpl
    Set part = model

    ' AddDimension2 opens a modal 'Modify' box when this user option is on (measured) - switch it off
    oldInputDim = swApp.GetUserPreferenceToggle(swInputDimValOnCreate)
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, False
    toggleChanged = True

    ' ---- base sketch on the FRONT plane (first RefPlane; names are localized) ----
    Set plane = StandardPlane(model, 0)
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 12, , "Cannot select front plane"
    model.SketchManager.InsertSketch True
    model.SketchManager.AddToDB = True                     ' otherwise off-screen geometry may silently fail
    model.SketchManager.CreateCornerRectangle 0, 0, 0, LENGTH_MM / 1000#, WIDTH_MM / 1000#, 0
    model.SketchManager.AddToDB = False
    model.SketchManager.InsertSketch True                  ' close sketch
    Set sketchFeat = model.FeatureByPositionReverse(0)
    If sketchFeat.GetTypeName2 <> "ProfileFeature" Then Err.Raise vbObjectError + 13, , "Base sketch not created"

    ' ---- extrude (sketch must be selected) ----
    model.ClearSelection2 True
    sketchFeat.Select2 False, 0
    Set feat = model.FeatureManager.FeatureExtrusion3(True, False, False, swEndCondBlind, swEndCondBlind, _
        THICKNESS_MM / 1000#, 0, False, False, False, False, 0, 0, False, False, False, False, _
        True, True, True, swStartSketchPlane, 0, False)
    If feat Is Nothing Then Err.Raise vbObjectError + 14, , "Extrusion failed"
    expectedVolume = LENGTH_MM * WIDTH_MM * THICKNESS_MM

    ' ---- holes: sketch on the top face, cut through all ----
    If HOLE_DIAMETER_MM > 0 Then
        Set topFace = PlanarFaceWithNormal(part, 0, 0, 1)
        If topFace Is Nothing Then Err.Raise vbObjectError + 15, , "Top face not found"
        model.ClearSelection2 True
        topFace.Select4 False, Nothing
        model.SketchManager.InsertSketch True
        model.SketchManager.AddToDB = True
        For i = 0 To HOLE_COUNT_X - 1
            For j = 0 To HOLE_COUNT_Y - 1
                x = HOLE_EDGE_OFFSET_MM + i * (LENGTH_MM - 2 * HOLE_EDGE_OFFSET_MM) / (HOLE_COUNT_X - 1)
                y = HOLE_EDGE_OFFSET_MM
                If HOLE_COUNT_Y > 1 Then y = HOLE_EDGE_OFFSET_MM + j * (WIDTH_MM - 2 * HOLE_EDGE_OFFSET_MM) / (HOLE_COUNT_Y - 1)
                ' sketch on a face parallel to the front plane: sketch X/Y = model X/Y here (verify for other faces)
                model.SketchManager.CreateCircleByRadius x / 1000#, y / 1000#, 0, HOLE_DIAMETER_MM / 2000#
            Next j
        Next i
        model.SketchManager.AddToDB = False
        model.SketchManager.InsertSketch True
        Set sketchFeat = model.FeatureByPositionReverse(0)
        model.ClearSelection2 True
        sketchFeat.Select2 False, 0
        Set feat = model.FeatureManager.FeatureCut4(True, False, False, swEndCondThroughAll, swEndCondBlind, 0, 0, _
            False, False, False, False, 0, 0, False, False, False, False, False, True, True, False, False, False, _
            swStartSketchPlane, 0, False, False)
        If feat Is Nothing Then Err.Raise vbObjectError + 16, , "Hole cut failed"
        expectedVolume = expectedVolume - HOLE_COUNT_X * HOLE_COUNT_Y * 3.14159265358979 * (HOLE_DIAMETER_MM / 2) ^ 2 * THICKNESS_MM
    End If

    ' ---- corner fillets: the 4 edges parallel to Z ----
    If CORNER_RADIUS_MM > 0 Then
        If SelectEdgesAlong(part, 0, 0, 1) <> 4 Then Err.Raise vbObjectError + 17, , "Expected 4 corner edges"
        Set feat = model.FeatureManager.FeatureFillet3(swFeatureFilletUniformRadius, CORNER_RADIUS_MM / 1000#, 0, 0, _
            swFeatureFilletType_Simple, 0, 0, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
        If feat Is Nothing Then Err.Raise vbObjectError + 18, , "Fillet failed"
        expectedVolume = expectedVolume - 4 * (CORNER_RADIUS_MM ^ 2 - 3.14159265358979 * CORNER_RADIUS_MM ^ 2 / 4) * THICKNESS_MM
    End If
    model.ClearSelection2 True

    ' ---- material (an unknown name is silently ignored -> read it back) ----
    If MATERIAL_NAME <> "" Then
        part.SetMaterialPropertyName2 model.ConfigurationManager.ActiveConfiguration.Name, MATERIAL_DB, MATERIAL_NAME
        If part.GetMaterialPropertyName2(model.ConfigurationManager.ActiveConfiguration.Name, db) <> MATERIAL_NAME Then
            Err.Raise vbObjectError + 19, , "Material '" & MATERIAL_NAME & "' not found in database '" & MATERIAL_DB & "'"
        End If
    End If
    If PROP_DESCRIPTION <> "" Then
        model.Extension.CustomPropertyManager("").Add3 "Opis", swCustomInfoText, PROP_DESCRIPTION, swCustomPropertyReplaceValue
    End If

    ' ---- rebuild + check the result against the inputs ----
    model.ForceRebuild3 False
    Set mp = model.Extension.CreateMassProperty
    If Abs(mp.Volume * 1000000000# - expectedVolume) > 0.001 * expectedVolume Then
        Err.Raise vbObjectError + 20, , "Volume " & Format(mp.Volume * 1000000000#, "0") & " mm3 <> expected " & Format(expectedVolume, "0")
    End If
    If Abs(ExtentMm(part, 0) - LENGTH_MM) > 0.01 Or Abs(ExtentMm(part, 1) - WIDTH_MM) > 0.01 _
       Or Abs(ExtentMm(part, 2) - THICKNESS_MM) > 0.01 Then
        Err.Raise vbObjectError + 21, , "Overall size differs from the inputs"
    End If

    ' ---- save ----
    If model.Extension.SaveAs(outPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, e, w) = False Or e <> 0 Then
        Err.Raise vbObjectError + 22, , "SaveAs failed, error " & e & " (" & outPath & ")"
    End If
    ' Measured: in a new part Extension.CreateMassProperty kept density 1 kg/m3 (mass ~0) after the material
    ' change; CreateMassProperty2 returned the right density at once. Check the density anyway.
    Set mp2 = model.Extension.CreateMassProperty2
    status = "CREATED"
    message = model.GetPathName & ";mass_kg=" & Format(mp2.Mass, "0.000") & ";density=" & Format(mp2.Density, "0") & _
              ";volume_mm3=" & Format(mp2.Volume * 1000000000#, "0")
    If mp2.Density < 100 Then message = message & ";WARNING mass not updated yet - reopen the part to read it"
    GoTo Cleanup
Failed:
    status = "ERROR": message = Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    If toggleChanged Then swApp.SetUserPreferenceToggle swInputDimValOnCreate, oldInputDim
    If Not model Is Nothing Then model.SketchManager.AddToDB = False
    On Error GoTo 0
    WriteResult status & ";" & message
    ' On error the unsaved new part stays open for inspection (it is never saved half-built).
    If SHOW_MESSAGES Then MsgBox status & vbCrLf & Replace(message, ";", vbCrLf), IIf(status = "CREATED", vbInformation, vbExclamation)
End Sub

' Standard planes in creation order: 0 = Front, 1 = Top, 2 = Right (names are localized - never select by name)
Private Function StandardPlane(model As SldWorks.ModelDoc2, index As Long) As SldWorks.Feature
    Dim f As SldWorks.Feature, n As Long
    Set f = model.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then
            If n = index Then Set StandardPlane = f: Exit Function
            n = n + 1
        End If
        Set f = f.GetNextFeature
    Loop
    Err.Raise vbObjectError + 30, , "Standard plane not found"
End Function

' First planar face whose outward normal is (nx, ny, nz) - returned as Entity so it can be selected
Private Function PlanarFaceWithNormal(part As SldWorks.PartDoc, nx As Double, ny As Double, nz As Double) As SldWorks.Entity
    Dim bodies As Variant, faces As Variant, f As SldWorks.Face2, n As Variant, i As Long
    bodies = part.GetBodies2(swSolidBody, True)
    If IsEmpty(bodies) Then Exit Function
    faces = bodies(0).GetFaces
    For i = 0 To UBound(faces)
        Set f = faces(i)
        n = f.Normal                                      ' (0,0,0) for non-planar faces
        If Abs(n(0) - nx) < 0.000001 And Abs(n(1) - ny) < 0.000001 And Abs(n(2) - nz) < 0.000001 Then
            Set PlanarFaceWithNormal = f
            Exit Function
        End If
    Next i
End Function

' Select every straight edge parallel to (dx, dy, dz); returns how many were selected
Private Function SelectEdgesAlong(part As SldWorks.PartDoc, dx As Double, dy As Double, dz As Double) As Long
    Dim bodies As Variant, edges As Variant, ed As SldWorks.Edge, c As SldWorks.Curve, p As Variant
    Dim ent As SldWorks.Entity, i As Long, model As SldWorks.ModelDoc2
    Set model = part
    model.ClearSelection2 True
    bodies = part.GetBodies2(swSolidBody, True)
    edges = bodies(0).GetEdges
    For i = 0 To UBound(edges)
        Set ed = edges(i)
        Set c = ed.GetCurve
        If c.IsLine Then
            p = c.LineParams                               ' root xyz, direction xyz
            If Abs(Abs(p(3) * dx + p(4) * dy + p(5) * dz) - 1) < 0.000001 Then
                Set ent = ed
                If ent.Select4(True, Nothing) <> False Then SelectEdgesAlong = SelectEdgesAlong + 1
            End If
        End If
    Next i
End Function

' Exact extent of the first solid body along axis 0=X, 1=Y, 2=Z in mm (GetPartBox is only approximate)
Private Function ExtentMm(part As SldWorks.PartDoc, axis As Long) As Double
    Dim bodies As Variant, b As SldWorks.Body2, d(2) As Double, x As Double, y As Double, z As Double, lo As Double
    bodies = part.GetBodies2(swSolidBody, True)
    Set b = bodies(0)
    d(axis) = -1
    b.GetExtremePoint d(0), d(1), d(2), x, y, z
    lo = Choose(axis + 1, x, y, z)
    d(axis) = 1
    b.GetExtremePoint d(0), d(1), d(2), x, y, z
    ExtentMm = (Choose(axis + 1, x, y, z) - lo) * 1000#
End Function

Private Sub WriteResult(text As String)
    Dim ff As Integer
    Debug.Print text
    If RESULT_PATH = "" Then Exit Sub
    On Error GoTo LogFailed
    ff = FreeFile
    Open RESULT_PATH For Append As #ff
    Print #ff, text
    Close #ff
    Exit Sub
LogFailed:
    Debug.Print "Result log failed: " & Err.Description
End Sub
