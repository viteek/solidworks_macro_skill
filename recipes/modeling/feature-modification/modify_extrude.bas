Attribute VB_Name = "ModifyExtrudeRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. Run on the disposable sketch-extrude result.
' Changes Recipe_Block depth from 10 to 15 mm in its sole configuration. No save.
Private Const NEW_DEPTH_M As Double = 0.015

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim target As SldWorks.Feature
    Dim definition As SldWorks.ExtrudeFeatureData2
    Dim readback As SldWorks.ExtrudeFeatureData2
    Dim selectionAccess As Boolean
    Dim oldDepthM As Double
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
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open the disposable recipe block."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    If model.GetConfigurationCount <> 1 Then Err.Raise vbObjectError + 4, , "Exactly one configuration required."
    If Not model.SketchManager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 5, , "Exit sketch edit mode."
    If NEW_DEPTH_M <= 0 Or NEW_DEPTH_M > 0.1 Then Err.Raise vbObjectError + 6, , "Depth must be >0 and <=0.1 m."
    Set part = model
    Set target = part.FeatureByName("Recipe_Block")
    If target Is Nothing Then Err.Raise vbObjectError + 7, , "Recipe_Block not found."
    Set definition = target.GetDefinition
    If definition Is Nothing Then Err.Raise vbObjectError + 8, , "No extrusion definition."
    If definition.GetEndCondition(True) <> swEndCondBlind Then Err.Raise vbObjectError + 9, , "Blind extrusion required."
    oldDepthM = definition.GetDepth(True)
    If Abs(oldDepthM - 0.01) > 0.000000001 Then Err.Raise vbObjectError + 10, , "Expected untouched 10 mm block."
    AssertVolume part, 0.04 * 0.02 * oldDepthM
    model.ClearSelection2 True
    If definition.AccessSelections(model, Nothing) = False Then Err.Raise vbObjectError + 11, , "AccessSelections failed."
    selectionAccess = True
    definition.SetDepth True, NEW_DEPTH_M
    If target.ModifyDefinition(definition, model, Nothing) = False Then Err.Raise vbObjectError + 12, , "ModifyDefinition failed."
    ' Successful ModifyDefinition releases access; never reuse its cached definition.
    selectionAccess = False
    Set definition = Nothing
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 13, , "Rebuild failed."
    Set readback = target.GetDefinition
    If readback Is Nothing Then Err.Raise vbObjectError + 14, , "Readback definition missing."
    If Abs(readback.GetDepth(True) - NEW_DEPTH_M) > 0.000000001 Then Err.Raise vbObjectError + 15, , "Depth readback mismatch."
    AssertVolume part, 0.04 * 0.02 * NEW_DEPTH_M
    model.ClearSelection2 True
    resultText = "RESULT modify_extrude OK; depth_m=" & readback.GetDepth(True) & "; volume_m3=" & 0.04 * 0.02 * NEW_DEPTH_M
    GoTo Cleanup
Failed:
    failureText = "FAIL modify_extrude " & Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Resume leaves the active VBA error handler before protected cleanup executes.
    On Error Resume Next
    Err.Clear
    If selectionAccess <> False Then
        definition.ReleaseSelectionAccess
        If Err.Number <> 0 Then cleanupText = "ReleaseSelectionAccess: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If Len(failureText) <> 0 Then resultText = failureText
    If Len(cleanupText) <> 0 Then
        resultText = "FAIL modify_extrude cleanup: " & cleanupText & "; prior=" & Replace(resultText, "RESULT modify_extrude OK; ", "")
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
    failureText = "FAIL modify_extrude result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub

Private Sub AssertVolume(ByVal part As SldWorks.PartDoc, ByVal expectedM3 As Double)
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim properties As Variant
    bodies = part.GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) Then Err.Raise vbObjectError + 20, , "No solid."
    If IsArray(bodies) = False Then Err.Raise vbObjectError + 21, , "Expected body array."
    If UBound(bodies) - LBound(bodies) + 1 <> 1 Then Err.Raise vbObjectError + 22, , "Expected one solid."
    Set body = bodies(LBound(bodies))
    properties = body.GetMassProperties(1#)
    If IsEmpty(properties) Then Err.Raise vbObjectError + 23, , "No mass properties."
    If Abs(CDbl(properties(3)) - expectedM3) > 0.00000000001 Then Err.Raise vbObjectError + 24, , "Volume mismatch."
End Sub

