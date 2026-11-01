Attribute VB_Name = "ConfigurationVariantsRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. Requires untouched disposable sketch-extrude fixture.
' Adds Recipe_10mm and Recipe_20mm; restores original configuration. Never saves.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim target As SldWorks.Feature
    Dim definition As SldWorks.ExtrudeFeatureData2
    Dim depth As SldWorks.Dimension
    Dim config As SldWorks.Configuration
    Dim originalConfig As String
    Dim names As Variant
    Dim depths As Variant
    Dim values As Variant
    Dim index As Long
    Dim setStatus As Long
    Dim resultHandle As Integer
    Dim resultOpened As Boolean
    Dim resultText As String
    Dim failureText As String
    Dim cleanupText As String
    Dim cleanupOK As Boolean
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
    If model.GetConfigurationCount <> 1 Then Err.Raise vbObjectError + 4, , "Expected sole original configuration."
    If Not model.SketchManager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 5, , "Exit sketch edit mode."
    Set part = model
    originalConfig = model.ConfigurationManager.ActiveConfiguration.Name
    Set target = part.FeatureByName("Recipe_Block")
    If target Is Nothing Then Err.Raise vbObjectError + 6, , "Recipe_Block missing."
    Set definition = target.GetDefinition
    If definition Is Nothing Then Err.Raise vbObjectError + 7, , "No extrusion definition."
    If definition.GetEndCondition(True) <> swEndCondBlind Then Err.Raise vbObjectError + 8, , "Blind extrusion required."
    If Abs(definition.GetDepth(True) - 0.01) > 0.000000001 Then Err.Raise vbObjectError + 9, , "Expected untouched 10 mm depth."
    Set depth = target.Parameter("D1")
    If depth Is Nothing Then Err.Raise vbObjectError + 10, , "D1 missing in fixture extrusion."
    values = depth.GetSystemValue3(swThisConfiguration, Empty)
    If IsEmpty(values) Then Err.Raise vbObjectError + 11, , "No dimension readback."
    If Abs(CDbl(values(0)) - definition.GetDepth(True)) > 0.000000001 Then Err.Raise vbObjectError + 12, , "D1 is not the fixture depth."
    AssertVolume part, 0.000008
    names = Array("Recipe_10mm", "Recipe_20mm")
    depths = Array(0.01, 0.02)
    For index = LBound(names) To UBound(names)
        Set config = model.GetConfigurationByName(CStr(names(index)))
        If Not config Is Nothing Then Err.Raise vbObjectError + 13, , "Variant already exists."
    Next index
    For index = LBound(names) To UBound(names)
        Set config = model.AddConfiguration3(CStr(names(index)), "Original API recipe variant", "", swConfigOption_DontActivate)
        If config Is Nothing Then Err.Raise vbObjectError + 14, , "Configuration creation failed."
        ' Config_names is used only with the specific-configurations setter option.
        setStatus = depth.SetSystemValue3(CDbl(depths(index)), swSetValue_InSpecificConfigurations, CStr(names(index)))
        If setStatus <> swSetValue_Successful Then Err.Raise vbObjectError + 15, , "SetSystemValue3 status=" & setStatus
    Next index
    ' Activate and rebuild each configuration before checking geometry, not cached values.
    For index = LBound(names) To UBound(names)
        If model.ShowConfiguration2(CStr(names(index))) = False Then Err.Raise vbObjectError + 16, , "Activation failed."
        If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 17, , "Rebuild failed."
        Set target = part.FeatureByName("Recipe_Block")
        Set definition = target.GetDefinition
        If Abs(definition.GetDepth(True) - CDbl(depths(index))) > 0.000000001 Then Err.Raise vbObjectError + 18, , "Configuration depth mismatch."
        AssertVolume part, 0.04 * 0.02 * CDbl(depths(index))
        Debug.Print "VARIANT " & names(index) & "; depth_m=" & definition.GetDepth(True)
    Next index
    If model.ShowConfiguration2(originalConfig) = False Then Err.Raise vbObjectError + 19, , "Original configuration restore failed."
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 20, , "Original rebuild failed."
    AssertVolume part, 0.000008
    resultText = "RESULT configuration_variants OK; configurations=3; original_restored=" & originalConfig
    GoTo Cleanup
Failed:
    failureText = "FAIL configuration_variants " & Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Resume leaves the active VBA error handler before protected cleanup executes.
    On Error Resume Next
    Err.Clear
    If Len(failureText) <> 0 And Len(originalConfig) <> 0 Then
        cleanupOK = model.ShowConfiguration2(originalConfig)
        If Err.Number <> 0 Then
            cleanupText = "Configuration restore: " & Err.Description
        ElseIf cleanupOK = False Then
            cleanupText = "Configuration restore returned False."
        End If
        Err.Clear
    End If
    On Error GoTo 0
    If Len(failureText) <> 0 Then resultText = failureText
    If Len(cleanupText) <> 0 Then
        resultText = "FAIL configuration_variants cleanup: " & cleanupText & "; prior=" & resultText
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
    failureText = "FAIL configuration_variants result log: " & Err.Number & ": " & Err.Description
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
    If IsEmpty(bodies) Then Err.Raise vbObjectError + 30, , "No solid."
    If IsArray(bodies) = False Then Err.Raise vbObjectError + 31, , "Expected body array."
    If UBound(bodies) - LBound(bodies) + 1 <> 1 Then Err.Raise vbObjectError + 32, , "Expected one solid."
    Set body = bodies(LBound(bodies))
    properties = body.GetMassProperties(1#)
    If IsEmpty(properties) Then Err.Raise vbObjectError + 33, , "No mass properties."
    If Abs(CDbl(properties(3)) - expectedM3) > 0.00000000001 Then Err.Raise vbObjectError + 34, , "Volume mismatch."
End Sub

