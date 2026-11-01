Option Explicit
Public Sub main()
    On Error GoTo Failed
    Dim app As SldWorks.SldWorks
    Dim original As SldWorks.ModelDoc2, source As SldWorks.ModelDoc2, created As SldWorks.ModelDoc2
    Dim assembly As SldWorks.AssemblyDoc
    Dim component As SldWorks.Component2
    Dim math As SldWorks.MathUtility, transform As SldWorks.MathTransform
    Dim values(0 To 15) As Double, matrixValues As Variant
    Dim errors As Long, warnings As Long, handle As Integer
    Dim result As String, originalName As String
    Set app = Application.SldWorks
    Set original = app.ActiveDoc
    originalName = original.GetTitle
    If Len(Dir$("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\fixture.sldasm")) > 0 Then Err.Raise vbObjectError + 1, , "Assembly output exists"
    Set source = app.OpenDoc6("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\block.sldprt", swDocPART, swOpenDocOptions_Silent, "", errors, warnings)
    If source Is Nothing Then Err.Raise vbObjectError + 2, , "Source open failed"
    If errors <> 0 Then Err.Raise vbObjectError + 3, , "Source open errors"
    Set created = app.NewDocument("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\assembly.asmdot", 0, 0#, 0#)
    If created Is Nothing Then Err.Raise vbObjectError + 4, , "Assembly creation failed"
    Set assembly = created
    Set component = assembly.AddComponent5("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\block.sldprt", swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0#, 0#, 0#)
    If component Is Nothing Then Err.Raise vbObjectError + 5, , "First component failed"
    Set component = assembly.AddComponent5("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\block.sldprt", swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0#, 0#, 0#)
    If component Is Nothing Then Err.Raise vbObjectError + 6, , "Second component failed"
    Set math = app.GetMathUtility
    values(0) = 1#: values(4) = 1#: values(8) = 1#: values(9) = 0.1: values(10) = 0.05: values(12) = 1#
    matrixValues = values
    Set transform = math.CreateTransform(matrixValues)
    If component.SetTransformAndSolve3(transform, True) = False Then Err.Raise vbObjectError + 7, , "Fixture transform failed"
    If created.EditRebuild3 = False Then Err.Raise vbObjectError + 8, , "Fixture rebuild failed"
    If created.Extension.SaveAs("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\fixture.sldasm", swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings) = False Then Err.Raise vbObjectError + 9, , "Fixture save failed"
    If errors <> 0 Then Err.Raise vbObjectError + 10, , "Fixture save errors"
    result = "RESULT fixture PASS; components=" & CStr(assembly.GetComponentCount(True))
    GoTo Cleanup
Failed:
    result = "RESULT fixture FAIL " & CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    If Not created Is Nothing Then app.CloseDoc created.GetTitle
    If Not source Is Nothing Then app.CloseDoc source.GetTitle
    Set original = app.ActivateDoc3(originalName, False, swDontRebuildActiveDoc, errors)
    If Err.Number <> 0 Then result = "RESULT fixture FAIL cleanup " & Err.Description
    handle = FreeFile
    Open "C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\fixture.log" For Output As #handle
    Print #handle, result
    Close #handle
End Sub
