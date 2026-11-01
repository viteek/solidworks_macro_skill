Attribute VB_Name = "InsertComponentAndTransform"
Option Explicit
Private Const SHOW_MESSAGES As Boolean = True

' NOT_LIVE_TESTED. Active saved/clean, nonempty ASSEMBLY COPY; edit assembly mode.
' Source: saved/clean part COPY already open in its own SOLIDWORKS window.
' Inserts one instance and sets a root-relative transform in the active configuration.
' Changes/rebuilds the in-memory assembly; never saves, closes or changes PDM state.
Private Const TRANSLATION_X_M As Double = 0.1
Private Const TRANSLATION_Y_M As Double = 0.05
Private Const TRANSLATION_Z_M As Double = 0#
Private Const READBACK_TOLERANCE As Double = 0.000000001
Private Const RESULT_PATH As String = ""
Private Const SOURCE_PART_PATH As String = ""

Public Sub main()
    On Error GoTo Failed
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim assembly As SldWorks.AssemblyDoc
    Dim sourceModel As SldWorks.ModelDoc2
    Dim sourceConfiguration As SldWorks.Configuration
    Dim component As SldWorks.Component2
    Dim math As SldWorks.MathUtility
    Dim desired As SldWorks.MathTransform
    Dim observed As SldWorks.MathTransform
    Dim values(0 To 15) As Double
    Dim transformInput As Variant
    Dim readback As Variant
    Dim sourcePath As String
    Dim configurationName As String
    Dim instanceName As String
    Dim countBefore As Long
    Dim i As Long
    Dim resultLog As Object
    Dim outcome As String
    Dim failureText As String

    outcome = "FAIL"
    Set resultLog = BeginResultLog()
    If Not resultLog Is Nothing Then resultLog.WriteLine "START insert_component_transform; terminal RESULT line is required"
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a nonempty assembly copy."
    If model.GetType <> swDocASSEMBLY Then Err.Raise vbObjectError + 2, , "Active document must be an assembly copy."
    If Len(model.GetPathName) = 0 Then Err.Raise vbObjectError + 3, , "Save the assembly test copy manually first."
    If model.GetSaveFlag <> False Then Err.Raise vbObjectError + 4, , "Assembly test copy must have no unsaved changes."
    Set assembly = model
    countBefore = assembly.GetComponentCount(True)
    If countBefore = 0 Then Err.Raise vbObjectError + 5, , "Use a nonempty assembly; the first inserted component can be fixed automatically."
    sourcePath = SOURCE_PART_PATH
    If Len(sourcePath) = 0 Then sourcePath = Trim$(InputBox("Full path of a saved part COPY already open in its own SOLIDWORKS window", "Insert test component"))
    If Len(sourcePath) = 0 Then
        outcome = "CANCELLED"
        GoTo Cleanup
    End If
    If Len(Dir$(sourcePath)) = 0 Then Err.Raise vbObjectError + 6, , "Source file was not found."
    Set sourceModel = swApp.GetOpenDocumentByName(sourcePath)
    If sourceModel Is Nothing Then Err.Raise vbObjectError + 7, , "Open the source part copy in its own window before running."
    If sourceModel.GetType <> swDocPART Then Err.Raise vbObjectError + 8, , "Source must be a part."
    If sourceModel.GetSaveFlag <> False Then Err.Raise vbObjectError + 9, , "Source part has unsaved changes."
    If StrComp(sourceModel.GetPathName, sourcePath, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 10, , "Source document path did not match."
    Set sourceConfiguration = sourceModel.ConfigurationManager.ActiveConfiguration
    If sourceConfiguration Is Nothing Then Err.Raise vbObjectError + 11, , "Source configuration is unavailable."
    configurationName = sourceConfiguration.Name
    Set math = swApp.GetMathUtility
    If math Is Nothing Then Err.Raise vbObjectError + 12, , "MathUtility is unavailable."
    ' 0..8 rotation; 9..11 translation in meters; 12 scale; 13..15 unused.
    values(0) = 1#: values(4) = 1#: values(8) = 1#
    values(9) = TRANSLATION_X_M: values(10) = TRANSLATION_Y_M: values(11) = TRANSLATION_Z_M
    values(12) = 1#
    transformInput = values
    Set desired = math.CreateTransform(transformInput)
    If desired Is Nothing Then Err.Raise vbObjectError + 13, , "CreateTransform returned Nothing."
    ' Source must be preloaded. True uses the explicitly named source configuration.
    Set component = assembly.AddComponent5(sourcePath, swAddComponentConfigOptions_CurrentSelectedConfig, _
        "", True, configurationName, 0#, 0#, 0#)
    If component Is Nothing Then Err.Raise vbObjectError + 14, , "AddComponent5 returned Nothing."
    instanceName = component.Name2
    If component.IsFixed <> False Then Err.Raise vbObjectError + 15, , "New instance is fixed; inspect the inserted test instance manually."
    If component.SetTransformAndSolve3(desired, True) = False Then Err.Raise vbObjectError + 16, , "Transform/solve failed."
    If model.EditRebuild3 = False Then Err.Raise vbObjectError + 17, , "Assembly rebuild failed."
    Set component = assembly.GetComponentByName(instanceName)
    If component Is Nothing Then Err.Raise vbObjectError + 18, , "Inserted component could not be reacquired."
    If assembly.GetComponentCount(True) <> countBefore + 1 Then Err.Raise vbObjectError + 19, , "Top-level component count did not increase by one."
    If StrComp(component.GetPathName, sourcePath, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 20, , "Inserted source path readback mismatch."
    If component.ReferencedConfiguration <> configurationName Then Err.Raise vbObjectError + 21, , "Inserted configuration readback mismatch."
    Set observed = component.Transform2
    If observed Is Nothing Then Err.Raise vbObjectError + 22, , "Transform readback is unavailable."
    readback = observed.ArrayData
    If IsArray(readback) = False Then Err.Raise vbObjectError + 23, , "Transform readback did not return an array."
    If UBound(readback) - LBound(readback) + 1 <> 16 Then Err.Raise vbObjectError + 24, , "Expected 16 transform values."
    For i = 0 To 12
        If Abs(CDbl(readback(LBound(readback) + i)) - values(i)) > READBACK_TOLERANCE Then Err.Raise vbObjectError + 25, , "Transform readback mismatch at element " & CStr(i)
    Next i
    Debug.Print "Inserted instance="; instanceName; " source="; component.GetPathName; " configuration="; component.ReferencedConfiguration
    Debug.Print "Translation_m="; readback(9); readback(10); readback(11); " scale="; readback(12)
    outcome = "PASS"
    GoTo Cleanup
Failed:
    failureText = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    Err.Clear
    Debug.Print "RESULT insert_component_transform "; outcome
    If Len(failureText) > 0 Then Debug.Print failureText
    If Len(instanceName) > 0 Then Debug.Print "Inserted instance may remain: "; instanceName
    If Err.Number <> 0 Then
        outcome = "FAIL"
        failureText = failureText & "; reporting failed: " & Err.Description
        Err.Clear
    End If
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "instance=" & instanceName & "; source=" & sourcePath & "; configuration=" & configurationName
        resultLog.WriteLine "expected_translation_m=" & CStr(TRANSLATION_X_M) & "," & CStr(TRANSLATION_Y_M) & "," & CStr(TRANSLATION_Z_M)
        resultLog.WriteLine "detail=" & failureText
        resultLog.WriteLine "RESULT insert_component_transform " & outcome
        resultLog.Close
        If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED: "; Err.Number; " "; Err.Description
    End If
    If Len(RESULT_PATH) = 0 Then
        If outcome = "PASS" Then
            If SHOW_MESSAGES <> False Then MsgBox "Inserted instance and verified root-relative transform: " & instanceName & vbCrLf & "Nothing was saved.", vbInformation
        ElseIf outcome = "FAIL" Then
            If SHOW_MESSAGES <> False Then MsgBox failureText & vbCrLf & "An inserted instance may remain in the disposable copy. Nothing was saved.", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub

Private Function BeginResultLog() As Object
    Dim files As Object
    If Len(RESULT_PATH) = 0 Then Exit Function
    Set files = CreateObject("Scripting.FileSystemObject")
    Set BeginResultLog = files.CreateTextFile(RESULT_PATH, False, True)
End Function
