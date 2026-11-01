Attribute VB_Name = "PersistentReferenceRecipe"
Option Explicit

Private Const RESULT_PATH As String = ""

' Original recipe. NOT_LIVE_TESTED. No save, close, or file I/O.
' Select exactly one supported object in a disposable part, outside sketch edit.
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim original As Object
    Dim resolved As Object
    Dim referenceBytes As Variant
    Dim restoredBytes As Variant
    Dim encoded As String
    Dim state As Long
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
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a disposable part."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Part required."
    If Not model.SketchManager.ActiveSketch Is Nothing Then Err.Raise vbObjectError + 3, , "Exit sketch edit mode."
    If model.SelectionManager.GetSelectedObjectCount2(-1) <> 1 Then Err.Raise vbObjectError + 4, , "Select exactly one object."
    Set original = model.SelectionManager.GetSelectedObject6(1, -1)
    If original Is Nothing Then Err.Raise vbObjectError + 5, , "Selection is unavailable."
    referenceBytes = model.Extension.GetPersistReference3(original)
    If IsEmpty(referenceBytes) Then Err.Raise vbObjectError + 6, , "Object has no persistent reference."
    If IsArray(referenceBytes) = False Then Err.Raise vbObjectError + 7, , "Expected a byte array."
    encoded = EncodeBytes(referenceBytes)
    restoredBytes = DecodeBytes(encoded)
    Set original = Nothing
    ' Rebuild makes this a stronger test than resolving the same cached pointer.
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 8, , "Rebuild failed."
    Set resolved = model.Extension.GetObjectByPersistReference3(restoredBytes, state)
    If state <> swPersistReferencedObject_Ok Then Err.Raise vbObjectError + 9, , "Reference state=" & state
    If resolved Is Nothing Then Err.Raise vbObjectError + 10, , "Resolution returned Nothing."
    ' Compare through the API, not VBA object pointer identity.
    Set original = model.Extension.GetObjectByPersistReference3(referenceBytes, state)
    If state <> swPersistReferencedObject_Ok Then Err.Raise vbObjectError + 11, , "Original reference no longer resolves."
    If original Is Nothing Then Err.Raise vbObjectError + 12, , "Original resolution returned Nothing."
    If swApp.IsSame(original, resolved) <> swObjectSame Then Err.Raise vbObjectError + 13, , "Roundtrip resolved a different object."
    Debug.Print "REFERENCE decimal-triplets=" & encoded
    resultText = "RESULT persistent_roundtrip OK; bytes=" & Len(encoded) / 3
    GoTo Cleanup
Failed:
    failureText = "FAIL persistent_roundtrip " & Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Resume leaves the active VBA error handler before protected cleanup executes.
    On Error Resume Next
    Err.Clear
    On Error GoTo 0
    If Len(failureText) <> 0 Then resultText = failureText
    If Len(cleanupText) <> 0 Then
        resultText = "FAIL persistent_roundtrip cleanup: " & cleanupText & "; prior=" & resultText
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
    failureText = "FAIL persistent_roundtrip result log: " & Err.Number & ": " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    Err.Clear
    If resultOpened <> False Then Close #resultHandle
    If Err.Number <> 0 Then Debug.Print "CLEANUP_FAILED result log: " & Err.Description
    On Error GoTo 0
    Debug.Print failureText
End Sub

Private Function EncodeBytes(ByVal values As Variant) As String
    Dim index As Long
    Dim byteValue As Long
    For index = LBound(values) To UBound(values)
        byteValue = CLng(values(index))
        If byteValue < 0 Or byteValue > 255 Then Err.Raise vbObjectError + 20, , "Invalid byte."
        EncodeBytes = EncodeBytes & Right$("000" & CStr(byteValue), 3)
    Next index
    If Len(EncodeBytes) = 0 Then Err.Raise vbObjectError + 21, , "Empty reference."
End Function

Private Function DecodeBytes(ByVal encoded As String) As Variant
    Dim values() As Byte
    Dim index As Long
    Dim byteValue As Long
    Dim token As String
    If Len(encoded) = 0 Or Len(encoded) Mod 3 <> 0 Then Err.Raise vbObjectError + 22, , "Invalid encoded length."
    ReDim values(0 To Len(encoded) / 3 - 1)
    For index = 0 To UBound(values)
        token = Mid$(encoded, index * 3 + 1, 3)
        If token Like "[0-9][0-9][0-9]" = False Then Err.Raise vbObjectError + 23, , "Reference contains a non-digit."
        byteValue = CLng(token)
        If byteValue > 255 Then Err.Raise vbObjectError + 24, , "Reference byte exceeds 255."
        values(index) = CByte(byteValue)
    Next index
    DecodeBytes = values
End Function
