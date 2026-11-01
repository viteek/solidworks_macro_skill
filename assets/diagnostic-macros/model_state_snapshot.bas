Attribute VB_Name = "ModelStateSnapshot"
Option Explicit

' Independent read-only observer. Collects state, never certifies the tested macro.
Const RESULT_PATH As String = ""

Sub main()
    Dim swApp As SldWorks.SldWorks, model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc, feature As SldWorks.Feature
    Dim sketch As SldWorks.Sketch, segment As SldWorks.SketchSegment, body As SldWorks.Body2
    Dim names As Variant, segments As Variant, bodies As Variant, properties As Variant
    Dim i As Long, count As Long, handle As Integer
    Dim length As Double, volume As Double
    Dim text As String, item As String, failure As String, isOpen As Boolean
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 100, , "No active test document"
    If RESULT_PATH = "" Then Err.Raise vbObjectError + 101, , "Specify a fresh observation file"
    If Dir(RESULT_PATH) <> "" Then Err.Raise vbObjectError + 102, , "Observation file already exists"
    text = "{""collection_status"":""OK"",""path"":" & JsonString(model.GetPathName) & _
           ",""document_type"":" & CStr(model.GetType) & ",""selection_count"":" & CStr(model.SelectionManager.GetSelectedObjectCount2(-1))
    If model.GetType <> swDocDRAWING Then
        text = text & ",""active_configuration"":" & JsonString(model.ConfigurationManager.ActiveConfiguration.Name)
        text = text & ",""configurations"":" & JsonStrings(model.GetConfigurationNames)
    End If
    text = text & ",""features"":" & "["
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        If count > 0 Then text = text & ","
        item = "{""name"":" & JsonString(feature.Name) & ",""type"":" & JsonString(feature.GetTypeName2)
        If feature.GetTypeName2 = "3DProfileFeature" Or feature.GetTypeName2 = "ProfileFeature" Then
            Set sketch = feature.GetSpecificFeature2
            If Not sketch Is Nothing Then
                segments = sketch.GetSketchSegments
                length = 0
                If IsEmpty(segments) = False Then
                    For i = LBound(segments) To UBound(segments)
                        Set segment = segments(i)
                        length = length + segment.GetLength
                    Next i
                    item = item & ",""segment_count"":" & CStr(UBound(segments) - LBound(segments) + 1)
                Else
                    item = item & ",""segment_count"":0"
                End If
                item = item & ",""length_m"":" & JsonNumber(length)
                item = item & ",""constrained_status"":" & CStr(sketch.GetConstrainedStatus)
            End If
        End If
        text = text & item & "}"
        count = count + 1
        Set feature = feature.GetNextFeature
    Loop
    text = text & "]"
    If model.GetType = swDocPART Then
        Set part = model
        bodies = part.GetBodies2(swSolidBody, False)
        count = 0
        If IsEmpty(bodies) = False Then
            count = UBound(bodies) - LBound(bodies) + 1
            For i = LBound(bodies) To UBound(bodies)
                Set body = bodies(i)
                properties = body.GetMassProperties(1#)
                If IsEmpty(properties) Then Err.Raise vbObjectError + 103, , "Mass property collection failed"
                volume = volume + CDbl(properties(3))
            Next i
        End If
        text = text & ",""solid_body_count"":" & CStr(count) & ",""volume_m3"":" & JsonNumber(volume)
    End If
    text = text & "}"
    handle = FreeFile
    Open RESULT_PATH For Output As #handle
    isOpen = True
    Print #handle, text
    Close #handle
    isOpen = False
    Exit Sub
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
    Resume CleanUp
CleanUp:
    On Error Resume Next
    If isOpen Then Close #handle
    If Err.Number <> 0 Then Debug.Print "Observer cleanup failed: " & Err.Description
    On Error GoTo 0
    Debug.Print "COLLECTION_FAILED " & failure
    If RESULT_PATH <> "" Then
        If Dir(RESULT_PATH) = "" Then
            handle = FreeFile
            Open RESULT_PATH For Output As #handle
            Print #handle, "{""collection_status"":""FAIL"",""error"":" & JsonString(failure) & "}"
            Close #handle
        End If
    End If
End Sub

Private Function JsonNumber(ByVal value As Double) As String
    Dim text As String
    text = Trim(Str(value))
    If Left(text, 1) = "." Then text = "0" & text
    If Left(text, 2) = "-." Then text = "-0" & Mid(text, 2)
    JsonNumber = text
End Function

Private Function JsonStrings(ByVal values As Variant) As String
    Dim i As Long, text As String
    text = "["
    If IsEmpty(values) = False Then
        For i = LBound(values) To UBound(values)
            If i > LBound(values) Then text = text & ","
            text = text & JsonString(CStr(values(i)))
        Next i
    End If
    JsonStrings = text & "]"
End Function

Private Function JsonString(ByVal value As String) As String
    Dim i As Long, code As Long, character As String, text As String
    text = """"
    For i = 1 To Len(value)
        character = Mid(value, i, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34: text = text & "\"""
            Case 92: text = text & "\\"
            Case 32 To 126: text = text & character
            Case Else: text = text & "\u" & Right("0000" & Hex(code), 4)
        End Select
    Next i
    JsonString = text & """"
End Function
