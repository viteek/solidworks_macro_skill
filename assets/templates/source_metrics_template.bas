Attribute VB_Name = "Main"
Option Explicit
' Read-only solid metrics from the exact active source configuration; SI -> millimetres.
Const METRICS_PATH As String = ""

Sub main()
    Dim app As SldWorks.SldWorks, model As SldWorks.ModelDoc2
    Set app = Application.SldWorks
    Set model = app.ActiveDoc
    WriteSourceMetrics model, METRICS_PATH
End Sub

Private Sub WriteSourceMetrics(model As SldWorks.ModelDoc2, path As String)
    Dim part As SldWorks.PartDoc, body As SldWorks.Body2
    Dim bodies As Variant, properties As Variant, i As Long, axis As Long
    Dim x As Double, y As Double, z As Double, direction(2) As Double, bounds(5) As Double
    Dim payload As String, item As String, ff As Integer, isOpen As Boolean
    Dim failure As String
    On Error GoTo Failed
    If model Is Nothing Then Err.Raise vbObjectError + 100, , "No active source model"
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 101, , "Solid geometry snapshot currently supports parts only"
    If model.GetPathName = "" Then Err.Raise vbObjectError + 102, , "Source model must have a saved path"
    If path = "" Or Dir(path) <> "" Then Err.Raise vbObjectError + 103, , "Provide a NEW metrics output path"
    Set part = model
    bodies = part.GetBodies2(swSolidBody, True)   ' visible only: hidden bodies are not exported to STEP
    If IsEmpty(bodies) Then Err.Raise vbObjectError + 104, , "Source contains no solid bodies"
    If UBound(bodies) - LBound(bodies) + 1 > 500 Then Err.Raise vbObjectError + 105, , "Geometry snapshot body limit exceeded"
    payload = "{""schema_version"":1,""units"":""mm"",""source_path"":" & JsonString(model.GetPathName) & _
              ",""configuration"":" & JsonString(model.ConfigurationManager.ActiveConfiguration.Name) & ",""bodies"":["
    For i = LBound(bodies) To UBound(bodies)
        Set body = bodies(i)
        properties = body.GetMassProperties(1#)
        If IsEmpty(properties) Then Err.Raise vbObjectError + 106, , "Cannot measure body properties"
        If properties(3) <= 0 Or properties(4) <= 0 Then Err.Raise vbObjectError + 107, , "Invalid source solid properties"
        For axis = 0 To 2
            direction(0) = 0: direction(1) = 0: direction(2) = 0
            direction(axis) = -1
            If body.GetExtremePoint(direction(0), direction(1), direction(2), x, y, z) = False Then
                Err.Raise vbObjectError + 108, , "Cannot obtain minimum extreme point"
            End If
            Select Case axis
                Case 0: bounds(axis) = x * 1000#
                Case 1: bounds(axis) = y * 1000#
                Case 2: bounds(axis) = z * 1000#
            End Select
            direction(axis) = 1
            If body.GetExtremePoint(direction(0), direction(1), direction(2), x, y, z) = False Then
                Err.Raise vbObjectError + 109, , "Cannot obtain maximum extreme point"
            End If
            Select Case axis
                Case 0: bounds(axis + 3) = x * 1000#
                Case 1: bounds(axis + 3) = y * 1000#
                Case 2: bounds(axis + 3) = z * 1000#
            End Select
        Next axis
        item = "{""volume_mm3"":" & JsonNumber(properties(3) * 1000000000#) & _
               ",""area_mm2"":" & JsonNumber(properties(4) * 1000000#) & _
               ",""centroid_mm"":" & JsonNumbers(Array(properties(0) * 1000#, properties(1) * 1000#, properties(2) * 1000#)) & _
               ",""bounds_mm"":" & JsonNumbers(Array(bounds(0), bounds(1), bounds(2), bounds(3), bounds(4), bounds(5))) & "}"
        If i > LBound(bodies) Then payload = payload & ","
        payload = payload & item
    Next i
    payload = payload & "]}"
    ff = FreeFile
    Open path For Output As #ff
    isOpen = True
    Print #ff, payload
    Close #ff
    Exit Sub
Failed:
    failure = Err.Number & ": " & Err.Description
    Resume CloseFile
CloseFile:
    On Error Resume Next
    If isOpen Then Close #ff
    On Error GoTo 0
    Err.Raise vbObjectError + 110, , failure
End Sub

Private Function JsonNumber(value As Double) As String
    Dim text As String
    text = Trim(Str(value))
    If Left(text, 1) = "." Then text = "0" & text
    If Left(text, 2) = "-." Then text = "-0" & Mid(text, 2)
    JsonNumber = text
End Function

Private Function JsonNumbers(values As Variant) As String
    Dim i As Long, text As String
    text = "["
    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then text = text & ","
        text = text & JsonNumber(CDbl(values(i)))
    Next i
    JsonNumbers = text & "]"
End Function

Private Function JsonString(value As String) As String
    Dim i As Long, code As Long, text As String, character As String
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
