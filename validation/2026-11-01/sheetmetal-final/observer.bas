Option Explicit
Const METRICS_PATH As String = ""
Private Const SHOW_MESSAGES As Boolean = True
Private Const INSPECTION_PATH As String = ""
Const EXPECTED_INPUT_PATH As String = ""
Const OUTPUT_PATH As String = ""
Const RESULT_PATH As String = ""
' Read-only solid metrics from the exact active source configuration; SI -> millimetres.

Private Sub CollectMetrics()
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

' NOT_LIVE_TESTED. Read-only active-part diagnostic; SOLIDWORKS 2018+.
' Reports bodies, SheetMetal/FlatPattern features and cut-list membership/properties.
' Reads current cut-list cache without UpdateCutList, rebuild, flatten, or save.

Private Sub InspectFixture()
    On Error GoTo Failed
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim feature As SldWorks.Feature
    Dim seen As Object
    Dim i As Long
    Dim failureText As String
    Dim resultLog As Object
    If Len(INSPECTION_PATH) > 0 Then
        Dim files As Object
        Set files = CreateObject("Scripting.FileSystemObject")
        Set resultLog = files.CreateTextFile(INSPECTION_PATH, False, True)
    End If
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a part."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Active document must be a part."
    Set part = model
    Debug.Print "SHEET METAL / CUT LIST: "; model.GetTitle
    bodies = part.GetBodies2(swSolidBody, False)
    If IsArray(bodies) <> False Then
        For i = LBound(bodies) To UBound(bodies)
            Set body = bodies(i)
            If Not body Is Nothing Then Debug.Print "Body: "; body.Name; " sheetMetal="; (body.IsSheetMetal <> False)
        Next i
    Else
        Debug.Print "No solid bodies in the active configuration."
    End If
    ' Feature names are unique within this part. Tree links can expose a feature twice.
    Set seen = CreateObject("Scripting.Dictionary")
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        InspectFeature feature, seen
        Set feature = feature.GetNextFeature
    Loop
    Debug.Print "END: cached cut-list data; freshness was not established."
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "RESULT inspect_sheetmetal_cutlist PASS"
        resultLog.Close
    End If
    Exit Sub
Failed:
    failureText = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    Debug.Print "RESULT inspect_sheetmetal_cutlist FAIL: "; failureText
    If SHOW_MESSAGES <> False Then MsgBox failureText, vbExclamation
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "RESULT inspect_sheetmetal_cutlist FAIL: " & failureText
        resultLog.Close
    End If
    If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED: "; Err.Number; " "; Err.Description
    On Error GoTo 0
End Sub

Private Sub InspectFeature(ByVal feature As SldWorks.Feature, ByVal seen As Object)
    Dim key As String
    Dim kind As String
    Dim child As SldWorks.Feature
    Dim folder As SldWorks.BodyFolder
    Dim sheetMetal As SldWorks.SheetMetalFeatureData
    Dim bodies As Variant
    Dim body As SldWorks.Body2
    Dim i As Long
    kind = feature.GetTypeName2
    key = kind & "|" & feature.Name
    If seen.Exists(key) <> False Then Exit Sub
    seen.Add key, True
    Select Case kind
        Case "SheetMetal"
            Set sheetMetal = feature.GetDefinition
            If sheetMetal Is Nothing Then
                Debug.Print "SheetMetal definition unavailable: "; feature.Name
            Else
                Debug.Print "SheetMetal: "; feature.Name; " thickness_m="; sheetMetal.Thickness
            End If
        Case "FlatPattern"
            Debug.Print "FlatPattern feature present: "; feature.Name; " (presence does not mean active/unsuppressed)"
        Case "CutListFolder"
            Set folder = feature.GetSpecificFeature2
            If folder Is Nothing Then Err.Raise vbObjectError + 10, , "CutListFolder has no BodyFolder object."
            Debug.Print "CutListFolder: "; feature.Name; " bodyCount="; folder.GetBodyCount
            bodies = folder.GetBodies
            If IsArray(bodies) <> False Then
                For i = LBound(bodies) To UBound(bodies)
                    Set body = bodies(i)
                    If Not body Is Nothing Then Debug.Print "  Member: "; body.Name; " sheetMetal="; (body.IsSheetMetal <> False)
                Next i
            End If
            InspectProperties feature
    End Select
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        InspectFeature child, seen
        Set child = child.GetNextSubFeature
    Loop
End Sub

Private Sub InspectProperties(ByVal feature As SldWorks.Feature)
    Dim properties As SldWorks.CustomPropertyManager
    Dim names As Variant
    Dim raw As String
    Dim evaluated As String
    Dim wasResolved As Boolean
    Dim linked As Boolean
    Dim result As Long
    Dim i As Long
    Set properties = feature.CustomPropertyManager
    If properties Is Nothing Then Exit Sub
    names = properties.GetNames
    If IsArray(names) = False Then Exit Sub
    For i = LBound(names) To UBound(names)
        ' Cached=True deliberately avoids demanding cross-configuration resolution.
        result = properties.Get6(CStr(names(i)), True, raw, evaluated, wasResolved, linked)
        Debug.Print "  Property: "; names(i); " status="; result; " raw="; raw; _
            " evaluated="; evaluated; " resolved="; wasResolved; " linked="; linked
    Next i
End Sub

' Observer for fixture preparation only: publish a NEW native copy of the test model.

Private Sub SaveFixture()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim errors As Long, warnings As Long, handle As Integer
    Dim result As String
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 100, , "No active fixture"
    If EXPECTED_INPUT_PATH = "" Then Err.Raise vbObjectError + 101, , "Specify expected test input"
    If StrComp(model.GetPathName, EXPECTED_INPUT_PATH, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 102, , "Active model is not the reviewed test copy"
    End If
    If OUTPUT_PATH = "" Or RESULT_PATH = "" Then Err.Raise vbObjectError + 103, , "Specify new output and log"
    If Dir(OUTPUT_PATH) <> "" Or Dir(RESULT_PATH) <> "" Then Err.Raise vbObjectError + 104, , "Output already exists"
    model.ClearSelection2 True
    If model.Extension.SaveAs(OUTPUT_PATH, swSaveAsCurrentVersion, _
        swSaveAsOptions_Silent + swSaveAsOptions_Copy, Nothing, errors, warnings) = False Then
        Err.Raise vbObjectError + 105, , "Cannot save fixture copy: " & CStr(errors)
    End If
    If errors <> 0 Or Dir(OUTPUT_PATH) = "" Then Err.Raise vbObjectError + 106, , "Fixture copy missing"
    If FileLen(OUTPUT_PATH) = 0 Then Err.Raise vbObjectError + 107, , "Fixture copy empty"
    result = "RESULT=PASS; fixture_copy=" & OUTPUT_PATH
WriteResult:
    On Error GoTo LogFailed
    Debug.Print result
    If RESULT_PATH <> "" Then
        If Dir(RESULT_PATH) = "" Then
            handle = FreeFile
            Open RESULT_PATH For Output As #handle
            Print #handle, result
            Close #handle
        End If
    End If
    Exit Sub
Failed:
    result = "RESULT=FAIL; " & CStr(Err.Number) & "; " & Err.Description
    Resume WriteResult
LogFailed:
    result = "RESULT_LOG_FAILED; " & CStr(Err.Number) & "; " & Err.Description
    Resume LogCleanup
LogCleanup:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    Debug.Print result
    On Error GoTo 0
End Sub
Sub main()
    CollectMetrics
    InspectFixture
    SaveFixture
End Sub
