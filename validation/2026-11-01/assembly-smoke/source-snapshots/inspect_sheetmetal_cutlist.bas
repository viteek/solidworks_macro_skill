Attribute VB_Name = "InspectSheetMetalCutList"
Option Explicit
Private Const SHOW_MESSAGES As Boolean = True
Private Const RESULT_PATH As String = ""

' NOT_LIVE_TESTED. Read-only active-part diagnostic; SOLIDWORKS 2018+.
' Reports bodies, SheetMetal/FlatPattern features and cut-list membership/properties.
' Reads current cut-list cache without UpdateCutList, rebuild, flatten, or save.

Public Sub main()
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
    If Len(RESULT_PATH) > 0 Then
        Dim files As Object
        Set files = CreateObject("Scripting.FileSystemObject")
        Set resultLog = files.CreateTextFile(RESULT_PATH, False, True)
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
