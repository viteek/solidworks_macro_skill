Attribute VB_Name = "SelectionDiagnostic"
Option Explicit

' Read-only observation; does not change selection, geometry or saved CAD files.
' Empty RESULT_PATH prints only to the Immediate window. A file must be new.
Const RESULT_PATH As String = ""
Const EXPECTED_SELECTION_TYPE As Long = -1  ' -1 = observe; otherwise a swSelectType_e value

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim selections As SldWorks.SelectionMgr
    Dim selected As Object
    Dim owner As Object  ' Component2 in assemblies, DrawingComponent in drawings
    Dim feature As SldWorks.Feature
    Dim count As Long, index As Long, selectionType As Long
    Dim handle As Integer, fileOpened As Boolean, valid As Boolean
    Dim message As String

    On Error GoTo Failed
    If RESULT_PATH <> "" Then
        If Dir(RESULT_PATH) <> "" Then Err.Raise vbObjectError + 100, , "Diagnostic output already exists"
        handle = FreeFile
        Open RESULT_PATH For Output As #handle
        fileOpened = True
    End If
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 101, , "No active document"
    Set selections = model.SelectionManager
    count = selections.GetSelectedObjectCount2(-1)
    WriteLine handle, fileOpened, "Document: " & model.GetPathName
    WriteLine handle, fileOpened, "Document type (swDocumentTypes_e): " & CStr(model.GetType)
    WriteLine handle, fileOpened, "Selection count: " & CStr(count)
    valid = (count > 0)
    For index = 1 To count
        Set selected = selections.GetSelectedObject6(index, -1)
        selectionType = selections.GetSelectedObjectType3(index, -1)
        WriteLine handle, fileOpened, "Selected index: " & CStr(index)
        WriteLine handle, fileOpened, "Selection type (swSelectType_e): " & CStr(selectionType)
        WriteLine handle, fileOpened, "Selection mark: " & CStr(selections.GetSelectedObjectMark(index))
        If selected Is Nothing Then
            valid = False
            WriteLine handle, fileOpened, "Object resolved: NO"
        Else
            WriteLine handle, fileOpened, "Object resolved: YES"
            WriteLine handle, fileOpened, "API type: " & TypeName(selected)
            If TypeOf selected Is SldWorks.Feature Then
                Set feature = selected
                WriteLine handle, fileOpened, "Feature type: " & feature.GetTypeName2
                WriteLine handle, fileOpened, "Display name: " & feature.Name
            End If
        End If
        Set owner = selections.GetSelectedObjectsComponent4(index, -1)
        If owner Is Nothing Then
            WriteLine handle, fileOpened, "Owning component: None"
        Else
            WriteLine handle, fileOpened, "Owning component API type: " & TypeName(owner)
            If TypeOf owner Is SldWorks.Component2 Then
                Dim component As SldWorks.Component2
                Set component = owner
                WriteLine handle, fileOpened, "Owning component: " & component.Name2
            End If
        End If
        If EXPECTED_SELECTION_TYPE <> -1 Then
            If selectionType <> EXPECTED_SELECTION_TYPE Then valid = False
        End If
    Next index
    WriteLine handle, fileOpened, "Selection operation: NOT_PERFORMED (read-only diagnostic)"
    If valid = False Then
        WriteLine handle, fileOpened, "RESULT=FAIL"
    ElseIf EXPECTED_SELECTION_TYPE = -1 Then
        WriteLine handle, fileOpened, "RESULT=OBSERVED (expected type not specified)"
    Else
        WriteLine handle, fileOpened, "RESULT=PASS"
    End If
CleanUp:
    On Error Resume Next
    If fileOpened Then Close #handle
    If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED " & Err.Description
    On Error GoTo 0
    Exit Sub
Failed:
    message = "RESULT=FAIL; " & CStr(Err.Number) & "; " & Err.Description
    Resume ReportFailure
ReportFailure:
    On Error Resume Next
    Debug.Print message
    If fileOpened Then Print #handle, message
    If fileOpened Then Close #handle
    If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED " & Err.Description
    On Error GoTo 0
End Sub

Private Sub WriteLine(ByVal handle As Integer, ByVal fileOpened As Boolean, ByVal text As String)
    Debug.Print text
    If fileOpened Then Print #handle, text
End Sub
