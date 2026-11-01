Attribute VB_Name = "InspectDrawingSemantics"
Option Explicit
Private Const SHOW_MESSAGES As Boolean = True
Private Const RESULT_PATH As String = ""

' NOT_LIVE_TESTED. Read-only diagnostic; SOLIDWORKS 2018+.
' Input: active drawing; output: VBA Immediate window (Ctrl+G).
' Includes sheet pseudo-views; no sheet activation, selection, rebuild, or save.

Public Sub main()
    On Error GoTo Failed
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim drawing As SldWorks.DrawingDoc
    Dim view As SldWorks.View
    Dim sheets As Variant
    Dim views As Variant
    Dim sheetIndex As Long
    Dim viewIndex As Long
    Dim failureText As String
    Dim resultLog As Object
    If Len(RESULT_PATH) > 0 Then
        Dim files As Object
        Set files = CreateObject("Scripting.FileSystemObject")
        Set resultLog = files.CreateTextFile(RESULT_PATH, False, True)
    End If
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a drawing."
    If model.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 2, , "Active document must be a drawing."
    Set drawing = model
    sheets = drawing.GetViews
    If IsEmpty(sheets) Then Err.Raise vbObjectError + 3, , "No sheet/view arrays returned."
    If IsArray(sheets) = False Then Err.Raise vbObjectError + 4, , "Unexpected GetViews result."
    Debug.Print "DRAWING SEMANTIC INSPECTION: "; model.GetTitle
    For sheetIndex = LBound(sheets) To UBound(sheets)
        views = sheets(sheetIndex)
        If IsArray(views) <> False Then
            For viewIndex = LBound(views) To UBound(views)
                Set view = views(viewIndex)
                If Not view Is Nothing Then
                    Debug.Print "Sheet-array "; sheetIndex; " view="; view.GetName2; " sheet-pseudo-view="; (viewIndex = LBound(views))
                    InspectAnnotations view
                    InspectTables view
                End If
            Next viewIndex
        End If
    Next sheetIndex
    Debug.Print "END: inspect text values against the drawing UI; no model changes requested."
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "RESULT inspect_drawing_semantics PASS"
        resultLog.Close
    End If
    Exit Sub
Failed:
    failureText = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' This diagnostic changes no document/UI state; reporting is still protected.
    On Error Resume Next
    Debug.Print "RESULT inspect_drawing_semantics FAIL: "; failureText
    If SHOW_MESSAGES <> False Then MsgBox failureText, vbExclamation
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "RESULT inspect_drawing_semantics FAIL: " & failureText
        resultLog.Close
    End If
    If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED: "; Err.Number; " "; Err.Description
    On Error GoTo 0
End Sub

Private Sub InspectAnnotations(ByVal view As SldWorks.View)
    Dim annotations As Variant
    Dim annotation As SldWorks.Annotation
    Dim note As SldWorks.Note
    Dim gtol As SldWorks.Gtol
    Dim weld As SldWorks.WeldSymbol
    Dim specific As Object
    Dim i As Long
    Dim j As Long
    Dim count As Long
    count = view.GetAnnotationCount
    Debug.Print "  Annotations="; count
    If count = 0 Then Exit Sub
    annotations = view.GetAnnotations
    If IsArray(annotations) = False Then Err.Raise vbObjectError + 10, , "Annotation count/array mismatch."
    For i = LBound(annotations) To UBound(annotations)
        Set annotation = annotations(i)
        If Not annotation Is Nothing Then
            Debug.Print "  Annotation type="; annotation.GetType
            Set specific = annotation.GetSpecificAnnotation
            If Not specific Is Nothing Then
                Select Case annotation.GetType
                    Case swNote
                        Set note = specific
                        Debug.Print "    Note: "; note.GetText
                    Case swGTol
                        Set gtol = specific
                        Debug.Print "    GTol stored frames="; gtol.GetFrameCount
                        For j = 0 To gtol.GetTextCount - 1
                            Debug.Print "      GTol text["; j; "]="; gtol.GetTextAtIndex(j)
                        Next j
                    Case swWeldSymbol
                        Set weld = specific
                        For j = 0 To weld.GetTextCount - 1
                            Debug.Print "      Weld text["; j; "]="; weld.GetTextAtIndex(j)
                        Next j
                End Select
            Else
                Debug.Print "    No specific object (PMI-only annotations require PMI APIs)."
            End If
        End If
    Next i
End Sub

Private Sub InspectTables(ByVal view As SldWorks.View)
    Dim tables As Variant
    Dim table As SldWorks.TableAnnotation
    Dim i As Long
    Dim row As Long
    Dim column As Long
    tables = view.GetTableAnnotations
    If IsEmpty(tables) Then Exit Sub
    If IsArray(tables) = False Then Exit Sub
    For i = LBound(tables) To UBound(tables)
        Set table = tables(i)
        If Not table Is Nothing Then
            Debug.Print "  Table type="; table.Type; " rows="; table.TotalRowCount; " columns="; table.TotalColumnCount
            ' Include hidden cells. Text is displayed content, not BOM component identity.
            For row = 0 To table.TotalRowCount - 1
                For column = 0 To table.TotalColumnCount - 1
                    Debug.Print "    ["; row; ","; column; "]="; table.DisplayedText2(row, column, True)
                Next column
            Next row
        End If
    Next i
End Sub
