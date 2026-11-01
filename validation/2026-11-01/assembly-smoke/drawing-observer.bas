Option Explicit
Public Sub main()
    On Error GoTo Failed
    Dim app As SldWorks.SldWorks
    Dim original As SldWorks.ModelDoc2, model As SldWorks.ModelDoc2
    Dim drawing As SldWorks.DrawingDoc, view As SldWorks.View
    Dim documents As Variant, sheets As Variant, views As Variant
    Dim i As Long, j As Long, k As Long, count As Long, handle As Integer
    Dim errors As Long, warnings As Long, result As String
    Set app = Application.SldWorks
    Set original = app.ActiveDoc
    If StrComp(original.GetPathName, "C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\block.sldprt", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1, , "Input activation not restored"
    documents = app.GetDocuments
    For i = LBound(documents) To UBound(documents)
        Set model = documents(i)
        If model.GetType = swDocDRAWING Then
            If Len(model.GetPathName) <> 0 Then Err.Raise vbObjectError + 2, , "Drawing should be unsaved"
            Set drawing = model
            sheets = drawing.GetViews
            For j = LBound(sheets) To UBound(sheets)
                views = sheets(j)
                For k = LBound(views) + 1 To UBound(views)
                    Set view = views(k)
                    If StrComp(view.GetReferencedModelName, "C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\block.sldprt", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 3, , "Observer reference mismatch"
                    count = count + 1
                Next k
            Next j
            If count < 3 Then Err.Raise vbObjectError + 4, , "Observer found fewer than 3 views"
            If Len(Dir$("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\fixture.slddrw")) > 0 Then Err.Raise vbObjectError + 5, , "Drawing output exists"
            If model.Extension.SaveAs("C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\fixture.slddrw", swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings) = False Then Err.Raise vbObjectError + 6, , "Drawing fixture save failed"
            If errors <> 0 Then Err.Raise vbObjectError + 7, , "Drawing save errors"
            app.CloseDoc model.GetTitle
        End If
    Next i
    If count < 3 Then Err.Raise vbObjectError + 8, , "No created drawing found"
    result = "RESULT drawing_observer PASS; model_views=" & CStr(count) & "; input_restored=True"
    GoTo Cleanup
Failed:
    result = "RESULT drawing_observer FAIL " & CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    handle = FreeFile
    Open "C:\Users\micha\.codex\tmp\sw-audit-fxy61gm3\assembly-smoke-live2\drawing-observed.log" For Output As #handle
    Print #handle, result
    Close #handle
    If Err.Number <> 0 Then Debug.Print "OBSERVER_LOG_FAILED: "; Err.Description
End Sub
