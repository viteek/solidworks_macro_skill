Attribute VB_Name = "CreateThreeDrawingViews"
Option Explicit
Private Const SHOW_MESSAGES As Boolean = True

' NOT_LIVE_TESTED. Creates a NEW UNSAVED drawing; never saves the input or output.
' Input: active, saved, clean part COPY and an existing local .drwdot template.
' Leave this empty to read the user's default drawing template preference.
Private Const DRAWING_TEMPLATE As String = ""

Public Sub main()
    On Error GoTo Failed
    Dim swApp As SldWorks.SldWorks
    Dim inputModel As SldWorks.ModelDoc2
    Dim drawingModel As SldWorks.ModelDoc2
    Dim drawing As SldWorks.DrawingDoc
    Dim view As SldWorks.View
    Dim sheets As Variant
    Dim views As Variant
    Dim templatePath As String
    Dim inputPath As String
    Dim inputFileName As String
    Dim referencedPath As String
    Dim outputTitle As String
    Dim sheetIndex As Long
    Dim viewIndex As Long
    Dim modelViewCount As Long
    Dim failureMessage As String
    Dim restored As Boolean
    Dim succeeded As Boolean

    Set swApp = Application.SldWorks
    Set inputModel = swApp.ActiveDoc
    If inputModel Is Nothing Then Err.Raise vbObjectError + 1, , "Open a saved part copy."
    If inputModel.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "Active document must be a part copy."
    inputPath = inputModel.GetPathName
    If Len(inputPath) = 0 Then Err.Raise vbObjectError + 3, , "The part must already have a saved path."
    If inputModel.GetSaveFlag <> False Then Err.Raise vbObjectError + 4, , "The part has unsaved changes. Save your test copy manually before running."
    If Len(Dir$(inputPath)) = 0 Then Err.Raise vbObjectError + 5, , "Saved input file was not found."
    inputFileName = Mid$(inputPath, InStrRev(inputPath, "\") + 1)
    templatePath = DRAWING_TEMPLATE
    If Len(templatePath) = 0 Then templatePath = swApp.GetUserPreferenceStringValue(swDefaultTemplateDrawing)
    If Len(templatePath) = 0 Then Err.Raise vbObjectError + 6, , "Set DRAWING_TEMPLATE or configure the default drawing template."
    If LCase$(Right$(templatePath, 7)) <> ".drwdot" Then Err.Raise vbObjectError + 7, , "Drawing template must be a .drwdot file."
    If Len(Dir$(templatePath)) = 0 Then Err.Raise vbObjectError + 8, , "Drawing template file was not found."
    Debug.Print "Input part: "; inputPath
    Debug.Print "Drawing template: "; templatePath

    ' Width/height apply only to custom paper; this recipe requests standard A size.
    Set drawingModel = swApp.NewDocument(templatePath, swDwgPaperAsize, 0#, 0#)
    If drawingModel Is Nothing Then Err.Raise vbObjectError + 9, , "NewDocument returned Nothing."
    outputTitle = drawingModel.GetTitle
    If drawingModel.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 10, , "Template did not create a drawing document."
    Set drawing = drawingModel
    ' Standard third-angle layout avoids localized '*Front' or feature/view names.
    If drawing.Create3rdAngleViews2(inputPath) = False Then Err.Raise vbObjectError + 11, , "Create3rdAngleViews2 failed."
    If drawingModel.EditRebuild3 = False Then Err.Raise vbObjectError + 12, , "New drawing rebuild failed."
    sheets = drawing.GetViews
    If IsArray(sheets) = False Then Err.Raise vbObjectError + 13, , "Drawing returned no sheet/view arrays."
    For sheetIndex = LBound(sheets) To UBound(sheets)
        views = sheets(sheetIndex)
        If IsArray(views) <> False Then
            ' First view in each sheet array represents the sheet, not a model view.
            For viewIndex = LBound(views) + 1 To UBound(views)
                Set view = views(viewIndex)
                If Not view Is Nothing Then
                    referencedPath = view.GetReferencedModelName
                    Debug.Print "View="; view.GetName2; " referenced model="; referencedPath
                    If StrComp(referencedPath, inputPath, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 14, , "View model path does not match the input part."
                    modelViewCount = modelViewCount + 1
                End If
            Next viewIndex
        End If
    Next sheetIndex
    If modelViewCount < 3 Then Err.Raise vbObjectError + 15, , "Readback found fewer than three model views."
    If Len(drawingModel.GetPathName) <> 0 Then Err.Raise vbObjectError + 16, , "Expected a new unsaved drawing."
    Debug.Print "Created UNSAVED drawing: "; outputTitle; " model views="; modelViewCount
    succeeded = True
    GoTo Cleanup
Failed:
    failureMessage = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Run restoration only after leaving the active error handler.
    On Error Resume Next
    Err.Clear
    If Not swApp Is Nothing Then
        If Not drawingModel Is Nothing Then
            restored = RestoreInput(swApp, inputFileName, inputPath)
            If Err.Number <> 0 Then
                restored = False
                failureMessage = failureMessage & vbCrLf & "Restoration call failed: " & Err.Description
                Err.Clear
            End If
            If restored = False Then
                succeeded = False
                failureMessage = failureMessage & vbCrLf & "Original document activation could not be verified."
            End If
        End If
    End If
    If Err.Number <> 0 Then
        succeeded = False
        failureMessage = failureMessage & vbCrLf & "Cleanup failed: " & Err.Description
        Err.Clear
    End If
    If succeeded <> False Then
        Debug.Print "RESULT create_three_views PASS; model_views="; modelViewCount
        If SHOW_MESSAGES <> False Then MsgBox "Created unsaved drawing: " & outputTitle & vbCrLf & _
            "At least three model views reference the input part. Original part is active; neither document was saved.", vbInformation
    Else
        Debug.Print "RESULT create_three_views FAIL: "; failureMessage
        If Len(outputTitle) > 0 Then Debug.Print "Unsaved output may remain open: "; outputTitle
        If SHOW_MESSAGES <> False Then MsgBox failureMessage & vbCrLf & "No save/close operation was requested; inspect any new drawing.", vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Function RestoreInput(ByVal swApp As SldWorks.SldWorks, ByVal inputFileName As String, ByVal inputPath As String) As Boolean
    On Error GoTo Failed
    Dim activated As SldWorks.ModelDoc2
    Dim errors As Long
    Dim restoreFailure As String
    Set activated = swApp.ActivateDoc3(inputFileName, False, swDontRebuildActiveDoc, errors)
    Debug.Print "Restore activation status="; errors
    If activated Is Nothing Then Exit Function
    If activated.GetType <> swDocPART Then Exit Function
    If StrComp(activated.GetPathName, inputPath, vbTextCompare) <> 0 Then Exit Function
    RestoreInput = True
    Exit Function
Failed:
    restoreFailure = CStr(Err.Number) & ": " & Err.Description
    Resume RestoreFailed
RestoreFailed:
    On Error Resume Next
    Debug.Print "RestoreInput FAILED: "; restoreFailure
    RestoreInput = False
    On Error GoTo 0
End Function
