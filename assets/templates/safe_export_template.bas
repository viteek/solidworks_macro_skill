Attribute VB_Name = "Main"
Option Explicit
' Complete ACTIVE-document exporter (DXF/DWG flat pattern, PDF, STEP/X_T/IGES/STL); SOLIDWORKS 2024 SP5 verified.
' Existing output is never deleted before a successful export: stage -> check -> publish (+ backup).
' For many files/assembly components combine with batch_folder / assembly_traverse (activate each part first).
Const OUTPUT_PATH As String = ""
Const CONFIGURATION_NAME As String = ""  ' empty = active configuration
Const ALLOW_OVERWRITE As Boolean = False
Const EXPORT_BEND_LINES As Boolean = True  ' measured SW2024 SP5: bend lines need bits 4 AND 8 (4 alone exports none)
Const EXPORT_SKETCHES As Boolean = False   ' other visible flat-pattern sketches (bit 8 is also set by EXPORT_BEND_LINES)
Const SHOW_MESSAGES As Boolean = True
Const RESULT_PATH As String = ""

Sub main()
    Dim swApp As SldWorks.SldWorks, model As SldWorks.ModelDoc2
    Dim selection As SldWorks.SelectionMgr, selectionSuspended As Boolean
    Dim originalConfiguration As String, configurationChanged As Boolean
    Dim stagedPath As String, backupPath As String, message As String, failure As String
    Dim cleanupFailed As Boolean
    Dim status As String, fso As Object, extension As String
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a document first"
    If model.GetPathName = "" Then Err.Raise vbObjectError + 2, , "Save the source document first"
    If OUTPUT_PATH = "" Then Err.Raise vbObjectError + 3, , "Set OUTPUT_PATH"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fso.GetParentFolderName(OUTPUT_PATH)) Then Err.Raise vbObjectError + 4, , "Output directory does not exist"
    If fso.FileExists(OUTPUT_PATH) And ALLOW_OVERWRITE = False Then
        status = "SKIP_EXISTING": message = "Existing output preserved": GoTo Cleanup
    End If
    extension = LCase(fso.GetExtensionName(OUTPUT_PATH))
    If Len(OUTPUT_PATH) > 220 Then Err.Raise vbObjectError + 5, , "Use an output path shorter than 221 characters"
    If model.GetType <> swDocDRAWING Then originalConfiguration = model.ConfigurationManager.ActiveConfiguration.Name
    If model.GetType = swDocDRAWING And CONFIGURATION_NAME <> "" Then Err.Raise vbObjectError + 10, , "Drawings do not have model configurations"
    If CONFIGURATION_NAME <> "" And CONFIGURATION_NAME <> originalConfiguration Then
        configurationChanged = True
        If model.ShowConfiguration2(CONFIGURATION_NAME) = False Then Err.Raise vbObjectError + 6, , "Cannot activate requested configuration"
        If model.ConfigurationManager.ActiveConfiguration.Name <> CONFIGURATION_NAME Then Err.Raise vbObjectError + 7, , "Wrong configuration active"
    End If
    Set selection = model.SelectionManager
    selection.SuspendSelectionList
    selectionSuspended = True
    model.ClearSelection2 True
    ' A same-folder stage keeps rename publication on one volume and protects the old output.
    Do
        stagedPath = fso.BuildPath(fso.GetParentFolderName(OUTPUT_PATH), "swstage-" & fso.GetBaseName(fso.GetTempName) & "." & extension)
    Loop While fso.FileExists(stagedPath)
    ExportToStage swApp, model, stagedPath, extension, message
    If Not fso.FileExists(stagedPath) Then Err.Raise vbObjectError + 8, , "Exporter created no file"
    If fso.GetFile(stagedPath).Size = 0 Then Err.Raise vbObjectError + 9, , "Exporter created an empty file"
    PublishStage stagedPath, OUTPUT_PATH, ALLOW_OVERWRITE, backupPath
    stagedPath = ""
    status = "EXPORTED"
    If backupPath <> "" Then message = message & "; previous output retained at " & backupPath
    GoTo Cleanup
Failed:
    failure = Err.Number & ": " & Err.Description
    status = "ERROR": message = failure
    Resume Cleanup   ' leave error-handling mode, otherwise On Error Resume Next below would not trap
Cleanup:
    On Error Resume Next
    If selectionSuspended Then selection.ResumeSelectionList2 False
    If Err.Number <> 0 Then
        message = message & "; selection restore failed: " & Err.Description
        cleanupFailed = True
    End If
    Err.Clear
    If configurationChanged Then
        If model.ShowConfiguration2(originalConfiguration) = False Then
            message = message & "; configuration restore failed"
            cleanupFailed = True
        End If
    End If
    If Err.Number <> 0 Then
        message = message & "; configuration restore error: " & Err.Description
        cleanupFailed = True
    End If
    Err.Clear
    If stagedPath <> "" Then
        If fso.FileExists(stagedPath) Then fso.DeleteFile stagedPath, False
    End If
    If Err.Number <> 0 Then
        message = message & "; staged file retained: " & stagedPath
        cleanupFailed = True
    End If
    On Error GoTo 0
    If cleanupFailed Then status = "ERROR_CLEANUP"
    WriteResult status, message
    If SHOW_MESSAGES Then MsgBox status & vbCrLf & message, vbInformation
End Sub

Private Sub ExportToStage(swApp As SldWorks.SldWorks, model As SldWorks.ModelDoc2, path As String, extension As String, ByRef message As String)
    Dim errors As Long, warnings As Long
    Dim part As SldWorks.PartDoc, drawing As SldWorks.DrawingDoc, pdf As SldWorks.ExportPdfData
    Dim bodies As Variant, body As SldWorks.Body2, options As Long, sheets As Variant
    Select Case extension
        Case "dxf", "dwg"
            If model.GetType <> swDocPART Then Err.Raise vbObjectError + 20, , "DXF/DWG template supports single-body sheet-metal parts only"
            Set part = model
            bodies = part.GetBodies2(swSolidBody, False)
            If IsEmpty(bodies) Then Err.Raise vbObjectError + 21, , "No solid bodies"
            If UBound(bodies) - LBound(bodies) <> 0 Then Err.Raise vbObjectError + 22, , "Multibody: select each FlatPattern feature and export separately (recipes.md section 3)"
            Set body = bodies(LBound(bodies))
            If body.IsSheetMetal = False Then Err.Raise vbObjectError + 23, , "Source is not sheet metal"
            options = 1  ' flat-pattern geometry (always required)
            ' Documented: 4 = bend lines, 8 = sketches. Measured on SW2024 SP5: bend lines (CENTERX2 lines)
            ' appear only when 8 is set too; mask 5 produced a DXF without bend lines.
            If EXPORT_BEND_LINES Then options = options Or 4 Or 8
            If EXPORT_SKETCHES Then options = options Or 8
            If part.ExportToDWG2(path, model.GetPathName, swExportToDWG_ExportSheetMetal, True, _
                Empty, False, False, options, Empty) = False Then Err.Raise vbObjectError + 24, , "ExportToDWG2 failed"
        Case "pdf"
            If model.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 25, , "PDF template supports drawings only"
            Set drawing = model
            Set pdf = swApp.GetExportFileData(swExportPdfData)
            If pdf Is Nothing Then Err.Raise vbObjectError + 26, , "No PDF export settings"
            sheets = drawing.GetSheetNames
            If IsEmpty(sheets) Then Err.Raise vbObjectError + 27, , "Drawing has no sheets"
            If pdf.SetSheets(swExportData_ExportSpecifiedSheets, sheets) = False Then Err.Raise vbObjectError + 28, , "Cannot select PDF sheets"
            pdf.ViewPdfAfterSaving = False
            If model.Extension.SaveAs(path, swSaveAsCurrentVersion, swSaveAsOptions_Silent Or swSaveAsOptions_Copy, pdf, errors, warnings) = False Then
                Err.Raise vbObjectError + 29, , "PDF save failed; code=" & errors
            End If
        Case "step", "stp", "x_t", "igs", "iges", "stl"
            If model.GetType <> swDocPART And model.GetType <> swDocASSEMBLY Then Err.Raise vbObjectError + 30, , "Model export requires a part or assembly"
            If model.Extension.SaveAs(path, swSaveAsCurrentVersion, swSaveAsOptions_Silent Or swSaveAsOptions_Copy, Nothing, errors, warnings) = False Then
                Err.Raise vbObjectError + 31, , "Model export failed; code=" & errors
            End If
        Case Else
            Err.Raise vbObjectError + 32, , "Unsupported output extension"
    End Select
    If errors <> 0 Then Err.Raise vbObjectError + 33, , "Exporter returned error flags=" & errors
    message = "Export warnings=" & warnings
End Sub

Private Sub PublishStage(stage As String, target As String, overwrite As Boolean, ByRef backup As String)
    Dim fso As Object, failure As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(target) Then
        If overwrite = False Then Err.Raise vbObjectError + 40, , "Output appeared during export; existing output preserved"
        backup = stage & ".previous"
        Name target As backup
    End If
    On Error GoTo Failed
    Name stage As target
    Exit Sub
Failed:
    failure = Err.Description
    Resume Restore
Restore:
    On Error Resume Next
    If backup <> "" And Not fso.FileExists(target) Then Name backup As target
    On Error GoTo 0
    Err.Raise vbObjectError + 41, , "Publication failed: " & failure & "; check original/backup at " & backup
End Sub

Private Sub WriteResult(status As String, message As String)
    Dim ff As Integer, isOpen As Boolean
    On Error GoTo Failed
    If RESULT_PATH = "" Then
        Debug.Print status & ";" & message
        Exit Sub
    End If
    ff = FreeFile
    Open RESULT_PATH For Append As #ff
    isOpen = True
    Print #ff, status & ";" & Replace(Replace(message, vbCr, " "), vbLf, " ")
    Close #ff
    Exit Sub
Failed:
    Debug.Print "Result log failed: " & Err.Description
    Resume CloseLog
CloseLog:
    On Error Resume Next
    If isOpen Then Close #ff
    On Error GoTo 0
End Sub
