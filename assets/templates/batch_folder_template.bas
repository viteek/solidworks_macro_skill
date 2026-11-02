Attribute VB_Name = "Main"
Option Explicit
' Process every file in a folder. SOLIDWORKS 2024 SP5 verified.
' Audit by default (SAVE_CHANGES=False: files opened read-only, only InspectModel runs).
' Skips documents already loaded and same-named files (SOLIDWORKS can't hold two), checks the returned path,
' closes only what it opened, appends a CSV log per file. No PDM operations.
Const SOURCE_FOLDER As String = ""              ' empty = folder picker (only when SHOW_MESSAGES = True)
Const FILE_MASK As String = "*.sldprt"
Const INCLUDE_SUBFOLDERS As Boolean = False
Const SAVE_CHANGES As Boolean = False
Const LOG_NAME As String = "macro_log.csv"
Const SHOW_MESSAGES As Boolean = False          ' False = unattended (Batch+/night run): no dialog on any path
Dim swApp As SldWorks.SldWorks
Dim logPath As String
Dim logHealthy As Boolean

Sub main()
    Dim files As Collection, initialActive As SldWorks.ModelDoc2
    Dim previousCommand As Boolean, commandChanged As Boolean
    Dim folder As String, i As Long, activateError As Long
    Dim failure As String
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    Set initialActive = swApp.ActiveDoc
    previousCommand = (swApp.CommandInProgress <> False)
    folder = SOURCE_FOLDER
    If folder = "" And SHOW_MESSAGES Then folder = PickFolder("Select the folder to process")
    If folder = "" Then Err.Raise vbObjectError + 1, , "Set SOURCE_FOLDER (or SHOW_MESSAGES=True to pick it)"
    folder = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(folder)
    logPath = folder & "\" & LOG_NAME
    logHealthy = True
    If Not TryLog("File;Status;Message") Then Err.Raise vbObjectError + 2, , "Cannot append batch log"
    Set files = New Collection
    CollectFiles folder, files
    swApp.CommandInProgress = True
    commandChanged = True
    For i = 1 To files.Count
        ProcessFile CStr(files(i))
        If Not logHealthy Then Exit For
    Next i
    GoTo Cleanup
Failed:
    failure = Err.Number & ": " & Err.Description
    If logPath <> "" Then logHealthy = TryLog(";FATAL;" & Csv(failure))
    Resume Cleanup   ' leave error-handling mode so the best-effort cleanup below can trap its own errors
Cleanup:
    ' Every cleanup operation is best-effort and cannot bypass the remaining restoration.
    On Error Resume Next
    If commandChanged Then swApp.CommandInProgress = previousCommand
    If Not initialActive Is Nothing Then swApp.ActivateDoc3 initialActive.GetTitle, False, swDontRebuildActiveDoc, activateError
    If Err.Number <> 0 Then Debug.Print "Cleanup: " & Err.Description
    On Error GoTo 0
    If failure <> "" Or Not logHealthy Then
        Debug.Print "Batch failed: " & failure & "; log healthy=" & logHealthy
        If SHOW_MESSAGES Then MsgBox "Batch failed. Inspect the log or Immediate Window.", vbExclamation
    ElseIf SHOW_MESSAGES Then
        MsgBox "Batch complete. Inspect " & logPath, vbInformation
    End If
End Sub

Private Sub ProcessFile(filePath As String)
    Dim model As SldWorks.ModelDoc2
    Dim errors As Long, warnings As Long, openedByMacro As Boolean
    Dim message As String, status As String, closeFailure As String
    Dim stillLoaded As Boolean
    On Error GoTo Failed
    status = "ERROR"
    If IsLoaded(filePath) Then
        status = "SKIP": message = "Document already loaded; user changes preserved"
        GoTo Cleanup
    End If
    Dim options As Long
    options = swOpenDocOptions_Silent
    If SAVE_CHANGES = False Then options = options Or swOpenDocOptions_ReadOnly
    Set model = swApp.OpenDoc6(filePath, DocTypeFromPath(filePath), options, "", errors, warnings)
    If model Is Nothing Then Err.Raise vbObjectError + 10, , "Open failed; code=" & errors
    If SamePath(model.GetPathName, filePath) = False Then
        ' Do not claim ownership or close a different document returned by SOLIDWORKS.
        Set model = Nothing
        Err.Raise vbObjectError + 11, , "Returned path differs from requested path"
    End If
    openedByMacro = True
    If errors <> 0 Then Err.Raise vbObjectError + 12, , "Open returned load errors=" & errors
    If SAVE_CHANGES Then
        ApplyChanges model
        If model.Save3(swSaveAsOptions_Silent, errors, warnings) = False Then Err.Raise vbObjectError + 13, , "Save failed; code=" & errors
        If errors <> 0 Then Err.Raise vbObjectError + 14, , "Save returned errors=" & errors
        status = "SAVED"
    Else
        InspectModel model
        status = "AUDIT"
    End If
    message = "warnings=" & warnings
    GoTo Cleanup
Failed:
    message = Err.Number & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    If openedByMacro And Not model Is Nothing Then
        ' Only a verified owned document is closed; audit does not run ApplyChanges.
        Err.Clear
        swApp.CloseDoc model.GetTitle
        If Err.Number <> 0 Then
            closeFailure = Err.Number & ": " & Err.Description
        Else
            ' CloseDoc has no return value; confirm the document is no longer loaded.
            Err.Clear
            stillLoaded = IsLoaded(filePath)
            If Err.Number <> 0 Then
                closeFailure = "Close verification failed: " & Err.Number & ": " & Err.Description
            ElseIf stillLoaded Then
                closeFailure = "Document remains loaded after CloseDoc"
            End If
        End If
    End If
    On Error GoTo 0
    If closeFailure <> "" Then
        message = "Processing status=" & status & "; " & message & "; Close failed or unconfirmed: " & closeFailure
        status = "ERROR"
    End If
    If Not TryLog(Csv(filePath) & ";" & status & ";" & Csv(message)) Then logHealthy = False
End Sub

Private Sub InspectModel(model As SldWorks.ModelDoc2)
    ' Add READ-ONLY checks here. ApplyChanges is never called in audit mode.
    Debug.Print model.GetPathName
End Sub

Private Sub ApplyChanges(model As SldWorks.ModelDoc2)
    ' Replace this guard with the actual task, including API status checks.
    Err.Raise vbObjectError + 20, , "ApplyChanges has not been implemented; nothing saved"
End Sub

Private Function IsLoaded(filePath As String) As Boolean
    Dim documents As Variant, doc As SldWorks.ModelDoc2, i As Long
    documents = swApp.GetDocuments
    If IsEmpty(documents) Then Exit Function
    For i = LBound(documents) To UBound(documents)
        Set doc = documents(i)
        If SamePath(doc.GetPathName, filePath) Then IsLoaded = True: Exit Function
        ' A same-named document from another folder is also a collision; do not open it.
        If LCase(CreateObject("Scripting.FileSystemObject").GetFileName(doc.GetPathName)) = _
           LCase(CreateObject("Scripting.FileSystemObject").GetFileName(filePath)) Then IsLoaded = True: Exit Function
    Next i
End Function

Private Function SamePath(leftPath As String, rightPath As String) As Boolean
    Dim fso As Object
    If leftPath = "" Or rightPath = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(leftPath) And fso.FileExists(rightPath) Then
        ' ShortPath normalizes 8.3 vs long names (C:\Users\MICHAL~1.WIK vs C:\Users\michal.wiktorowicz)
        SamePath = (StrComp(fso.GetFile(leftPath).ShortPath, fso.GetFile(rightPath).ShortPath, vbTextCompare) = 0)
    Else
        SamePath = (StrComp(fso.GetAbsolutePathName(leftPath), fso.GetAbsolutePathName(rightPath), vbTextCompare) = 0)
    End If
End Function

Private Function DocTypeFromPath(path As String) As Long
    Select Case LCase(Mid(path, InStrRev(path, ".") + 1))
        Case "sldprt": DocTypeFromPath = swDocPART
        Case "sldasm": DocTypeFromPath = swDocASSEMBLY
        Case "slddrw": DocTypeFromPath = swDocDRAWING
        Case Else: Err.Raise vbObjectError + 30, , "Unsupported extension"
    End Select
End Function

Private Sub CollectFiles(folder As String, files As Collection)
    Dim fso As Object, current As Object, item As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set current = fso.GetFolder(folder)
    For Each item In current.Files
        If LCase(item.Name) Like LCase(FILE_MASK) Then
            If Left(item.Name, 2) <> "~$" Then files.Add item.Path
        End If
    Next item
    If INCLUDE_SUBFOLDERS Then
        For Each item In current.SubFolders
            ' Do not follow junctions/symlinks out of the selected tree (FILE_ATTRIBUTE_REPARSE_POINT).
            If (item.Attributes And 1024) = 0 Then CollectFiles item.Path, files
        Next item
    End If
End Sub

Private Function Csv(value As String) As String
    Csv = """" & Replace(Replace(Replace(value, """", """"""), vbCr, " "), vbLf, " ") & """"
End Function

Private Function TryLog(line As String) As Boolean
    Dim ff As Integer, isOpen As Boolean, message As String
    On Error GoTo Failed
    ff = FreeFile
    Open logPath For Append As #ff
    isOpen = True
    Print #ff, line
    Close #ff
    TryLog = True
    Exit Function
Failed:
    message = Err.Description
    Resume CloseLog
CloseLog:
    On Error Resume Next
    If isOpen Then Close #ff
    Debug.Print "Log unavailable: " & message
    On Error GoTo 0
End Function

Private Function PickFolder(title As String) As String
    Dim sh As Object, fld As Object
    Set sh = CreateObject("Shell.Application")
    Set fld = sh.BrowseForFolder(0, title, 0)
    If Not fld Is Nothing Then PickFolder = fld.Self.Path
End Function
