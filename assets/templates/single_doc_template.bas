Attribute VB_Name = "Main"
' ============================================================================
' Macro:    <NAME>
' Purpose:  <ONE LINE: what it does>
' Works on: active <part / assembly / drawing>
' Requires: SOLIDWORKS 2018+, references: SOLIDWORKS Type Library,
'           SOLIDWORKS Constant type library
' Usage:    open the document, Tools > Macro > Run
' ============================================================================
Option Explicit

' ---- Configuration ----------------------------------------------------------
Const MACRO_TITLE As String = "<NAME>"

' ---- Globals ----------------------------------------------------------------
Dim swApp As SldWorks.SldWorks

Sub main()
    Dim swModel As SldWorks.ModelDoc2

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        swApp.SendMsgToUser2 "Open a document first.", swMbWarning, swMbOk
        Exit Sub
    End If

    ' Restrict to the supported document type(s)
    If swModel.GetType <> swDocPART Then
        swApp.SendMsgToUser2 "This macro works on parts only.", swMbWarning, swMbOk
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ProcessModel swModel

    swApp.SendMsgToUser2 MACRO_TITLE & ": done.", swMbInformation, swMbOk
    Exit Sub

ErrHandler:
    swApp.SendMsgToUser2 MACRO_TITLE & " - error " & Err.Number & ": " & Err.Description, swMbStop, swMbOk
End Sub

Private Sub ProcessModel(swModel As SldWorks.ModelDoc2)
    ' TODO: the actual work
End Sub
