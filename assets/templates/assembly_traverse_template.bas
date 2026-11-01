Attribute VB_Name = "Main"
' ============================================================================
' Macro:    <NAME>
' Purpose:  walk all components of the active assembly and process each
'           unique part/subassembly file once (per referenced configuration)
' Requires: SOLIDWORKS 2018+, references: SOLIDWORKS Type Library,
'           SOLIDWORKS Constant type library
' ============================================================================
Option Explicit

' ---- Configuration ----------------------------------------------------------
Const MACRO_TITLE As String = "<NAME>"
Const SKIP_SUPPRESSED As Boolean = True
Const SKIP_EXCLUDED_FROM_BOM As Boolean = True
Const INCLUDE_SUBASSEMBLIES As Boolean = False   ' process subassembly files themselves too

' ---- Globals ----------------------------------------------------------------
Dim swApp As SldWorks.SldWorks
Dim processed As Object      ' Scripting.Dictionary: key = path|config -> instance count
Dim logText As String

Sub main()
    Dim swModel As SldWorks.ModelDoc2
    Dim swAssy As SldWorks.AssemblyDoc
    Dim swRoot As SldWorks.Component2

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        swApp.SendMsgToUser2 "Open an assembly first.", swMbWarning, swMbOk
        Exit Sub
    End If
    If swModel.GetType <> swDocASSEMBLY Then
        swApp.SendMsgToUser2 "This macro works on assemblies only.", swMbWarning, swMbOk
        Exit Sub
    End If

    On Error GoTo ErrHandler

    Set swAssy = swModel
    ' Lightweight components return Nothing from GetModelDoc2
    swAssy.ResolveAllLightWeightComponents False

    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = vbTextCompare
    logText = ""

    Set swRoot = swModel.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    If Not swRoot Is Nothing Then TraverseComponent swRoot

    swApp.SendMsgToUser2 MACRO_TITLE & ": processed " & processed.Count & " unique file(s)." & _
                         IIf(logText <> "", vbCrLf & vbCrLf & logText, ""), swMbInformation, swMbOk
    Exit Sub

ErrHandler:
    swApp.SendMsgToUser2 MACRO_TITLE & " - error " & Err.Number & ": " & Err.Description, swMbStop, swMbOk
End Sub

Private Sub TraverseComponent(swParent As SldWorks.Component2)
    Dim children As Variant
    Dim i As Long
    Dim swComp As SldWorks.Component2
    Dim swCompModel As SldWorks.ModelDoc2
    Dim key As String
    Dim isAssembly As Boolean

    children = swParent.GetChildren
    If IsEmpty(children) Then Exit Sub

    For i = 0 To UBound(children)
        Set swComp = children(i)

        If SKIP_SUPPRESSED And (swComp.IsSuppressed <> False) Then GoTo NextComp
        If SKIP_EXCLUDED_FROM_BOM And (swComp.ExcludeFromBOM <> False) Then GoTo NextComp

        Set swCompModel = swComp.GetModelDoc2
        If swCompModel Is Nothing Then
            logText = logText & "Not loaded: " & swComp.Name2 & vbCrLf
            GoTo NextComp
        End If

        isAssembly = (swCompModel.GetType = swDocASSEMBLY)
        key = swComp.GetPathName & "|" & swComp.ReferencedConfiguration

        If processed.Exists(key) Then
            processed(key) = processed(key) + 1
        ElseIf (Not isAssembly) Or INCLUDE_SUBASSEMBLIES Then
            processed.Add key, 1
            ProcessComponent swComp, swCompModel
        End If

        ' Recurse into subassemblies (children are visible even if the subassembly itself is skipped)
        If isAssembly Then TraverseComponent swComp
NextComp:
    Next i
End Sub

Private Sub ProcessComponent(swComp As SldWorks.Component2, swCompModel As SldWorks.ModelDoc2)
    ' Use swComp.ReferencedConfiguration for configuration-specific data
    ' TODO: the actual work
End Sub
