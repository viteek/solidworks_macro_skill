Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks, model As SldWorks.ModelDoc2
    Dim manager As SldWorks.CustomPropertyManager
    Dim errors As Long, warnings As Long, addStatus As Long, ok As Boolean
    On Error GoTo Failed
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    Set manager = model.Extension.CustomPropertyManager("")
    addStatus = manager.Add3("AuditRunnerProbe", swCustomInfoText, "disposable test only", swCustomPropertyReplaceValue)
    model.SetSaveFlag
    ok = model.Save3(swSaveAsOptions_Silent, errors, warnings)
    Debug.Print "Private input deliberately saved; ok=" & CStr(ok) & "; errors=" & errors
    Exit Sub
Failed:
    Debug.Print "Probe error=" & CStr(Err.Number) & "; " & Err.Description
    Resume Finished
Finished:
End Sub
