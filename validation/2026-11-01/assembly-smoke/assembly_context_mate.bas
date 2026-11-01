Attribute VB_Name = "AssemblyContextMate"
Option Explicit
Private Const SHOW_MESSAGES As Boolean = True

' NOT_LIVE_TESTED. SOLIDWORKS 2019+; references: SOLIDWORKS and swconst type libraries.
' Run only on a disposable assembly copy. Select two planar faces on different
' resolved, top-level components while the assembly itself is active.
' Creates one coincident mate; rebuilds and clears selection; never saves.
' API evidence and negative cases: README.md and ../../references/assembly-mates.md.
Private Const RESULT_PATH As String = ""

Private Sub RunRecipe()
    On Error GoTo Failed
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim assembly As SldWorks.AssemblyDoc
    Dim selections As SldWorks.SelectionMgr
    Dim firstFace As SldWorks.Face2
    Dim secondFace As SldWorks.Face2
    Dim firstComponent As SldWorks.Component2
    Dim secondComponent As SldWorks.Component2
    Dim parent As SldWorks.Component2
    Dim firstEntity As SldWorks.Entity
    Dim secondEntity As SldWorks.Entity
    Dim data As SldWorks.SelectData
    Dim definition As SldWorks.CoincidentMateFeatureData
    Dim created As SldWorks.Feature
    Dim mate As SldWorks.Mate2
    Dim mateEntity As SldWorks.MateEntity2
    Dim component As SldWorks.Component2
    Dim surface As SldWorks.Surface
    Dim featureName As String
    Dim expectedFirst As String
    Dim expectedSecond As String
    Dim readFirst As String
    Dim readSecond As String
    Dim isWarning As Boolean
    Dim errorCode As Long
    Dim i As Long
    Dim resultLog As Object
    Dim outcome As String
    Dim failureText As String
    Dim selectionTouched As Boolean

    outcome = "FAIL"
    Set resultLog = BeginResultLog()
    If Not resultLog Is Nothing Then resultLog.WriteLine "START assembly_context_mate; terminal RESULT line is required"
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open an assembly copy."
    If model.GetType <> swDocASSEMBLY Then Err.Raise vbObjectError + 2, , "Active document must be an assembly."
    Set assembly = model
    Set selections = model.SelectionManager
    If selections.GetSelectedObjectCount2(-1) <> 2 Then Err.Raise vbObjectError + 3, , "Select exactly two planar component faces."
    For i = 1 To 2
        If selections.GetSelectedObjectType3(i, -1) <> swSelFACES Then Err.Raise vbObjectError + 4, , "Both selections must be faces."
    Next i
    Set firstFace = selections.GetSelectedObject6(1, -1)
    Set secondFace = selections.GetSelectedObject6(2, -1)
    Set firstComponent = selections.GetSelectedObjectsComponent4(1, -1)
    Set secondComponent = selections.GetSelectedObjectsComponent4(2, -1)
    If firstComponent Is Nothing Then Err.Raise vbObjectError + 5, , "First face has no assembly component context."
    If secondComponent Is Nothing Then Err.Raise vbObjectError + 6, , "Second face has no assembly component context."
    If firstComponent.Name2 = secondComponent.Name2 Then Err.Raise vbObjectError + 7, , "Select different component instances."
    expectedFirst = firstComponent.Name2
    expectedSecond = secondComponent.Name2
    Set parent = firstComponent.GetParent
    If Not parent Is Nothing Then Err.Raise vbObjectError + 8, , "This recipe requires top-level components."
    Set parent = secondComponent.GetParent
    If Not parent Is Nothing Then Err.Raise vbObjectError + 9, , "This recipe requires top-level components."
    If firstComponent.IsFixed <> False Then
        If secondComponent.IsFixed <> False Then Err.Raise vbObjectError + 10, , "Both components are fixed; float one manually on the test copy."
    End If
    Set surface = firstFace.GetSurface
    If surface Is Nothing Then Err.Raise vbObjectError + 11, , "First face has no surface."
    If surface.IsPlane = False Then Err.Raise vbObjectError + 12, , "First face is not planar."
    Set surface = secondFace.GetSurface
    If surface Is Nothing Then Err.Raise vbObjectError + 13, , "Second face has no surface."
    If surface.IsPlane = False Then Err.Raise vbObjectError + 14, , "Second face is not planar."

    ' These entities originate in the active ASSEMBLY selection list, not a part.
    Set firstEntity = firstFace
    Set secondEntity = secondFace
    selectionTouched = True
    model.ClearSelection2 True
    Set data = selections.CreateSelectData
    data.Mark = 1
    If firstEntity.Select4(False, data) = False Then Err.Raise vbObjectError + 15, , "First marked selection failed."
    If secondEntity.Select4(True, data) = False Then Err.Raise vbObjectError + 16, , "Second marked selection failed."
    Set definition = assembly.CreateMateData(swMateCOINCIDENT)
    If definition Is Nothing Then Err.Raise vbObjectError + 17, , "CreateMateData returned Nothing."
    definition.MateAlignment = swMateAlignCLOSEST
    Set created = assembly.CreateMate(definition)
    If created Is Nothing Then Err.Raise vbObjectError + 18, , "CreateMate returned Nothing; inspect conflicting mates."
    featureName = created.Name
    If model.EditRebuild3 = False Then Err.Raise vbObjectError + 19, , "Rebuild failed; inspect the new mate on the test copy."
    ' Reacquire the feature and mate after rebuild before reading dependent objects.
    Set created = assembly.FeatureByName(featureName)
    If created Is Nothing Then Err.Raise vbObjectError + 20, , "Created feature could not be reacquired."
    errorCode = created.GetErrorCode2(isWarning)
    If errorCode <> 0 Then
        Debug.Print "Mate feature status="; errorCode; " warning="; isWarning
        Err.Raise vbObjectError + 21, , "Mate has a feature error/warning; review What's Wrong."
    End If
    Set mate = created.GetSpecificFeature2
    If mate Is Nothing Then Err.Raise vbObjectError + 22, , "Mate readback failed."
    If mate.Type <> swMateCOINCIDENT Then Err.Raise vbObjectError + 26, , "Readback mate type is not coincident."
    If mate.GetMateEntityCount <> 2 Then Err.Raise vbObjectError + 23, , "Readback did not contain two entities."
    Debug.Print "Created mate: "; created.Name; " type="; mate.Type; " alignment="; mate.Alignment
    For i = 0 To mate.GetMateEntityCount - 1
        Set mateEntity = mate.MateEntity(i)
        If mateEntity Is Nothing Then Err.Raise vbObjectError + 24, , "Missing mate entity during readback."
        Set component = mateEntity.ReferenceComponent
        If component Is Nothing Then Err.Raise vbObjectError + 25, , "Missing component during mate readback."
        If i = 0 Then readFirst = component.Name2 Else readSecond = component.Name2
        Debug.Print "  Entity "; i; " component="; component.Name2
    Next i
    If (readFirst = expectedFirst And readSecond = expectedSecond) Or _
        (readFirst = expectedSecond And readSecond = expectedFirst) Then
        Debug.Print "Component instance readback matched both inputs."
    Else
        Err.Raise vbObjectError + 27, , "Mate entity components differ from selected inputs."
    End If
    outcome = "PASS"
    GoTo Cleanup
Failed:
    failureText = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    ' Resume leaves the active error handler before any protected cleanup runs.
    On Error Resume Next
    Err.Clear
    If selectionTouched <> False Then model.ClearSelection2 True
    If Err.Number <> 0 Then
        outcome = "FAIL"
        failureText = failureText & "; selection cleanup failed: " & Err.Description
        Err.Clear
    End If
    Debug.Print "RESULT assembly_context_mate "; outcome; " feature="; featureName; " components="; readFirst; ","; readSecond
    If Len(failureText) > 0 Then Debug.Print failureText
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "feature=" & featureName & "; components=" & readFirst & "," & readSecond
        resultLog.WriteLine "detail=" & failureText
        resultLog.WriteLine "RESULT assembly_context_mate " & outcome
        resultLog.Close
        If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED: "; Err.Number; " "; Err.Description
    End If
    If Len(RESULT_PATH) = 0 Then
        If outcome = "PASS" Then
            If SHOW_MESSAGES <> False Then MsgBox "Mate created and read back. Review the assembly; nothing was saved.", vbInformation
        Else
            If SHOW_MESSAGES <> False Then MsgBox failureText & vbCrLf & "A created mate may remain. Inspect the disposable copy; nothing was saved.", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub

Private Function BeginResultLog() As Object
    Dim files As Object
    If Len(RESULT_PATH) = 0 Then Exit Function
    Set files = CreateObject("Scripting.FileSystemObject")
    ' overwrite=False is create-only, including a race with another writer.
    Set BeginResultLog = files.CreateTextFile(RESULT_PATH, False, True)
End Function

Public Sub main()
    Dim fixtureModel As SldWorks.ModelDoc2
    Dim fixtureAssembly As SldWorks.AssemblyDoc
    Dim fixtureComponents As Variant, fixtureBodies As Variant, fixtureFaces As Variant
    Dim fixtureComponent As SldWorks.Component2
    Dim fixturePart As SldWorks.PartDoc
    Dim fixtureBody As SldWorks.Body2
    Dim fixtureEntity As SldWorks.Entity
    Dim n As Long
    Set fixtureModel = Application.SldWorks.ActiveDoc
    Set fixtureAssembly = fixtureModel
    fixtureComponents = fixtureAssembly.GetComponents(True)
    fixtureModel.ClearSelection2 True
    For n = 0 To 1
        Set fixtureComponent = fixtureComponents(n)
        Set fixturePart = fixtureComponent.GetModelDoc2
        fixtureBodies = fixturePart.GetBodies2(swSolidBody, False)
        Set fixtureBody = fixtureBodies(0)
        fixtureFaces = fixtureBody.GetFaces
        Set fixtureEntity = fixtureComponent.GetCorrespondingEntity(fixtureFaces(0))
        If fixtureEntity.Select4(True, Nothing) = False Then Err.Raise vbObjectError + 501, , "Assembly fixture selection failed"
    Next n
    RunRecipe
End Sub
