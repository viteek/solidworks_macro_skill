Attribute VB_Name = "CreateMateReference"
Option Explicit
Private Const SHOW_MESSAGES As Boolean = True

' NOT_LIVE_TESTED. Select one to three planar faces in an active PART copy.
' Inserts an explicit 12-argument mate reference, then reads all three slots.
' Adds a feature and clears selection; does not save or alter configurations.
Private Const RESULT_PATH As String = ""
Private Const REFERENCE_NAME As String = ""

Private Sub RunRecipe()
    On Error GoTo Failed
    Dim swApp As SldWorks.SldWorks
    Dim model As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim selections As SldWorks.SelectionMgr
    Dim faces(0 To 2) As SldWorks.Entity
    Dim face As SldWorks.Face2
    Dim surface As SldWorks.Surface
    Dim existing As SldWorks.Feature
    Dim created As SldWorks.Feature
    Dim reference As SldWorks.MateReference
    Dim entity As Object
    Dim referenceName As String
    Dim selectedCount As Long
    Dim slotCount As Long
    Dim presentCount As Long
    Dim i As Long
    Dim resultLog As Object
    Dim outcome As String
    Dim failureText As String
    Dim selectionTouched As Boolean

    outcome = "FAIL"
    Set resultLog = BeginResultLog()
    If Not resultLog Is Nothing Then resultLog.WriteLine "START mate_reference; terminal RESULT line is required"
    Set swApp = Application.SldWorks
    Set model = swApp.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 1, , "Open a part copy."
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 2, , "This recipe requires an active part."
    Set part = model
    Set selections = model.SelectionManager
    selectedCount = selections.GetSelectedObjectCount2(-1)
    If selectedCount < 1 Or selectedCount > 3 Then Err.Raise vbObjectError + 3, , "Select one, two, or three planar faces in reference priority order."
    For i = 0 To selectedCount - 1
        If selections.GetSelectedObjectType3(i + 1, -1) <> swSelFACES Then Err.Raise vbObjectError + 4, , "Every selection must be a face."
        Set face = selections.GetSelectedObject6(i + 1, -1)
        Set surface = face.GetSurface
        If surface Is Nothing Then Err.Raise vbObjectError + 5, , "Selected face has no surface."
        If surface.IsPlane = False Then Err.Raise vbObjectError + 6, , "Use planar faces for this coincident reference recipe."
        Set faces(i) = face
    Next i
    referenceName = REFERENCE_NAME
    If Len(referenceName) = 0 Then referenceName = Trim$(InputBox("Unique mate reference name", "Create mate reference", "RecipeMateReference"))
    If Len(referenceName) = 0 Then
        outcome = "CANCELLED"
        GoTo Cleanup
    End If
    Set existing = part.FeatureByName(referenceName)
    If Not existing Is Nothing Then Err.Raise vbObjectError + 7, , "A feature already uses that name."
    selectionTouched = True
    model.ClearSelection2 True
    ' Exactly 12 parameters: name; primary(entity/type/alignment/alignAxes);
    ' secondary(entity/type/alignment/alignAxes); tertiary(entity/type/alignment).
    ' There is NO tertiary AlignAxes parameter. Unselected entity slots are Nothing.
    Set created = model.FeatureManager.InsertMateReference2(referenceName, _
        faces(0), swMateReferenceType_Coincident, swMateReferenceAlignment_Any, False, _
        faces(1), swMateReferenceType_Coincident, swMateReferenceAlignment_Any, False, _
        faces(2), swMateReferenceType_Coincident, swMateReferenceAlignment_Any)
    If created Is Nothing Then Err.Raise vbObjectError + 8, , "InsertMateReference2 returned Nothing."
    Set reference = created.GetSpecificFeature2
    If reference Is Nothing Then Err.Raise vbObjectError + 9, , "MateReference readback returned Nothing."
    slotCount = reference.ReferenceEntityCount
    Debug.Print "Mate reference: "; reference.Name; " slots="; slotCount
    ' ReferenceEntityCount describes three slots, not necessarily three populated entities.
    For i = 0 To slotCount - 1
        Set entity = reference.ReferenceEntity2(i)
        If entity Is Nothing Then
            Debug.Print "  Slot "; i; " empty; type="; reference.ReferenceType(i)
        Else
            presentCount = presentCount + 1
            Debug.Print "  Slot "; i; " selection type="; reference.ReferenceEntityType(i); _
                " mate type="; reference.ReferenceType(i); " alignment="; reference.ReferenceAlignment(i)
            If reference.ReferenceType(i) <> swMateReferenceType_Coincident Then Err.Raise vbObjectError + 10, , "Mate reference type readback mismatch."
            If reference.ReferenceAlignment(i) <> swMateReferenceAlignment_Any Then Err.Raise vbObjectError + 11, , "Mate reference alignment readback mismatch."
        End If
    Next i
    If presentCount <> selectedCount Then Err.Raise vbObjectError + 12, , "Populated slot count differs from selected entity count."
    outcome = "PASS"
    GoTo Cleanup
Failed:
    failureText = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    Err.Clear
    If selectionTouched <> False Then model.ClearSelection2 True
    If Err.Number <> 0 Then
        outcome = "FAIL"
        failureText = failureText & "; selection cleanup failed: " & Err.Description
        Err.Clear
    End If
    Debug.Print "RESULT mate_reference "; outcome; " name="; referenceName; " populated="; presentCount; " slots="; slotCount
    If Len(failureText) > 0 Then Debug.Print failureText
    If Not resultLog Is Nothing Then
        resultLog.WriteLine "name=" & referenceName & "; populated=" & CStr(presentCount) & "; slots=" & CStr(slotCount)
        resultLog.WriteLine "detail=" & failureText
        resultLog.WriteLine "RESULT mate_reference " & outcome
        resultLog.Close
        If Err.Number <> 0 Then Debug.Print "RESULT_LOG_FAILED: "; Err.Number; " "; Err.Description
    End If
    If Len(RESULT_PATH) = 0 Then
        If outcome = "PASS" Then
            If SHOW_MESSAGES <> False Then MsgBox "Mate reference created and read back. Nothing was saved.", vbInformation
        ElseIf outcome = "FAIL" Then
            If SHOW_MESSAGES <> False Then MsgBox failureText & vbCrLf & "A created reference may remain. Review the disposable copy.", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub

Private Function BeginResultLog() As Object
    Dim files As Object
    If Len(RESULT_PATH) = 0 Then Exit Function
    Set files = CreateObject("Scripting.FileSystemObject")
    Set BeginResultLog = files.CreateTextFile(RESULT_PATH, False, True)
End Function

Public Sub main()
    Dim fixtureModel As SldWorks.ModelDoc2
    Dim fixturePart As SldWorks.PartDoc
    Dim fixtureBodies As Variant, fixtureFaces As Variant
    Dim fixtureBody As SldWorks.Body2
    Dim fixtureEntity As SldWorks.Entity
    Dim n As Long
    Set fixtureModel = Application.SldWorks.ActiveDoc
    Set fixturePart = fixtureModel
    fixtureBodies = fixturePart.GetBodies2(swSolidBody, False)
    Set fixtureBody = fixtureBodies(0)
    fixtureFaces = fixtureBody.GetFaces
    fixtureModel.ClearSelection2 True
    For n = 0 To 2
        Set fixtureEntity = fixtureFaces(n)
        If fixtureEntity.Select4(True, Nothing) = False Then Err.Raise vbObjectError + 500, , "Fixture selection failed"
    Next n
    RunRecipe
End Sub
