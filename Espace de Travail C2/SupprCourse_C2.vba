Private Sub Annuler_Click()
 Unload Me
End Sub

Private Sub Supprimer_Click()
 Dim strng As String
    Dim lCol As Long, lRow As Long
    If MsgBox("Êtes-vous sûr de vouloir supprimer cette course ?", vbYesNo + vbQuestion, "Confirmation de Suppression") = vbYes Then
    Sheets("Programme des Courses C2").Select
    For r = 0 To TableauCourses.ListCount - 1
        If TableauCourses.Selected(r) Then
        Rows(r + 1).Delete Shift:=xlUp
    End If
    Next
    Sheets("Gestion Concept2").Select
    MsgBox "La course à été supprimée avec succès !", vbInformation, "Confirmation Suppression"
    End If
    Unload Me
End Sub
Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Programme des Courses C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.)
    TableauCourses.RowSource = "A1:I200"
    TableauCourses.ColumnWidths = "60;40;45;0;140;60;0;0;0"
    Sheets("Gestion Concept2").Select
End Sub
