Private Sub AjoutCourses_Click()
    CreationCourse_C2.Show
End Sub

Private Sub SuppressionCourse_Click()
    SupprCourse_C2.Show
End Sub
Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Programme des Courses C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "A1:F200"
    TableauPrgCourses.ColumnWidths = "60;40;50;0;140;60"
    Sheets("Gestion Concept2").Select
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub