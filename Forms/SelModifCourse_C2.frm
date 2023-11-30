VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelModifCourse_C2 
   Caption         =   "Sélectionner la Course à Modifier"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8680.001
   OleObjectBlob   =   "SelModifCourse_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelModifCourse_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Modifier_Click()
    ' Vérifiez si une ligne est sélectionnée
    If TableauCourses.ListIndex <> -1 Then
        ' Récupérez l'index de la ligne sélectionnée
        Dim r As Integer
        r = TableauCourses.ListIndex
        ' Vérifiez si la première ligne (entête de colonne) est sélectionnée
        If r = 0 Then
            MsgBox "La première ligne (entête de colonne) ne peut pas être modifiée.", vbExclamation, "Erreur de Modification"
        Else
            ' Ouvrez le UserForm de modification en passant la ligne sélectionnée en paramètre
            Dim CourseModif_C2 As Long
            CourseModif_C2 = r + 1
            Sheets("Réglages Régate").Cells(27, "B").value = CourseModif_C2
            If Sheets("Programme des Courses C2").Cells(CourseModif_C2, 52).value = "Indiv" Then ModifCourse_C2_Indiv.Show
            If Sheets("Programme des Courses C2").Cells(CourseModif_C2, 52).value = "Relais" Then ModifCourse_C2_Relais.Show
            If Sheets("Programme des Courses C2").Cells(CourseModif_C2, 52).value = "Equipe" Then ModifCourse_C2_Equipes.Show
            CourseModif_C2 = 0
            Sheets("Réglages Régate").Cells(27, "B").value = 0
            Unload Me
        End If
    Else
        MsgBox "Veuillez sélectionner une course à modifier.", vbExclamation, "Aucune Course Sélectionnée"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Feuille à Sélectionner
    Sheets("Programme des Courses C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties).
    TableauCourses.RowSource = "A1:I200"
    TableauCourses.ColumnWidths = "60;40;45;0;140;60;0;0;0"
    Sheets("Gestion Concept2").Select
End Sub


