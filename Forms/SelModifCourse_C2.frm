VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelModifCourse_C2 
   Caption         =   "S�lectionner la Course � Modifier"
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
    ' V�rifiez si une ligne est s�lectionn�e
    If TableauCourses.ListIndex <> -1 Then
        ' R�cup�rez l'index de la ligne s�lectionn�e
        Dim r As Integer
        r = TableauCourses.ListIndex
        ' V�rifiez si la premi�re ligne (ent�te de colonne) est s�lectionn�e
        If r = 0 Then
            MsgBox "La premi�re ligne (ent�te de colonne) ne peut pas �tre modifi�e.", vbExclamation, "Erreur de Modification"
        Else
            ' Ouvrez le UserForm de modification en passant la ligne s�lectionn�e en param�tre
            Dim CourseModif_C2 As Long
            CourseModif_C2 = r + 1
            Sheets("R�glages R�gate").Cells(27, "B").value = CourseModif_C2
            If Sheets("Programme des Courses C2").Cells(CourseModif_C2, 52).value = "Indiv" Then ModifCourse_C2_Indiv.Show
            If Sheets("Programme des Courses C2").Cells(CourseModif_C2, 52).value = "Relais" Then ModifCourse_C2_Relais.Show
            If Sheets("Programme des Courses C2").Cells(CourseModif_C2, 52).value = "Equipe" Then ModifCourse_C2_Equipes.Show
            CourseModif_C2 = 0
            Sheets("R�glages R�gate").Cells(27, "B").value = 0
            Unload Me
        End If
    Else
        MsgBox "Veuillez s�lectionner une course � modifier.", vbExclamation, "Aucune Course S�lectionn�e"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Feuille � S�lectionner
    Sheets("Programme des Courses C2").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties).
    TableauCourses.RowSource = "A1:I200"
    TableauCourses.ColumnWidths = "60;40;45;0;140;60;0;0;0"
    Sheets("Gestion Concept2").Select
End Sub


