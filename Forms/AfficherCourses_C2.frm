VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AfficherCourses_C2 
   Caption         =   "Gestion du Programme des Courses"
   ClientHeight    =   7840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13780
   OleObjectBlob   =   "AfficherCourses_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AfficherCourses_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AjoutCourses_Equipes_Click()
    CreationCourse_C2_Equipes.Show
End Sub

Private Sub AjoutCourses_Indiv_Click()
    CreationCourse_C2_Indiv.Show
End Sub

Private Sub AjoutCourses_Relais_Click()
    CreationCourse_C2_Relais.Show
End Sub

Private Sub SuppressionCourse_Click()
    SupprCourse_C2.Show
End Sub
Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Programme des Courses C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "A1:BD200"
    TableauPrgCourses.ColumnWidths = "60;40;50;0;140;250;0;0;140;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;60;60;60;60;60;0;60"
    Sheets("Gestion Concept2").Select
    Me.AjoutCourses_Indiv.Caption = "Créer une Course" & vbCrLf & "Individuelle"
    Me.AjoutCourses_Equipes.Caption = "Créer une Course" & vbCrLf & "par Equipes"
    Me.AjoutCourses_Relais.Caption = "Créer une Course" & vbCrLf & "en Relai"
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub

