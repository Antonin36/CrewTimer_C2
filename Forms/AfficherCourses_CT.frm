VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AfficherCourses_CT 
   Caption         =   "Gestion du Programme des Courses"
   ClientHeight    =   7848
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13632
   OleObjectBlob   =   "AfficherCourses_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AfficherCourses_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AjoutCourses_Click()
    CreationCourse_CT.Show
End Sub

Private Sub SuppressionCourse_Click()
    SupprCourse_CT.Show
End Sub
Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Programme des Courses").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "A1:F200"
    TableauPrgCourses.ColumnWidths = "60;40;50;0;140;60"
    Sheets("Gestion CrewTimer").Select
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub
