VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestEngagements_C2 
   Caption         =   "Gestion des Engagements"
   ClientHeight    =   8140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17980
   OleObjectBlob   =   "GestEngagements_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestEngagements_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreerEngagement_Click()
    AjoutEngagement_C2.Show
End Sub

Private Sub ModifierEngagement_Click()
    SelModifEngagement_C2.Show
End Sub

Private Sub SupprEngagement_Click()
    SupprEngagement_C2.Show
End Sub

Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Import GOAL C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "C1:G1000"
    TableauPrgCourses.ColumnWidths = "150;200;400;150;150"
    Sheets("Gestion Concept2").Select
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub



