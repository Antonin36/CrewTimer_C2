VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestEpreuves_C2 
   Caption         =   "Gestion des Epreuves"
   ClientHeight    =   8140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17980
   OleObjectBlob   =   "GestEpreuves_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestEpreuves_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreerEpreuve_Click()
    AjoutEpreuve_C2.Show
End Sub

Private Sub ModifierEpreuve_Click()
    SelModifEpreuve_C2.Show
End Sub

Private Sub SupprEpreuve_Click()
    SupprEpreuve_C2.Show
End Sub

Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Stockage Epreuves C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "A1:E200"
    TableauPrgCourses.ColumnWidths = "80;200;500;80;60"
    Sheets("Gestion Concept2").Select
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub


