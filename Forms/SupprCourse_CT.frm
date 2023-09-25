VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SupprCourse_CT 
   Caption         =   "Suppression d'une Course"
   ClientHeight    =   5424
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8640.001
   OleObjectBlob   =   "SupprCourse_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SupprCourse_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Annuler_Click()
 Unload Me
End Sub

Private Sub Supprimer_Click()
 Dim strng As String
    Dim lCol As Long, lRow As Long
    If MsgBox("Êtes-vous sûr de vouloir supprimer cette course ?", vbYesNo + vbQuestion, "Confirmation de Suppression") = vbYes Then
    Sheets("Programme des Courses").Select
    For r = 0 To TableauCourses.ListCount - 1
        If TableauCourses.Selected(r) Then
        Rows(r + 1).Delete Shift:=xlUp
    End If
    Next
    Sheets("Gestion CrewTimer").Select
    MsgBox "La course à été supprimée avec succès !", vbInformation, "Confirmation Suppression"
    End If
    Unload Me
End Sub
Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Programme des Courses").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "A1:I200"
    TableauCourses.ColumnWidths = "60;40;45;0;140;60;0;0;0"
    Sheets("Gestion CrewTimer").Select
End Sub
