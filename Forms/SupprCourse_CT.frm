VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SupprCourse_CT 
   Caption         =   "Suppression d'une Course"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   8680.001
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
    auMoinsUneSelection = False
    If MsgBox("Êtes-vous sûr de vouloir supprimer cette course ?", vbYesNo + vbQuestion, "Confirmation de Suppression") = vbYes Then
    Sheets("Programme des Courses CT").Select
    For r = 0 To TableauCourses.ListCount - 1
            If TableauCourses.Selected(r) Then
                auMoinsUneSelection = True
                Exit For
            End If
        Next r
        
        ' Si aucune ligne n'est sélectionnée, affichez un message d'erreur et quittez la procédure
        If Not auMoinsUneSelection Then
            MsgBox "Veuillez sélectionner au moins une ligne à supprimer.", vbExclamation, "Erreur de Suppression"
            Exit Sub
        End If
    For r = 0 To TableauCourses.ListCount - 1
        If TableauCourses.Selected(r) Then
            If r = 0 Then
            MsgBox "La première ligne (entête de colonne) ne peut pas être supprimée.", vbExclamation, "Erreur de Suppression"
            Exit Sub
        Else
            Rows(r + 1).Delete Shift:=xlUp
        End If
    End If
    Next
    Sheets("Gestion CrewTimer").Select
    MsgBox "La course à été supprimée avec succès !", vbInformation, "Confirmation Suppression"
    End If
    Unload Me
End Sub
Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Sheets("Programme des Courses CT").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "A1:I200"
    TableauCourses.ColumnWidths = "60;40;45;0;140;60;0;0;0"
    Sheets("Gestion CrewTimer").Select
End Sub
