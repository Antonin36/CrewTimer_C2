VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SupprEngagement_CT 
   Caption         =   "Suppression d'un Engagement"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8680.001
   OleObjectBlob   =   "SupprEngagement_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SupprEngagement_CT"
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
    If MsgBox("�tes-vous s�r de vouloir supprimer cet engagement ?", vbYesNo + vbQuestion, "Confirmation de Suppression") = vbYes Then
    Sheets("Import GOAL CT").Select
    For r = 0 To TableauCourses.ListCount - 1
            If TableauCourses.Selected(r) Then
                auMoinsUneSelection = True
                Exit For
            End If
        Next r
        
        ' Si aucune ligne n'est s�lectionn�e, affichez un message d'erreur et quittez la proc�dure
        If Not auMoinsUneSelection Then
            MsgBox "Veuillez s�lectionner au moins une ligne � supprimer.", vbExclamation, "Erreur de Suppression"
            Exit Sub
        End If
    For r = 0 To TableauCourses.ListCount - 1
        If TableauCourses.Selected(r) Then
            If r = 0 Then
            MsgBox "La premi�re ligne (ent�te de colonne) ne peut pas �tre supprim�e.", vbExclamation, "Erreur de Suppression"
            Exit Sub
        Else
            Rows(r + 1).Delete Shift:=xlUp
        End If
    End If
    Next
    Sheets("Gestion CrewTimer").Select
    MsgBox "L'engagement � �t� supprim� avec succ�s !", vbInformation, "Confirmation Suppression"
    End If
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    ' Feuille � S�lectionner
    Sheets("Import GOAL CT").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "C1:G1000"
    TableauCourses.ColumnWidths = "150;150;400;150;150"
    Sheets("Gestion CrewTimer").Select
End Sub



