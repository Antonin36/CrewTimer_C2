VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SupprEpreuve_C2 
   Caption         =   "Suppression d'une Epreuve"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8680.001
   OleObjectBlob   =   "SupprEpreuve_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SupprEpreuve_C2"
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
    If MsgBox("�tes-vous s�r de vouloir supprimer cette �preuve ?", vbYesNo + vbQuestion, "Confirmation de Suppression") = vbYes Then
    Sheets("Stockage Epreuves C2").Select
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
    Sheets("Gestion Concept2").Select
    MsgBox "L'�preuve � �t� supprim�e avec succ�s !", vbInformation, "Confirmation Suppression"
    End If
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    ' Feuille � S�lectionner
    Sheets("Stockage Epreuves C2").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "A1:E200"
    TableauCourses.ColumnWidths = "80;200;500;80;60"
    Sheets("Gestion Concept2").Select
End Sub


