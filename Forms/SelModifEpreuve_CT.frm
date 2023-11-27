VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelModifEpreuve_CT 
   Caption         =   "Sélectionner l'Epreuve à Modifier"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8680.001
   OleObjectBlob   =   "SelModifEpreuve_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelModifEpreuve_CT"
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
            Dim EpreuveModif_CT As Long
            EpreuveModif_CT = r + 1
            Sheets("Réglages Régate").Cells(30, "B").Value = EpreuveModif_CT
            ModifEpreuve_CT.Show
            EpreuveModif_CT = 0
            Sheets("Réglages Régate").Cells(30, "B").Value = 0
            Unload Me
        End If
    Else
        MsgBox "Veuillez sélectionner une épreuve à modifier.", vbExclamation, "Aucune Epreuve Sélectionnée"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Feuille à Sélectionner
    Sheets("Stockage Epreuves CT").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "A1:E200"
    TableauCourses.ColumnWidths = "80;200;500;80;60"
    Sheets("Gestion CrewTimer").Select
End Sub


