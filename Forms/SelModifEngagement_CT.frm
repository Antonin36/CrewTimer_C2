VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelModifEngagement_CT 
   Caption         =   "Sélectionner l'Engagement à Modifier"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8680.001
   OleObjectBlob   =   "SelModifEngagement_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelModifEngagement_CT"
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
            Dim EngagementModif_CT As Long
            EngagementModif_CT = r + 1
            Sheets("Réglages Régate").Cells(30, "D").Value = EngagementModif_CT
            ModifEngagement_CT.Show
            EngagementModif_CT = 0
            Sheets("Réglages Régate").Cells(30, "D").Value = 0
            Unload Me
        End If
    Else
        MsgBox "Veuillez sélectionner un engagement à modifier.", vbExclamation, "Aucun Engagement Sélectionné"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Feuille à Sélectionner
    Sheets("Import GOAL CT").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "C1:G1000"
    TableauCourses.ColumnWidths = "150;200;400;150;150"
    Sheets("Gestion CrewTimer").Select
End Sub



