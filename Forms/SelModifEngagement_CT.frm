VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelModifEngagement_CT 
   Caption         =   "S�lectionner l'Engagement � Modifier"
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
    ' V�rifiez si une ligne est s�lectionn�e
    If TableauCourses.ListIndex <> -1 Then
        ' R�cup�rez l'index de la ligne s�lectionn�e
        Dim r As Integer
        r = TableauCourses.ListIndex
        ' V�rifiez si la premi�re ligne (ent�te de colonne) est s�lectionn�e
        If r = 0 Then
            MsgBox "La premi�re ligne (ent�te de colonne) ne peut pas �tre modifi�e.", vbExclamation, "Erreur de Modification"
        Else
            ' Ouvrez le UserForm de modification en passant la ligne s�lectionn�e en param�tre
            Dim EngagementModif_CT As Long
            EngagementModif_CT = r + 1
            Sheets("R�glages R�gate").Cells(30, "D").Value = EngagementModif_CT
            ModifEngagement_CT.Show
            EngagementModif_CT = 0
            Sheets("R�glages R�gate").Cells(30, "D").Value = 0
            Unload Me
        End If
    Else
        MsgBox "Veuillez s�lectionner un engagement � modifier.", vbExclamation, "Aucun Engagement S�lectionn�"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Feuille � S�lectionner
    Sheets("Import GOAL CT").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "C1:G1000"
    TableauCourses.ColumnWidths = "150;200;400;150;150"
    Sheets("Gestion CrewTimer").Select
End Sub



