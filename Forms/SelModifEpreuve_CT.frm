VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelModifEpreuve_CT 
   Caption         =   "S�lectionner l'Epreuve � Modifier"
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
            Dim EpreuveModif_CT As Long
            EpreuveModif_CT = r + 1
            Sheets("R�glages R�gate").Cells(30, "B").Value = EpreuveModif_CT
            ModifEpreuve_CT.Show
            EpreuveModif_CT = 0
            Sheets("R�glages R�gate").Cells(30, "B").Value = 0
            Unload Me
        End If
    Else
        MsgBox "Veuillez s�lectionner une �preuve � modifier.", vbExclamation, "Aucune Epreuve S�lectionn�e"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Feuille � S�lectionner
    Sheets("Stockage Epreuves CT").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauCourses.RowSource = "A1:E200"
    TableauCourses.ColumnWidths = "80;200;500;80;60"
    Sheets("Gestion CrewTimer").Select
End Sub


