VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifEpreuve_CT 
   Caption         =   "Modification d'une Epreuve"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780.001
   OleObjectBlob   =   "ModifEpreuve_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifEpreuve_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CourseAModifier_CT As Long
Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Sauvegarder_Click()
            Dim CategSel As String
            Dim i As Long
            For i = 0 To Categ.ListCount - 1
            If Categ.Selected(i) Then
                Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, 6 + i).value = Categ.List(i)
                CategSel = CategSel & Categ.List(i) & " / "
            End If
            Next i
            CategSel = Left(CategSel, Len(CategSel) - 3)
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "A").value = CodeEpreuve.Text
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "B").value = Nom_Epreuve.Text
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "C").value = CategSel
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "D").value = Taille.Text
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "E").value = Barreur.Text
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "AV").value = CodeEpreuve.Text
            Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, "F").value = TypePart.Text
            Sheets("Stockage Epreuves CT").Select
            Range("A1").Select
            Sheets("Gestion CrewTimer").Select
          MsgBox "L'�preuve � �t� modifi�e avec succ�s !", vbOKOnly + vbInformation, "Epreuve Modifi�e"
      Unload Me
End Sub
Private Sub UserForm_Initialize()
CourseAModifier_CT = Sheets("R�glages R�gate").Cells(30, "B").value
    Me.Categ.AddItem ("Jeune (J10)")
    Me.Categ.AddItem ("Jeune (J11)")
    Me.Categ.AddItem ("Jeune (J12)")
    Me.Categ.AddItem ("Jeune (J13)")
    Me.Categ.AddItem ("Jeune (J14)")
    Me.Categ.AddItem ("Junior (J15)")
    Me.Categ.AddItem ("Junior (J16)")
    Me.Categ.AddItem ("Junior (J17)")
    Me.Categ.AddItem ("Junior (J18)")
    Me.Categ.AddItem ("S�nior -23")
    Me.Categ.AddItem ("S�nior")
    Me.Taille.AddItem ("1")
    Me.Taille.AddItem ("2")
    Me.Taille.AddItem ("3")
    Me.Taille.AddItem ("4")
    Me.Taille.AddItem ("5")
    Me.Taille.AddItem ("6")
    Me.Taille.AddItem ("7")
    Me.Taille.AddItem ("8")
    Me.Barreur.AddItem ("Oui")
    Me.Barreur.AddItem ("Non")
    Me.TypePart.AddItem ("Homme")
    Me.TypePart.AddItem ("Femme")
    Me.TypePart.AddItem ("Mixte")
    ' Ajoutez ici le code n�cessaire pour charger les donn�es de la ligne sp�cifi�e dans les contr�les de l'UserForm
    ' Par exemple, si vous avez un contr�le TextBoxNomCourse, vous pouvez le remplir comme ceci :
    CodeEpreuve.Text = Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, 1).value
    Nom_Epreuve.Text = Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, 2).value
    Taille.Text = Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, 4).value
    Barreur.Text = Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, 5).value
    TypePart.Text = Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, 6).value
    'Ne Pas R�cup�rer G et D
    ' Ajoutez d'autres lignes similaires pour les autres contr�les que vous souhaitez initialiser
    For colcateg = 8 To 47
    ' V�rifiez si la cellule sp�cifi�e contient une valeur
    If Not IsEmpty(Sheets("Stockage Epreuves CT").Cells(CourseAModifier_CT, colcateg).value) Then
        ' Assurez-vous que l'index est dans la plage valide pour la ListBox
        If colcateg - 8 < Categ.ListCount Then
            ' Ajoutez l'option � la ListBox
            Categ.Selected(colcateg - 8) = True
        End If
    End If
Next colcateg
End Sub


