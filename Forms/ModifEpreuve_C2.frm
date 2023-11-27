VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifEpreuve_C2 
   Caption         =   "Modification d'une Epreuve"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780.001
   OleObjectBlob   =   "ModifEpreuve_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifEpreuve_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CourseAModifier_C2 As Long
Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Sauvegarder_Click()
            Dim CategSel As String
            Dim i As Long
            For i = 0 To Categ.ListCount - 1
            If Categ.Selected(i) Then
                Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, 8 + i).Value = Categ.List(i)
                CategSel = CategSel & Categ.List(i) & " / "
            End If
            Next i
            CategSel = Left(CategSel, Len(CategSel) - 3)
            Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, "A").Value = CodeEpreuve.Text
            Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, "B").Value = Nom_Epreuve.Text
            Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, "C").Value = CategSel
            Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, "D").Value = Taille.Text
            Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, "AV").Value = CodeEpreuve.Text
            Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, "F").Value = TypePart.Text
            Sheets("Stockage Epreuves C2").Select
            Range("A1").Select
            Sheets("Gestion Concept2").Select
          MsgBox "L'épreuve à été créée avec succès !", vbOKOnly + vbInformation, "Epreuve Créée"
      Unload Me
End Sub
Private Sub UserForm_Initialize()
CourseAModifier_C2 = Sheets("Réglages Régate").Cells(31, "B").Value
    Me.Categ.AddItem ("Jeune (J10)")
    Me.Categ.AddItem ("Jeune (J11)")
    Me.Categ.AddItem ("Jeune (J12)")
    Me.Categ.AddItem ("Jeune (J13)")
    Me.Categ.AddItem ("Jeune (J14)")
    Me.Categ.AddItem ("Junior (J15)")
    Me.Categ.AddItem ("Junior (J16)")
    Me.Categ.AddItem ("Junior (J17)")
    Me.Categ.AddItem ("Junior (J18)")
    Me.Categ.AddItem ("Sénior -23")
    Me.Categ.AddItem ("Sénior")
    Me.Taille.AddItem ("1")
    Me.Taille.AddItem ("2")
    Me.Taille.AddItem ("3")
    Me.Taille.AddItem ("4")
    Me.Taille.AddItem ("5")
    Me.Taille.AddItem ("6")
    Me.Taille.AddItem ("7")
    Me.Taille.AddItem ("8")
    Me.TypePart.AddItem ("Homme")
    Me.TypePart.AddItem ("Femme")
    Me.TypePart.AddItem ("Mixte")
    ' Ajoutez ici le code nécessaire pour charger les données de la ligne spécifiée dans les contrôles de l'UserForm
    ' Par exemple, si vous avez un contrôle TextBoxNomCourse, vous pouvez le remplir comme ceci :
    CodeEpreuve.Text = Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, 1).Value
    Nom_Epreuve.Text = Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, 2).Value
    Taille.Text = Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, 4).Value
    TypePart.Text = Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, 6).Value
    'Ne Pas Récupérer G et D
    For colcateg = 8 To 47
    ' Vérifiez si la cellule spécifiée contient une valeur
    If Not IsEmpty(Sheets("Stockage Epreuves C2").Cells(CourseAModifier_C2, colcateg).Value) Then
        ' Assurez-vous que l'index est dans la plage valide pour la ListBox
        If colcateg - 8 < Categ.ListCount Then
            ' Ajoutez l'option à la ListBox
            Categ.Selected(colcateg - 8) = True
        End If
    End If
Next colcateg
End Sub



