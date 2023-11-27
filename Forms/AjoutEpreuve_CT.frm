VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AjoutEpreuve_CT 
   Caption         =   "Ajout d'une Epreuve"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780.001
   OleObjectBlob   =   "AjoutEpreuve_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AjoutEpreuve_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Sauvegarder_Click()
            Dim LastRow As Long
            LastRow = Sheets("Stockage Epreuves CT").Cells(Rows.Count, "A").End(xlUp).Row + 1
            Dim CategSel As String
            Dim i As Long
            For i = 0 To Categ.ListCount - 1
            If Categ.Selected(i) Then
                Sheets("Stockage Epreuves CT").Cells(LastRow, 8 + i).Value = Categ.List(i)
                CategSel = CategSel & Categ.List(i) & " / "
            End If
            Next i
            CategSel = Left(CategSel, Len(CategSel) - 3)
            Sheets("Stockage Epreuves CT").Cells(LastRow, "A").Value = CodeEpreuve.Text
            Sheets("Stockage Epreuves CT").Cells(LastRow, "B").Value = Nom_Epreuve.Text
            Sheets("Stockage Epreuves CT").Cells(LastRow, "C").Value = CategSel
            Sheets("Stockage Epreuves CT").Cells(LastRow, "D").Value = Taille.Text
            Sheets("Stockage Epreuves CT").Cells(LastRow, "E").Value = Barreur.Text
            Sheets("Stockage Epreuves CT").Cells(LastRow, "AV").Value = CodeEpreuve.Text
            Sheets("Stockage Epreuves CT").Cells(LastRow, "F").Value = TypePart.Text
            Sheets("Stockage Epreuves CT").Select
            Range("A1").Select
            Sheets("Gestion CrewTimer").Select
          MsgBox "L'épreuve à été créée avec succès !", vbOKOnly + vbInformation, "Epreuve Créée"
      Unload Me
End Sub
Private Sub UserForm_Initialize()
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
    Me.Barreur.AddItem ("Oui")
    Me.Barreur.AddItem ("Non")
    Me.TypePart.AddItem ("Homme")
    Me.TypePart.AddItem ("Femme")
    Me.TypePart.AddItem ("Mixte")
End Sub

