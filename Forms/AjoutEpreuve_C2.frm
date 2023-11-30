VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AjoutEpreuve_C2 
   Caption         =   "Ajout d'une Epreuve"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780.001
   OleObjectBlob   =   "AjoutEpreuve_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AjoutEpreuve_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Sauvegarder_Click()
            Dim LastRow As Long
            LastRow = Sheets("Stockage Epreuves C2").Cells(Rows.Count, "A").End(xlUp).Row + 1
            Dim CategSel As String
            Dim i As Long
            For i = 0 To Categ.ListCount - 1
            If Categ.Selected(i) Then
                Sheets("Stockage Epreuves C2").Cells(LastRow, 8 + i).value = Categ.List(i)
                CategSel = CategSel & Categ.List(i) & " / "
            End If
            Next i
            CategSel = Left(CategSel, Len(CategSel) - 3)
            Sheets("Stockage Epreuves C2").Cells(LastRow, "A").value = CodeEpreuve.Text
            Sheets("Stockage Epreuves C2").Cells(LastRow, "B").value = Nom_Epreuve.Text
            Sheets("Stockage Epreuves C2").Cells(LastRow, "C").value = CategSel
            Sheets("Stockage Epreuves C2").Cells(LastRow, "D").value = Taille.Text
            Sheets("Stockage Epreuves C2").Cells(LastRow, "AV").value = CodeEpreuve.Text
            Sheets("Stockage Epreuves C2").Cells(LastRow, "F").value = TypePart.Text
            Sheets("Stockage Epreuves C2").Select
            Range("A1").Select
            Sheets("Gestion CrewTimer").Select
          MsgBox "L'�preuve � �t� cr��e avec succ�s !", vbOKOnly + vbInformation, "Epreuve Cr��e"
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
    Me.TypePart.AddItem ("Homme")
    Me.TypePart.AddItem ("Femme")
    Me.TypePart.AddItem ("Mixte")
End Sub



