VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestEpreuves_CT 
   Caption         =   "Gestion des Epreuves"
   ClientHeight    =   8140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17980
   OleObjectBlob   =   "GestEpreuves_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestEpreuves_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreerEpreuve_Click()
    AjoutEpreuve_CT.Show
End Sub

Private Sub Import_GOAL_Click()
Dim user_selected_filename As String
Dim cheminAccesComplet As String
Dim nomFichier As String
Dim extensionFichier As String
Dim feuilleDestination As Worksheet
Dim derniereLigne As Long

    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Fichiers Export Epreuves GOAL", "*.xls"
        .Title = "S�lectionner l'Export des Epreuves de GOAL"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename = .SelectedItems(1)
    End With
    Application.DisplayAlerts = False
    ' R�cup�rez le chemin d'acc�s complet
    cheminAccesComplet = user_selected_filename

    ' R�cup�rez le nom du fichier (y compris l'extension)
    nomFichier = Right(cheminAccesComplet, Len(cheminAccesComplet) - InStrRev(cheminAccesComplet, "\"))

    ' R�cup�rez l'extension du fichier
    extensionFichier = Right(nomFichier, Len(nomFichier) - InStrRev(nomFichier, "."))


    Sheets("Stockage Import Cat�gories CT").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.CutCopyMode = False
    
    ' Sp�cifiez la feuille de destination dans le classeur actif
    Set feuilleDestination = ThisWorkbook.Sheets("Stockage Import Cat�gories CT")

    ' Ouvrez l'autre classeur Excel
    Workbooks.Open cheminAccesComplet
    Workbooks(nomFichier).Unprotect

    ' Copiez les donn�es depuis l'autre classeur (� ajuster en fonction de votre structure)
    Workbooks(nomFichier).Sheets("Export").UsedRange.Copy

    ' Collez les donn�es dans la feuille de destination � partir de la ligne apr�s la derni�re ligne utilis�e
    feuilleDestination.Cells(1, 1).PasteSpecial xlPasteValues

    ' Fermez l'autre classeur sans enregistrer les modifications
    Workbooks(nomFichier).Close SaveChanges:=False
    Application.DisplayAlerts = True
    Call clearConnectionsAndQueries
    Sheets("Stockage Import Cat�gories CT").Select
    Dim feuille As Worksheet
    Dim texte As String
    Dim mot As Variant
    Dim cellule As Range
    
    ' Remplacez "Nom de votre feuille" par le nom de votre feuille
    Set feuille = ThisWorkbook.Sheets("Stockage Import Cat�gories CT")
    
    ' Tableau des cha�nes � rechercher
    Dim chaines As Variant
    chaines = Array("H1", "H2", "H4", "H8", "F1", "F2", "F4", "F8", "M1", "M2", "M4", "M8", "HR4", "FR4", "MR4")
    
    ' Boucler � travers les cellules non vides de la colonne 1 (colonne A) de la feuille
    For Each cellule In feuille.Range("A2:A" & feuille.Cells(feuille.Rows.Count, 1).End(xlUp).Row)
        If Not IsEmpty(cellule.Value) Then
            ' R�cup�rer le texte de la cellule
            texte = cellule.Value
            
            ' Boucler � travers les cha�nes
            For Each mot In chaines
                ' Trouver la position de la cha�ne dans le texte
                Dim position As Long
                position = InStr(1, texte, mot, vbTextCompare)
                
                ' Si la cha�ne est trouv�e, placer le r�sultat dans la colonne C de la m�me ligne
                If position > 0 Then
                    cellule.Offset(0, 2).NumberFormat = "@"
                    cellule.Offset(0, 2).Value = Left(texte, position - 1)
                    Exit For
                End If
            Next mot
            If InStr(1, cellule.Value, "1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la m�me ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).Value = 1
            End If
            If InStr(1, cellule.Value, "2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la m�me ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).Value = 2
            End If
            If InStr(1, cellule.Value, "4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la m�me ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).Value = 4
            End If
            If InStr(1, cellule.Value, "8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la m�me ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).Value = 8
            End If
            If InStr(1, cellule.Value, "+", vbTextCompare) > 0 Then
                ' Marquer dans la colonne E de la m�me ligne
                cellule.Offset(0, 4).NumberFormat = "@"
                cellule.Offset(0, 4).Value = "Oui"
            Else
                cellule.Offset(0, 4).NumberFormat = "@"
                cellule.Offset(0, 4).Value = "Non"
            End If
            If InStr(1, cellule.Value, "H1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Homme"
            End If
            If InStr(1, cellule.Value, "H2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Homme"
            End If
            If InStr(1, cellule.Value, "H4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Homme"
            End If
            If InStr(1, cellule.Value, "H8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Homme"
            End If
            If InStr(1, cellule.Value, "F1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Femme"
            End If
            If InStr(1, cellule.Value, "F2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Femme"
            End If
            If InStr(1, cellule.Value, "F4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Femme"
            End If
            If InStr(1, cellule.Value, "F8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Femme"
            End If
            If InStr(1, cellule.Value, "M1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Mixte"
            End If
            If InStr(1, cellule.Value, "M2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Mixte"
            End If
            If InStr(1, cellule.Value, "M4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Mixte"
            End If
            If InStr(1, cellule.Value, "M8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Mixte"
            End If
            If InStr(1, cellule.Value, "FR4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Femme"
            End If
            If InStr(1, cellule.Value, "HR4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Homme"
            End If
            If InStr(1, cellule.Value, "MR4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne F de la m�me ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).Value = "Mixte"
            End If
        End If
    Next cellule
    Range("A2:E999").Select
    Selection.Copy
    Sheets("Stockage Epreuves CT").Select
    Range("A2").Select
    ActiveSheet.Paste
    Range("C2:C999").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "@"
    Range("A1").Select
    
                    
End Sub

Private Sub ModifierEpreuve_Click()
    SelModifEpreuve_CT.Show
End Sub

Private Sub SupprEpreuve_Click()
    SupprEpreuve_CT.Show
End Sub

Private Sub UserForm_Initialize()
' Feuille � S�lectionner
    Me.Import_GOAL.Caption = "Importer les Epreuves" & vbCrLf & "depuis GOAL"
    Sheets("Stockage Epreuves CT").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "A1:E200"
    TableauPrgCourses.ColumnWidths = "80;500;500;80;60"
    Sheets("Gestion CrewTimer").Select
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub

