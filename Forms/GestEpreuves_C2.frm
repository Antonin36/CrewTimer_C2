VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestEpreuves_C2 
   Caption         =   "Gestion des Epreuves"
   ClientHeight    =   8140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17980
   OleObjectBlob   =   "GestEpreuves_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestEpreuves_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreerEpreuve_Click()
    AjoutEpreuve_C2.Show
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
        .Title = "Sélectionner l'Export des Epreuves de GOAL"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename = .SelectedItems(1)
    End With
    Application.DisplayAlerts = False
    ' Récupérez le chemin d'accès complet
    cheminAccesComplet = user_selected_filename

    ' Récupérez le nom du fichier (y compris l'extension)
    nomFichier = Right(cheminAccesComplet, Len(cheminAccesComplet) - InStrRev(cheminAccesComplet, "\"))

    ' Récupérez l'extension du fichier
    extensionFichier = Right(nomFichier, Len(nomFichier) - InStrRev(nomFichier, "."))


    Sheets("Stockage Import Catégories C2").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.CutCopyMode = False
    
    ' Spécifiez la feuille de destination dans le classeur actif
    Set feuilleDestination = ThisWorkbook.Sheets("Stockage Import Catégories C2")

    ' Ouvrez l'autre classeur Excel
    Workbooks.Open cheminAccesComplet
    Workbooks(nomFichier).Unprotect

    ' Copiez les données depuis l'autre classeur (à ajuster en fonction de votre structure)
    Workbooks(nomFichier).Sheets("Export").UsedRange.Copy

    ' Collez les données dans la feuille de destination à partir de la ligne après la dernière ligne utilisée
    feuilleDestination.Cells(1, 1).PasteSpecial xlPasteValues

    ' Fermez l'autre classeur sans enregistrer les modifications
    Workbooks(nomFichier).Close SaveChanges:=False
    Application.DisplayAlerts = True
    Call clearConnectionsAndQueries
    Sheets("Stockage Import Catégories C2").Select
    Dim feuille As Worksheet
    Dim texte As String
    Dim mot As Variant
    Dim cellule As Range
    
    ' Remplacez "Nom de votre feuille" par le nom de votre feuille
    Set feuille = ThisWorkbook.Sheets("Stockage Import Catégories C2")
    
    ' Tableau des chaînes à rechercher
    Dim chaines As Variant
    chaines = Array("H1", "H2", "H4", "H8", "F1", "F2", "F4", "F8", "M1", "M2", "M4", "M8", "HR4", "FR4")
    
    ' Boucler à travers les cellules non vides de la colonne 1 (colonne A) de la feuille
    For Each cellule In feuille.Range("A2:A" & feuille.Cells(feuille.Rows.Count, 1).End(xlUp).Row)
        If Not IsEmpty(cellule.value) Then
            ' Récupérer le texte de la cellule
            texte = cellule.value
            
            ' Boucler à travers les chaînes
            For Each mot In chaines
                ' Trouver la position de la chaîne dans le texte
                Dim position As Long
                position = InStr(1, texte, mot, vbTextCompare)
                
                ' Si la chaîne est trouvée, placer le résultat dans la colonne C de la même ligne
                If position > 0 Then
                    cellule.Offset(0, 2).NumberFormat = "@"
                    cellule.Offset(0, 2).value = Left(texte, position - 1)
                    Exit For
                End If
            Next mot
            If InStr(1, cellule.value, "+", vbTextCompare) > 0 Then
                ' Marquer dans la colonne E de la même ligne
                cellule.Offset(0, 4).NumberFormat = "@"
                cellule.Offset(0, 4).value = "Oui"
            Else
                cellule.Offset(0, 4).NumberFormat = "@"
                cellule.Offset(0, 4).value = "Non"
            End If
            If InStr(1, cellule.value, "H1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 1
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Homme"
            End If
            If InStr(1, cellule.value, "H2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 2
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Homme"
            End If
            If InStr(1, cellule.value, "H4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 4
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Homme"
            End If
            If InStr(1, cellule.value, "H8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 8
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Homme"
            End If
            If InStr(1, cellule.value, "F1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 1
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Femme"
            End If
            If InStr(1, cellule.value, "F2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 2
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Femme"
            End If
            If InStr(1, cellule.value, "F4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 4
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Femme"
            End If
            If InStr(1, cellule.value, "F8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 8
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Femme"
            End If
            If InStr(1, cellule.value, "M1", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 1
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Mixte"
            End If
            If InStr(1, cellule.value, "M2", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 2
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Mixte"
            End If
            If InStr(1, cellule.value, "M4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 4
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Mixte"
            End If
            If InStr(1, cellule.value, "M8", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 8
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Mixte"
            End If
            If InStr(1, cellule.value, "FR4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 4
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Femme"
            End If
            If InStr(1, cellule.value, "HR4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 4
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Homme"
            End If
            If InStr(1, cellule.value, "MR4", vbTextCompare) > 0 Then
                ' Marquer dans la colonne D de la même ligne
                cellule.Offset(0, 3).NumberFormat = "@"
                cellule.Offset(0, 3).value = 4
                ' Marquer dans la colonne F de la même ligne
                cellule.Offset(0, 5).NumberFormat = "@"
                cellule.Offset(0, 5).value = "Mixte"
            End If
        End If
    Next cellule
    Range("A2:G999").Select
    Selection.Copy
    Sheets("Stockage Epreuves C2").Select
    Range("A2").Select
    ActiveSheet.Paste
    Range("C2:C999").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "@"
    Range("A1").Select
    
                    
End Sub

Private Sub ModifierEpreuve_Click()
    SelModifEpreuve_C2.Show
End Sub

Private Sub SupprEpreuve_Click()
    SupprEpreuve_C2.Show
End Sub

Private Sub UserForm_Initialize()
' Feuille à Sélectionner
    Me.Import_GOAL.Caption = "Importer les Epreuves" & vbCrLf & "depuis GOAL"
    Sheets("Stockage Epreuves C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauPrgCourses.RowSource = "A1:E200"
    TableauPrgCourses.ColumnWidths = "80;500;500;80;60"
    Sheets("Gestion Concept2").Select
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub


