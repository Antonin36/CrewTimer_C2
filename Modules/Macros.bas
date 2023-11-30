Attribute VB_Name = "Macros"
Sub Affiche_Reglages()
    ReglagesRegate.Show
End Sub
Sub AfficheGestInscriptions_CT()
    GestIns_CT.Show
End Sub
Sub AfficheGestInscriptions_C2()
    GestIns_C2.Show
End Sub
Sub AfficheGestImp_PostTirages_CT()
        Sheets("Impressions Tirages CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Gestion CrewTimer").Select
End Sub
Sub AfficheGestImp_PostTirages_C2()
        Sheets("Impressions Tirages C2").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Gestion Concept2").Select
End Sub
Sub AfficheGestImp_PostResultats_CT()
        Sheets("Impressions R�sultats CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Gestion CrewTimer").Select
End Sub
Sub AfficheGestImp_PostResultats_C2()
        Sheets("Impressions R�sultats C2").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Impressions Concept2").Select
End Sub
Sub AfficheGestImp_CT()
    GestImp_CT.Show
End Sub
Sub AfficheGestImp_C2()
    GestImp_C2.Show
End Sub
Sub AfficheAccueilPostEmargement()

End Sub
Sub Affiche_Tirages_CT()
    GestionTirages_CT.Show
End Sub
Sub Affiche_Tirages_C2()
    GestionTirages_C2.Show
End Sub
Sub Affiche_Imp_Tirages_CT()
    ImpTirages_CT.Show
End Sub
Sub Affiche_Imp_Tirages_C2()
    ImpTirages_C2.Show
End Sub
Sub Affiche_Imp_Resultat_CT()
    ImpResultats_CT.Show
End Sub
Sub Affiche_Imp_Resultat_C2()
    ImpResultats_C2.Show
End Sub
Sub Retour_Accueil()
    Sheets("Accueil").Select
End Sub
Sub Affiche_Gest_CT()
    If Sheets("R�glages R�gate").Range("E16").value = "Indoor" Then
    MsgBox "Vous avez param�tr� une r�gate Indoor, l'acc�s � la gestion CrewTimer est impossible. Merci de v�rifier vos param�tres de r�gate.", vbOKOnly + vbExclamation, "Acc�s Impossible"
    Else
    Sheets("Gestion CrewTimer").Select
    End If
End Sub
Sub Affiche_Gest_C2()
'MsgBox "En cours de cr�ation...", vbCritical, "Acc�s Interdit"
If Sheets("R�glages R�gate").Range("E16").value = "Mer" Or Sheets("R�glages R�gate").Range("E16").value = "Rivi�re" Then
    MsgBox "Vous avez param�tr� une r�gate Rivi�re ou Mer, l'acc�s � la gestion Concept2 est impossible. Merci de v�rifier vos param�tres de r�gate.", vbOKOnly + vbExclamation, "Acc�s Impossible"
    Else
    Sheets("Gestion Concept2").Select
    End If
End Sub
Sub Affiche_Impr_CT()
    Sheets("Impressions CT").Select
End Sub
Sub Affiche_Impr_C2()
    Sheets("Impressions C2").Select
End Sub
Sub Affiche_Impr_ReinitImpressions_CT()
    ActiveWorkbook.ActiveSheet.Select
    Range("A13:H420").Select
    Selection.ClearContents
    Range("A13").Select
    Sheets("Gestion CrewTimer").Select
End Sub
Sub Affiche_Impr_ReinitImpressions_C2()
    ActiveWorkbook.ActiveSheet.Select
    Range("A13:H420").Select
    Selection.ClearContents
    Range("A13").Select
    Sheets("Gestion Concept2").Select
End Sub
Sub Affiche_Export_CT()
    Sheets("Feuille CrewTimer").Select
End Sub
Sub Affiche_Export_C2()
    Sheets("Feuille Concept2").Select
End Sub
Sub Affiche_Gest_Course_CT()
    AfficherCourses_CT.Show
End Sub
Sub Affiche_Gest_Course_C2()
    AfficherCourses_C2.Show
End Sub
Sub GenererRAC2()
Dim reponse As VbMsgBoxResult
    reponse = MsgBox("�tes-vous s�r de vouloir g�n�rer les fichiers RAC2 ?", vbQuestion + vbYesNo, "Confirmation G�n�ration")
        If reponse = vbYes Then
        Dim dossier As String
    Dim cheminFichier As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "S�lectionnez le r�pertoire d'enregistrement"
         If .Show = 0 Then Exit Sub 'user cancels
         dossier = .SelectedItems(1)
    End With
    ' Convertir le dictionnaire JSON en une cha�ne JSON
    Dim jsonString As String
    Dim Json As String
    Dim LastRow As Long
    Dim i As Long
    LastRow = Sheets("Programme des Courses C2").Cells(Sheets("Programme des Courses C2").Rows.Count, "BC").End(xlUp).Row
    For i = 2 To LastRow
    Do While Sheets("Programme des Courses C2").Cells(i, 55).value = ""
    i = i + 1
    Loop
    Json = Sheets("Programme des Courses C2").Cells(i, 55).value
    jsonString = Sheets("Programme des Courses C2").Cells(i, 55).value ' Le param�tre Whitespace ajoute une indentation pour une meilleure lisibilit�
    cheminFichier = dossier & "\" & Sheets("R�glages R�gate").Range("D4").value & "_" & Sheets("Programme des Courses C2").Cells(i, 3).value & "_" & Sheets("Programme des Courses C2").Cells(i, 4).value & ".rac2"  ' Remplacez "votre_fichier.json" par le nom de fichier de votre choix
    ' Enregistrer la cha�ne JSON dans le fichier
            Open cheminFichier For Output As #1
            Print #1, jsonString
            Close #1
    Next i
    Sheets("Gestion Concept2").Select
    MsgBox "Les fichiers ont �t� g�n�r�s avec succ�s !", vbInformation, "Fichiers G�n�r�s"
    Else
        Exit Sub
    End If
End Sub
Sub Import_GOAL_C2()
Dim user_selected_filename As String
   
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Fichiers Export GOAL", "*.csv"
        .Title = "S�lectionner l'Export GOAL"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename = .SelectedItems(1)
    End With

    Sheets("Import GOAL C2").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.CutCopyMode = False
    
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & user_selected_filename, Destination:=Range("$A$1"))
        .Name = "ImportGOAL"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 6
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Call clearConnectionsAndQueries
    Dim foundCellFinGOAL As Range
    Set foundCellFinGOAL = Cells.Find(What:="EQUIPAGES INCOMPLETS", After:=ActiveCell, LookIn:= _
    xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)

    If Not foundCellFinGOAL Is Nothing Then
        ' S�lectionne les lignes o� la correspondance est trouv�e
        Rows(foundCellFinGOAL.Row & ":" & foundCellFinGOAL.Row + 5000).Select
        ' Efface le contenu des cellules s�lectionn�es
        Selection.ClearContents
    End If
    Range("A1:EZ999").Select
    Selection.Replace What:="é", Replacement:="�", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("A1").Select
    Sheets("Gestion Concept2").Select
    MsgBox "L'import du fichier GOAL � �t� r�ussi avec succ�s !", vbInformation, "Import GOAL"
End Sub
Sub Import_GOAL_CT()
Attribute Import_GOAL_CT.VB_ProcData.VB_Invoke_Func = " \n14"
Dim user_selected_filename As String
   
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Fichiers Export GOAL", "*.csv"
        .Title = "S�lectionner l'Export GOAL"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename = .SelectedItems(1)
    End With

    Sheets("Import GOAL CT").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.CutCopyMode = False
    
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & user_selected_filename, Destination:=Range("$A$1"))
        .Name = "ImportGOAL"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 6
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Call clearConnectionsAndQueries
    Dim foundCellFinGOAL As Range
    Set foundCellFinGOAL = Cells.Find(What:="EQUIPAGES INCOMPLETS", After:=ActiveCell, LookIn:= _
    xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)

    If Not foundCellFinGOAL Is Nothing Then
        ' S�lectionne les lignes o� la correspondance est trouv�e
        Rows(foundCellFinGOAL.Row & ":" & foundCellFinGOAL.Row + 5000).Select
        ' Efface le contenu des cellules s�lectionn�es
        Selection.ClearContents
    End If
    Range("A1:EZ999").Select
    Selection.Replace What:="é", Replacement:="�", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("A1").Select
    Sheets("Gestion CrewTimer").Select
    MsgBox "L'import du fichier GOAL � �t� r�ussi avec succ�s !", vbInformation, "Import GOAL"
End Sub
Sub FermerSauvegarder()
Dim answer2 As Integer
answer2 = MsgBox("Voulez-vous fermer le syst�me ?", vbYesNo + vbQuestion, "Fermeture Syst�me")
  If answer2 = vbYes Then
    ActiveWorkbook.Save
    Application.Quit
    Else
    Exit Sub
  End If
End Sub
Sub EffTiragesetCT()
Dim answer1 As Integer
answer1 = MsgBox("Confirmez-vous l'effacement de la feuille CrewTimer ainsi que des Tirages ?", vbYesNo + vbExclamation, "Effacement CrewTimer et Tirages")
  If answer1 = vbYes Then
    Sheets("Feuille CrewTimer").Select
        Range("A8:R999").Select
        Selection.EntireRow.Delete
        Sheets("Pr�paration Tirages CT").Select
        Range("A2:R999").Select
        Selection.EntireRow.Delete
        Sheets("Feuille CrewTimer").Select
    MsgBox "La feuille CrewTimer ainsi que les tirages ont �t� effac�s !", vbOKOnly + vbInformation, "CrewTimer et Tirages Effac�s"
  Else
    Exit Sub
  End If
End Sub
Sub EffTiragesetC2()
Dim answer1 As Integer
answer1 = MsgBox("Confirmez-vous l'effacement de la feuille Concept2 ainsi que des Tirages ?", vbYesNo + vbExclamation, "Effacement Concept2 et Tirages")
  If answer1 = vbYes Then
    Sheets("Feuille Concept2").Select
        Range("A8:R999").Select
        Selection.EntireRow.Delete
        Sheets("Pr�paration Tirages C2").Select
        Range("A2:R999").Select
        Selection.EntireRow.Delete
        ActiveWorkbook.Worksheets("Programme des Courses C2").Columns("BC").ClearContents
        Sheets("Feuille Concept2").Select
    MsgBox "La feuille Concept2 ainsi que les tirages ont �t� effac�s !", vbOKOnly + vbInformation, "Concept2 et Tirages Effac�s"
  Else
    Exit Sub
  End If
End Sub
Sub clearConnectionsAndQueries()
Dim cn
Dim qt As QueryTable
Dim lo As ListObject
Dim ws As Worksheet
For Each cn In ThisWorkbook.Connections
    cn.Delete
Next
For Each ws In ThisWorkbook.Worksheets
    For Each qt In ws.QueryTables
        qt.Delete
    Next qt
    On Error Resume Next 'Ignore error if there's no query in table.
    For Each lo In ws.ListObjects
        lo.QueryTable.Delete
    Next lo
    On Error GoTo 0
Next ws
End Sub
Sub ImportResultat_CT()
    Dim user_selected_filename2 As String
   
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Fichier Export R�sultats CrewTimer", "*.csv"
        .Title = "S�lectionner l'Export R�sultat CrewTimer"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename2 = .SelectedItems(1)
    End With

    Sheets("Import Resultats CT").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.CutCopyMode = False
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & user_selected_filename2, Destination:=Range("$A$1"))
        .Name = "r12685"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Call clearConnectionsAndQueries
    Dim dCol As Long, Col As Long
    Dim tColSup, Flg As Boolean
    ' # Liste des colonnes � conserver
    ' Respecter l'orthographe de chaque terme
    tColSup = Split("EventNum,Event,Place,Crew,Bow,Stroke,AdjTime,Delta", ",")
    ' Avec la feuille
    With Sheets("Import Resultats CT")
    ' Derni�re colonne
    dCol = .Cells(1, Columns.Count).End(xlToLeft).Column
    ' Pour chaque colonne
    For Col = dCol To 1 Step -1
      ' V�rifier si nom de colonne trouv� dans celles � supprimer
      Flg = IsError(Application.Match(.Cells(1, Col).value, tColSup, 0))
      ' Si c'est le cas on supprime
      If Flg Then .Cells(1, Col).EntireColumn.Delete Shift:=xlToLeft
    Next Col
    End With
    Sheets("Gestion CrewTimer").Select
    MsgBox "L'import du fichier r�sultat � �t� r�ussi avec succ�s !", vbInformation, "Import R�sultats"
End Sub

Sub ImportResultat_C2()
    Dim user_selected_filename2 As String
    Dim LastRow As Long
    Dim DetailedResultsRow As Long
    
     Sheets("Import Resultats C2").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.CutCopyMode = False
   Do
   ' Trouver la derni�re ligne utilis�e dans la feuille "Import Resultats C2"
        LastRow = Sheets("Import Resultats C2").Cells(Sheets("Import Resultats C2").Rows.Count, "A").End(xlUp).Row

   ' Si la derni�re ligne utilis�e est la premi�re ligne (peut-�tre la feuille est vide), alors commencez � la ligne 1, sinon, � la derni�re ligne utilis�e + 1
        If LastRow = 1 Then
            LastRow = 1
        Else
            LastRow = LastRow + 1
        End If
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Fichier Export R�sultats Concept2", "*.txt"
        .Title = "S�lectionner l'Export R�sultat Concept2"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename2 = .SelectedItems(1)
    End With

    
    With Sheets("Import Resultats C2").QueryTables.Add(Connection:="TEXT;" & user_selected_filename2, Destination:=Sheets("Import Resultats C2").Cells(LastRow, 1))
        .Name = "r12685"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Call clearConnectionsAndQueries
    DetailedResultsRow = WorksheetFunction.Match("Detailed Results", Sheets("Import Resultats C2").Range("A:A"), 0)
    On Error Resume Next
        StartRow = LastRow
        EndRow = DetailedResultsRow + 3
        Sheets("Import Resultats C2").Rows(StartRow & ":" & EndRow).Delete Shift:=xlUp
    MsgBox "L'import du fichier r�sultat � �t� r�ussi avec succ�s !", vbInformation, "Import R�sultats"
    If MsgBox("Voulez-vous importer un autre fichier ?", vbQuestion + vbYesNo, "Importer un autre fichier ?") = vbNo Then
            Sheets("Gestion Concept2").Select
            Exit Do
        End If
    Loop
End Sub
Function CalculerAge(dateActuelle As Date, dateNaissance As Date) As Integer
    ' Calculer l'�ge
    Dim age As Integer
    age = DateDiff("yyyy", dateNaissance, dateActuelle)

    ' V�rifier si l'anniversaire est d�j� pass� cette ann�e
    If DateSerial(Year(dateActuelle), Month(dateNaissance), Day(dateNaissance)) > dateActuelle Then
        age = age - 1 ' D�cr�menter l'�ge si l'anniversaire n'est pas encore pass�
    End If

    ' Retourner l'�ge calcul�
    CalculerAge = age
End Function




