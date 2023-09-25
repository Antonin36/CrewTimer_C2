Attribute VB_Name = "Macros"
Sub Affiche_Reglages()
Attribute Affiche_Reglages.VB_ProcData.VB_Invoke_Func = " \n14"
' Affiche_Reglages Macro

    ReglagesRegate.Show
    
End Sub
Sub AfficheGestImp_PostTirages_CT()
        Sheets("Impressions Tirages CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Impressions CT").Select
End Sub
Sub AfficheGestImp_PostResultats_CT()
        Sheets("Impressions Résultats CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Impressions CT").Select
End Sub
Sub Affiche_Tirages()
' Affiche_Tirages Macro

    GestionTirages_CT.Show
    
End Sub
Sub Affiche_Imp_Tirages_CT()
' Affiche_Tirages Macro

    ImpTirages_CT.Show
    
End Sub
Sub Affiche_Imp_Resultat_CT()
' Affiche_Tirages Macro

    ImpResultats_CT.Show
    
End Sub
Sub Retour_Accueil()
Attribute Retour_Accueil.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Retour_Accueil Macro

    Sheets("Accueil").Select
End Sub

Sub Affiche_Gest_CT()
'
' Retour_Accueil Macro
    If Sheets("Réglages Régate").Range("E16").Value = "Indoor" Then
    MsgBox "Vous avez paramétré une régate Indoor, l'accès à la gestion CrewTimer est impossible. Merci de vérifier vos paramètres de régate.", vbOKOnly + vbExclamation, "Accès Impossible"
    Else
    Sheets("Gestion CrewTimer").Select
    End If
End Sub
Sub Affiche_Gest_C2()
MsgBox "En cours de création...", vbCritical, "Accès Interdit"
'If Sheets("Réglages Régate").Range("E16").Value = "Mer" Or Sheets("Réglages Régate").Range("E16").Value = "Rivière" Then
    'MsgBox "Vous avez paramétré une régate Rivière ou Mer, l'accès à la gestion Concept2 est impossible. Merci de vérifier vos paramètres de régate.", vbOKOnly + vbExclamation, "Accès Impossible"
   ' Else
    'Sheets("Gestion CrewTimer").Select
    'End If
End Sub
Sub Affiche_Impr_CT()
'
' Retour_Accueil Macro

    Sheets("Impressions CT").Select
End Sub
Sub Affiche_Impr_ReinitImpressions_CT()
'
' Retour_Accueil Macro

    Sheets("Impressions CT").Select
End Sub
Sub Affiche_Export_CT()
'
' Affiche_Export_CT Macro

    Sheets("Feuille CrewTimer").Select
End Sub

Sub Affiche_Gest_Course_CT()
'
' Affiche_Gest_Course_CT Macro
    AfficherCourses_CT.Show
End Sub

Sub Import_GOAL()
Attribute Import_GOAL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Import_GOAL Macro
'

'
Dim user_selected_filename As String
   
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Text Files", "*.csv"
        .Title = "Sélectionner l'export GOAL"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename = .SelectedItems(1)
    End With

    Sheets("Import GOAL").Select
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
    Sheets("Gestion CrewTimer").Select
    MsgBox "L'import du fichier GOAL à été réussi avec succès !", vbInformation, "Import GOAL"
End Sub
Sub FermerSauvegarder()
Dim answer2 As Integer
answer2 = MsgBox("Voulez-vous fermer le système ?", vbYesNo + vbQuestion, "Fermeture Système")
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
        Range("A8:K999").Select
        Selection.EntireRow.Delete
        Sheets("Préparation Tirages").Select
        Range("A2:K999").Select
        Selection.EntireRow.Delete
        Sheets("Feuille CrewTimer").Select
    MsgBox "La feuille CrewTimer ainsi que les tirages ont été effacés !", vbOKOnly + vbInformation, "CrewTimer et Tirages Effacés"
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
Sub ImportResultat()
'
' ImportResultat Macro
'

'
    Sheets("Import Resultats").Select
    Dim user_selected_filename2 As String
   
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Text Files", "*.csv"
        .Title = "Sélectionner l'export Résultat CrewTimer"
        If .Show = 0 Then Exit Sub 'user cancels
        user_selected_filename2 = .SelectedItems(1)
    End With

    Sheets("Import Resultats").Select
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
    Sheets("Gestion CrewTimer").Select
    MsgBox "L'import du fichier résultat à été réussi avec succès !", vbInformation, "Import Résultats"
End Sub



