VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestionTirages_CT 
   Caption         =   "Gestion des Tirages"
   ClientHeight    =   7776
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18024
   OleObjectBlob   =   "GestionTirages_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestionTirages_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreationTirages_Click()
    Dim LastRow As Long
    Dim partants As Long
    Dim numlignegoal As Long
    Dim Equipage As String
    Dim rg As Range
    Dim i As Long, j As Long, k As Long, l As Long
    Dim cat As Long
    Dim trigramme As String
    partants = 0
    numlignegoal = 2
    
    'Trouver la derni�re Ligne Utilis�e en Colonne A de la Feuille Origine
    LastRow = Sheets("Programme des Courses").Cells(Sheets("Programme des Courses").Rows.Count, "A").End(xlUp).Row

    'Trouver la derni�re ligne non utilis�e en colonne A de la Feuille Destinataire
    j = Sheets("Pr�paration Tirages").Cells(Sheets("Pr�paration Tirages").Rows.Count, "A").End(xlUp).Row + 1
    
    'Coller Chaque Ligne contenant Oui en H
    For i = 1 To LastRow
            If Sheets("Programme des Courses").Cells(i, 8).Value = "Oui" Then
                partants = 0
                Equipage = ""
                trigramme = ""
                Do While partants < Sheets("R�glages R�gate").Range("E14").Value
                    Sheets("Programme des Courses").Rows(i).Copy Destination:=Worksheets("Pr�paration Tirages").Range("A" & j)
                    Sheets("Pr�paration Tirages").Cells(j, 1).Value = Sheets("Pr�paration Tirages").Cells(j, 7)
                    Dim A As String
                    A = Sheets("Pr�paration Tirages").Cells(j, 3).Value & "-" & Sheets("Pr�paration Tirages").Cells(j, 4).Value
                    Dim B As String
                    B = Sheets("Pr�paration Tirages").Cells(j, 6).Value & "-" & Sheets("Pr�paration Tirages").Cells(j, 4).Value
                    Sheets("Pr�paration Tirages").Cells(j, 3).Value = A
                    Sheets("Pr�paration Tirages").Cells(j, 4).Value = B
                    Sheets("Pr�paration Tirages").Cells(j, 5).Value = A
                    Sheets("Pr�paration Tirages").Cells(j, 6).Value = Sheets("Pr�paration Tirages").Cells(j, 9).Value
                    Dim u As Integer
                    For u = 10 To 50
                    Set rg = Sheets("Import GOAL").Cells(numlignegoal, 3).Find(What:=Sheets("Pr�paration Tirages").Cells(j, u).Value, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not rg Is Nothing Then Exit For
                    If u = 50 Then partants = partants + 1
                    Next u

                    If Not rg Is Nothing Then
                        Equipage = Sheets("Import GOAL").Cells(numlignegoal, 5).Value & " (" & Sheets("Import GOAL").Cells(numlignegoal, 6).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 7).Value
                        If Sheets("Import GOAL").Cells(numlignegoal, 18).Value <> "" Then
                            Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 18).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 19).Value
                            If Sheets("Import GOAL").Cells(numlignegoal, 30).Value <> "" Then
                                Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 30).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 31).Value
                                If Sheets("Import GOAL").Cells(numlignegoal, 42).Value <> "" Then
                                    Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 42).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 43).Value
                                    If Sheets("Import GOAL").Cells(numlignegoal, 54).Value <> "" Then
                                        Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 54).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 55).Value
                                        If Sheets("Import GOAL").Cells(numlignegoal, 66).Value <> "" Then
                                            Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 66).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 67).Value
                                            If Sheets("Import GOAL").Cells(numlignegoal, 78).Value <> "" Then
                                                Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 78).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 79).Value
                                                If Sheets("Import GOAL").Cells(numlignegoal, 90).Value <> "" Then
                                                    Equipage = Equipage & " / " & Sheets("Import GOAL").Cells(numlignegoal, 90).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 91).Value
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf Sheets("Import GOAL").Cells(numlignegoal, 104).Value <> "" Then
                            Equipage = Equipage & " / Bar : " & Sheets("Import GOAL").Cells(numlignegoal, 104).Value & " " & Sheets("Import GOAL").Cells(numlignegoal, 105).Value
                        Else
                            Equipage = Equipage & ")"
                        End If
                        Sheets("Pr�paration Tirages").Cells(j, 7).Value = Equipage
                        Sheets("Pr�paration Tirages").Cells(j, 8).Value = Sheets("Import GOAL").Cells(numlignegoal, 5).Value
                        Sheets("Pr�paration Tirages").Cells(j, 9).Value = Sheets("Import GOAL").Cells(numlignegoal, 3).Value
                        Sheets("Pr�paration Tirages").Cells(j, 11).Value = Sheets("Import GOAL").Cells(numlignegoal, 5).Value
                        If Sheets("R�glages R�gate").Range("E16").Value = "Rivi�re" Then
                            Sheets("Pr�paration Tirages").Cells(j, 10).Value = partants + 1
                            numlignegoal = numlignegoal + 1
                            j = j + 1
                            partants = partants + 1
                        Else
                            Sheets("Pr�paration Tirages").Cells(j, 10) = Sheets("Import GOAL").Cells(numlignegoal, 4).Value
                            numlignegoal = numlignegoal + 1
                            j = j + 1
                            partants = partants + 1
                        End If
                    End If
                Loop
                
            End If
    Next i
    
    Sheets("Pr�paration Tirages").Select
    Columns("H:H").Select
    Selection.Replace What:="SAINTE CROIX AVN 04", Replacement:="AVN4", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MANOSQUE AC", Replacement:="ACDM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ESPARRON DE VERDON CN", Replacement:="CNEV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAVINES LE LAC ASP", Replacement:="ASP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EMBRUN CA", Replacement:="CAEM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NICE CN", Replacement:="CNNI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CANNES MANDELIEU RCCM", Replacement:="RCCM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MENTON SCA", Replacement:="SCAM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VILLEFRANCHE SN", Replacement:="SNVI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE ASPTT", Replacement:="AMSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE AAS", Replacement:="AAS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ISTRES ANO", Replacement:="ANOI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CASSIS AC", Replacement:="ACDC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE CA", Replacement:="CAM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT CHAMAS CASC", Replacement:="CASC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PEYROLLES CNPA", Replacement:="CNPA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARIGNANE CMS", Replacement:="CMSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARTIGUES AC", Replacement:="MAAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE RC", Replacement:="RCMA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA CIOTAT SN", Replacement:="SNLC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VITROLLES SA", Replacement:="VSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="HYERES ASPTT", Replacement:="ASAH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SIX OURS ACSF", Replacement:="ACSF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VERDON AC", Replacement:="ACV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT CASSIEN ASC", Replacement:="ASTC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA SEYNE SUR MER AV", Replacement:="ASEY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULON AV", Replacement:="ATON", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINTE MAXIME CA", Replacement:="CAMX", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SALETTES CN", Replacement:="CNDS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AVIGNON SN", Replacement:="SNAV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CADEROUSSE SN", Replacement:="SNCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONACO SN", Replacement:="SNMO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("A1").Select
    Dim LastRow2 As Long
    Dim partants2 As Long
   'Find last used row in a Column A of Sheet1
      LastRow2 = Sheets("Programme des Courses").Cells(Sheets("Programme des Courses").Rows.Count, "A").End(xlUp).Row

   'Find first row where values should be posted in Sheet2
      k = Sheets("Pr�paration Tirages").Cells(Sheets("Pr�paration Tirages").Rows.Count, "A").End(xlUp).Row + 1
 
   
   'Paste each row that contains "Mavs" in column A of Sheet1 into Sheet2
   For l = 1 To LastRow2
           If Sheets("Programme des Courses").Cells(l, 8).Value = "Non" Then
           partants2 = 1
           Do While partants2 < Sheets("R�glages R�gate").Range("E14").Value
               Sheets("Programme des Courses").Rows(l).Copy Destination:=Worksheets("Pr�paration Tirages").Range("A" & k)
                Sheets("Pr�paration Tirages").Cells(k, 1).Value = Sheets("Pr�paration Tirages").Cells(k, 7)
                Dim C As String
                C = Sheets("Pr�paration Tirages").Cells(k, 3).Value & "-" & Sheets("Pr�paration Tirages").Cells(k, 4).Value
                Dim D As String
                D = Sheets("Pr�paration Tirages").Cells(k, 6).Value & "-" & Sheets("Pr�paration Tirages").Cells(k, 4).Value
                Sheets("Pr�paration Tirages").Cells(k, 3).Value = C
                Sheets("Pr�paration Tirages").Cells(k, 4).Value = D
                Sheets("Pr�paration Tirages").Cells(k, 5).Value = C
                Sheets("Pr�paration Tirages").Cells(k, 6).Value = Sheets("Pr�paration Tirages").Cells(k, 9).Value
                Sheets("Pr�paration Tirages").Cells(k, 7).Value = "TBD"
                Sheets("Pr�paration Tirages").Cells(k, 8).Value = "TBD"
                Sheets("Pr�paration Tirages").Cells(k, 9).Value = "TBD"
                Sheets("Pr�paration Tirages").Cells(k, 10).Value = partants2
                Sheets("Pr�paration Tirages").Cells(k, 11).Value = ""
                k = k + 1
                partants2 = partants2 + 1
            Loop
            End If
   Next l
   Dim LastRow3 As Long
   'Find last used row in a Column A of Sheet1
    LastRow3 = Sheets("Pr�paration Tirages").Cells(Sheets("Pr�paration Tirages").Rows.Count, "A").End(xlUp).Row
      
   For w = 1 To LastRow3
           If Sheets("Pr�paration Tirages").Cells(l, 8).Value = "Non" Or Sheets("Pr�paration Tirages").Cells(w, 8).Value = "Oui" Then
           Sheets("Pr�paration Tirages").Rows(w).EntireRow.Delete
           End If
   Next w
       MsgBox "Les tirages ont �t� cr��s avec succ�s !", vbOKOnly + vbInformation, "Tirages Cr��s"
       Sheets("Gestion CrewTimer").Select
End Sub
Private Sub UserForm_Initialize()
Dim random_method As String
random_method = ""
If MsgBox("Voulez-vous utiliser un tirage al�atoire ?", vbYesNo + vbQuestion, "Tirages Al�atoires ?") = vbYes Then
'Mettre cr�er une colonne random, en ER
    random_method = "Al�atoire"
    Sheets("Import GOAL").Select
    Range("ER1").Value = "Random"
    Range("ER2").Select
    Dim rand As Long
    rand = 998
    For rand = 1 To rand
    ActiveCell.Value = Rnd()
    ActiveCell.Offset(1, 0).Select
Next rand
'Trier la table
ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Add2 Key:=Range( _
        "ER2:ER999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL").Sort
        .SetRange Range("A1:ER999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Effacer la colonne random
    Range("ER1:ER999").Value = ""
    Range("A1").Select
    
    ElseIf MsgBox("Voulez-vous utiliser un tirage o� le num�ro du bateau est l'ordre de d�part ? (T�te de Rivi�re UNIQUEMENT)", vbYesNo + vbQuestion, "Tirages par Num�ro de Bateau ?") = vbYes Then
    'Proc�der au tirage via l'ordre croissant des num�ros de bateau
    random_method = "Par l'ordre croissant des num�ros de bateau"
        Sheets("Import GOAL").Select
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Add2 Key:=Range( _
        "D2:D999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With
    Else
    'Tirage Ordre Alphab�tique Nom Court
    MsgBox "Le Tirage va �tre effectu� dans l'ordre alphab�tique des noms courts des clubs.", vbOKOnly + vbInformation, "Tirage Normal"
    random_method = "Par l'ordre alphab�tique des noms courts des clubs"
    Sheets("Import GOAL").Select
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL").Sort.SortFields.Add2 Key:=Range( _
        "E2:E999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Prg Courses selon Ordre Alphab�tique Cat�g
    Sheets("Programme des Courses").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    If MsgBox("Le mode de tirage d�fini est : " + random_method + ". Confirmez-vous ce choix ?", vbYesNo + vbInformation, "Confirmation Mode de Tirage") = vbYes Then
    ' Feuille � S�lectionner
    Sheets("Pr�paration Tirages").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
    Sheets("Gestion CrewTimer").Select
    Exit Sub
    Else
    Call UserForm_Initialize
    End If
    ' Feuille � S�lectionner
    Sheets("Pr�paration Tirages").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
    Sheets("Gestion CrewTimer").Select
End Sub

Private Sub ValidTirages_Click()
Dim answer1 As Integer
answer1 = MsgBox("Confirmez-vous la validation des tirages ?", vbYesNo + vbExclamation, "Confirmation Validation Tirages")
  If answer1 = vbYes Then
  Sheets("Pr�paration Tirages").Select
    Range("A2:K999").Select
    Selection.Copy
    Sheets("Feuille CrewTimer").Select
    Range("A8").Select
    ActiveSheet.Paste
    Sheets("Gestion CrewTimer").Select
    MsgBox "Les tirages ont bien �t� valid�s et transf�r�s dans la table pour l'export vers CrewTimer !", vbOKOnly + vbInformation, "Tirages Valid�s"
    Unload Me
  Else
    Exit Sub
  End If
End Sub
Private Sub SupprTirages_Click()
Dim answer2 As Integer
answer2 = MsgBox("Confirmez-vous l'invallidation des tirages ?", vbYesNo + vbExclamation, "Confirmation Invalidation Tirages")
  If answer2 = vbYes Then
    Sheets("Pr�paration Tirages").Select
    Range("A2:K999").Select
    Selection.EntireRow.Delete
    Range("A1").Select
    MsgBox "Les tirages ont bien �t� invalid�s.", vbOKOnly + vbInformation, "Tirages Invalid�s"
    Call UserForm_Initialize
  Else
    Exit Sub
  End If
End Sub
Private Sub Quit_Click()
Sheets("Programme des Courses").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Gestion CrewTimer").Select
 Unload Me
End Sub
