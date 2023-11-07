Private alea As Boolean
Private numCollection As New Collection ' Déclarez une collection pour stocker les numéros de ligne

' Fonction pour ajouter un numéro de ligne à la collection
Sub AddToCollection(col As Collection, item As Long)
    On Error Resume Next
    col.Add item, CStr(item) ' Utilisez CStr pour convertir le numéro de ligne en une clé de chaîne unique
    On Error GoTo 0
End Sub
Function IsInCollection(col As Collection, val As Long) As Boolean
    On Error Resume Next
    Dim item As Variant
    IsInCollection = False
    For Each item In col
        If item = val Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    On Error GoTo 0
End Function

' Fonction pour vider la collection
Sub ClearCollection(col As Collection)
    Set col = New Collection
End Sub
Private Sub CreationTirages_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
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
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ClearCollection numCollection
    'Trouver la dernière Ligne Utilisée en Colonne A de la Feuille Origine
    LastRow = Sheets("Programme des Courses CT").Cells(Sheets("Programme des Courses CT").Rows.Count, "A").End(xlUp).Row

    'Trouver la dernière ligne non utilisée en colonne A de la Feuille Destinataire
    j = Sheets("Préparation Tirages CT").Cells(Sheets("Préparation Tirages CT").Rows.Count, "A").End(xlUp).Row + 1
    Dim limgoal As Long
    limgoal = Sheets("Import GOAL CT").Cells(Sheets("Import GOAL CT").Rows.Count, "C").End(xlUp).Row + 1
    
    'Coller Chaque Ligne contenant Oui en H
    For i = 1 To LastRow
            If Sheets("Programme des Courses CT").Cells(i, 8).Value = "Oui" Then
                partants = 0
                Equipage = ""
                trigramme = ""
                numlignegoal = 2
                'numlignegoal = numlignegoal + 1
                Do While partants < Sheets("Réglages Régate").Range("E14").Value
                    Sheets("Programme des Courses CT").Rows(i).Copy Destination:=Worksheets("Préparation Tirages CT").Range("A" & j)
                    Sheets("Préparation Tirages CT").Cells(j, 1).Value = Sheets("Préparation Tirages CT").Cells(j, 7)
                    Dim A As String
                    A = Sheets("Préparation Tirages CT").Cells(j, 3).Value & "_" & Sheets("Préparation Tirages CT").Cells(j, 4).Value
                    Dim B As String
                    B = Sheets("Préparation Tirages CT").Cells(j, 6).Value & "_" & Sheets("Préparation Tirages CT").Cells(j, 4).Value
                    Sheets("Préparation Tirages CT").Cells(j, 3).Value = A
                    Sheets("Préparation Tirages CT").Cells(j, 4).Value = B
                    Sheets("Préparation Tirages CT").Cells(j, 5).Value = A
                    Sheets("Préparation Tirages CT").Cells(j, 6).Value = Sheets("Préparation Tirages CT").Cells(j, 9).Value
                    Dim u As Integer
                    For u = 10 To 50
                    Do
                    If Not IsInCollection(numCollection, numlignegoal) Then
                        Exit Do ' Sort de la boucle Do si numlignegoal n'est pas dans la collection
                    Else
                        numlignegoal = numlignegoal + 1 ' Incrémente numlignegoal
                    End If
                    Loop
                    Dim casegoal As String
                    Dim casetirage As String
                    casegoal = Sheets("Import GOAL CT").Cells(numlignegoal, 3).Value
                    casetirage = Sheets("Préparation Tirages CT").Cells(j, u).Value
                    If casegoal = casetirage Then Exit For
                    
                    If u = 50 Then numlignegoal = numlignegoal + 1
                    If numlignegoal = limgoal Then partants = partants + 1
                    Next u
                   
                    If casegoal = casetirage Then
                        Equipage = Sheets("Import GOAL CT").Cells(numlignegoal, 5).Value & " (" & Sheets("Import GOAL CT").Cells(numlignegoal, 6).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 7).Value
                        If Sheets("Import GOAL CT").Cells(numlignegoal, 18).Value <> "" Then
                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 18).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 19).Value
                            If Sheets("Import GOAL CT").Cells(numlignegoal, 30).Value <> "" Then
                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 30).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 31).Value
                                If Sheets("Import GOAL CT").Cells(numlignegoal, 42).Value <> "" Then
                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 42).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 43).Value
                                    If Sheets("Import GOAL CT").Cells(numlignegoal, 54).Value <> "" Then
                                        Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 54).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 55).Value
                                        If Sheets("Import GOAL CT").Cells(numlignegoal, 66).Value <> "" Then
                                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 66).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 67).Value
                                            If Sheets("Import GOAL CT").Cells(numlignegoal, 78).Value <> "" Then
                                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 78).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 79).Value
                                                If Sheets("Import GOAL CT").Cells(numlignegoal, 90).Value <> "" Then
                                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 90).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 91).Value
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Sheets("Import GOAL CT").Cells(numlignegoal, 104).Value <> "" Then
                            Equipage = Equipage & " / Bar : " & Sheets("Import GOAL CT").Cells(numlignegoal, 104).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 105).Value
                        End If
                        Equipage = Equipage & ")"
                        Sheets("Préparation Tirages CT").Cells(j, 7).Value = Equipage
                        Sheets("Préparation Tirages CT").Cells(j, 8).Value = Sheets("Import GOAL CT").Cells(numlignegoal, 5).Value
                        Sheets("Préparation Tirages CT").Cells(j, 9).Value = Sheets("Import GOAL CT").Cells(numlignegoal, 3).Value
                        Sheets("Préparation Tirages CT").Cells(j, 11).Value = Sheets("Import GOAL CT").Cells(numlignegoal, 5).Value
                        If Sheets("Réglages Régate").Range("E16").Value = "Rivière" Then
                            If Sheets("Réglages Régate").Range("G16").Value = "TDR" Then
                            Sheets("Préparation Tirages CT").Cells(j, 10) = Sheets("Import GOAL CT").Cells(numlignegoal, 4).Value
                            Else
                            Sheets("Préparation Tirages CT").Cells(j, 10).Value = partants + 1
                            End If
                            numCollection.Add numlignegoal
                            numlignegoal = 2
                            j = j + 1
                            partants = partants + 1
                            casegoal = ""
                            casetirage = ""
                        Else
                            Sheets("Préparation Tirages CT").Cells(j, 10) = Sheets("Import GOAL CT").Cells(numlignegoal, 4).Value
                            numCollection.Add numlignegoal
                            numlignegoal = 2
                            j = j + 1
                            partants = partants + 1
                            casegoal = ""
                            casetirage = ""
                        End If
                    End If
                Loop
                
            End If
    Next i
    
    Sheets("Préparation Tirages CT").Select
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
      LastRow2 = Sheets("Programme des Courses CT").Cells(Sheets("Programme des Courses CT").Rows.Count, "A").End(xlUp).Row

   'Find first row where values should be posted in Sheet2
      k = Sheets("Préparation Tirages CT").Cells(Sheets("Préparation Tirages CT").Rows.Count, "A").End(xlUp).Row + 1
 
   
   'Paste each row that contains "Mavs" in column A of Sheet1 into Sheet2
   For l = 1 To LastRow2
           If Sheets("Programme des Courses CT").Cells(l, 8).Value = "Non" Then
           partants2 = 0
           Do While partants2 < Sheets("Réglages Régate").Range("E14").Value
               Sheets("Programme des Courses CT").Rows(l).Copy Destination:=Worksheets("Préparation Tirages CT").Range("A" & k)
                Sheets("Préparation Tirages CT").Cells(k, 1).Value = Sheets("Préparation Tirages CT").Cells(k, 7)
                Dim C As String
                C = Sheets("Préparation Tirages CT").Cells(k, 3).Value & "_" & Sheets("Préparation Tirages CT").Cells(k, 4).Value
                Dim D As String
                D = Sheets("Préparation Tirages CT").Cells(k, 6).Value & "_" & Sheets("Préparation Tirages CT").Cells(k, 4).Value
                Sheets("Préparation Tirages CT").Cells(k, 3).Value = C
                Sheets("Préparation Tirages CT").Cells(k, 4).Value = D
                Sheets("Préparation Tirages CT").Cells(k, 5).Value = C
                Sheets("Préparation Tirages CT").Cells(k, 6).Value = Sheets("Préparation Tirages CT").Cells(k, 9).Value
                Sheets("Préparation Tirages CT").Cells(k, 7).Value = "TBD"
                Sheets("Préparation Tirages CT").Cells(k, 8).Value = "TBD"
                Sheets("Préparation Tirages CT").Cells(k, 9).Value = "TBD"
                Sheets("Préparation Tirages CT").Cells(k, 10).Value = partants2 + 1
                Sheets("Préparation Tirages CT").Cells(k, 11).Value = ""
                k = k + 1
                partants2 = partants2 + 1
            Loop
            End If
   Next l
   Dim LastRow3 As Long
   'Find last used row in a Column A of Sheet1
    LastRow3 = Sheets("Préparation Tirages CT").Cells(Sheets("Préparation Tirages CT").Rows.Count, "A").End(xlUp).Row
      
   For w = 1 To LastRow3
           If Sheets("Préparation Tirages CT").Cells(l, 8).Value = "Non" Or Sheets("Préparation Tirages CT").Cells(w, 8).Value = "Oui" Then
           Sheets("Préparation Tirages CT").Rows(w).EntireRow.Delete
           End If
   Next w
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
       MsgBox "Les tirages ont été créés avec succès !", vbOKOnly + vbInformation, "Tirages Créés"
       
       Sheets("Préparation Tirages CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Gestion CrewTimer").Select
            Sheets("Programme des Courses CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Préparation Tirages CT").Select
        Rows("1:1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$Z$403").AutoFilter Field:=7, Criteria1:="( )"
        Rows("2:1048576").Select
        Selection.Delete Shift:=xlUp
        Rows("1:1").Select
        ActiveSheet.ShowAllData
        Selection.AutoFilter
        Range("A1").Select
        If Sheets("Réglages Régate").Range("G16").Value = "Rand" Then
        Sheets("Réglages Régate").Select
        Sheets("Réglages Régate").Range("G16").Value = ""
        Sheets("Feuille CrewTimer").Select
        ' Générer des valeurs aléatoires dans la colonne M
        ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("M8:M1000").FormulaR1C1 = "=RAND()"

        ' Tri des données
        With ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("A8:A1000"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
                :="Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", DataOption:=xlSortNormal
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("B8:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("M8:M1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            .SetRange ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("A7:N1000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        ' Suppression de la colonne M
        ActiveWorkbook.Worksheets("Feuille CrewTimer").Columns("M:M").Delete Shift:=xlToLeft
        End If
        Sheets("Gestion CrewTimer").Select
End Sub
Private Sub UserForm_Initialize()
alea = False
Dim random_method As String
random_method = ""
Sheets("Réglages Régate").Select
            Sheets("Réglages Régate").Range("G16").Value = ""
    Sheets("Gestion CrewTimer").Select
If MsgBox("Voulez-vous utiliser un tirage aléatoire ?", vbYesNo + vbQuestion, "Tirages Aléatoires ?") = vbYes Then
alea = True
'Mettre créer une colonne random, en ER
    Sheets("Réglages Régate").Select
    Sheets("Réglages Régate").Range("G16").Value = "Rand"
    Sheets("Import GOAL CT").Select
    random_method = "Aléatoire"
    Sheets("Import GOAL CT").Select
    Range("ER1").Value = "Random"
    Range("ER2").Select
    Dim rand As Long
    rand = 998
    For rand = 1 To rand
    ActiveCell.Value = Rnd()
    ActiveCell.Offset(1, 0).Select
Next rand
'Trier la table
ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "ER2:ER999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL CT").Sort
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
    
    ElseIf MsgBox("Voulez-vous utiliser un tirage où le numéro du bateau est l'ordre de départ ? (Tête de Rivière UNIQUEMENT)", vbYesNo + vbQuestion, "Tirages par Numéro de Bateau ?") = vbYes Then
    'Procéder au tirage via l'ordre croissant des numéros de bateau
    random_method = "Par l'ordre croissant des numéros de bateau"
    Sheets("Réglages Régate").Select
    Sheets("Réglages Régate").Range("G16").Value = "TDR"
        Sheets("Import GOAL CT").Select
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "D2:D999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL CT").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With
    
    Else
    'Tirage Ordre Alphabétique Nom Court
    MsgBox "Le Tirage va être effectué dans l'ordre alphabétique des noms courts des clubs.", vbOKOnly + vbInformation, "Tirage Normal"
    random_method = "Par l'ordre alphabétique des noms courts des clubs"
    Sheets("Import GOAL CT").Select
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "E2:E999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL CT").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Prg Courses selon Ordre Alphabétique Catég
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    If MsgBox("Le mode de tirage défini est : " + random_method + ". Confirmez-vous ce choix ?", vbYesNo + vbInformation, "Confirmation Mode de Tirage") = vbYes Then
    ' Feuille à Sélectionner
    Sheets("Préparation Tirages CT").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
    Sheets("Gestion CrewTimer").Select
    Exit Sub
    Else
    Call UserForm_Initialize
    End If
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Feuille à Sélectionner
    Sheets("Préparation Tirages CT").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
        Sheets("Gestion CrewTimer").Select
End Sub

Private Sub ValidTirages_Click()
Dim answer1 As Integer
answer1 = MsgBox("Confirmez-vous la validation des tirages ?", vbYesNo + vbExclamation, "Confirmation Validation Tirages")
  If answer1 = vbYes Then
  Sheets("Préparation Tirages CT").Select
    Range("A2:K999").Select
    Selection.Copy
    Sheets("Feuille CrewTimer").Select
    Range("A8").Select
    ActiveSheet.Paste
    Sheets("Gestion CrewTimer").Select
    Sheets("Programme des Courses CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Réglages Régate").Select
            Sheets("Réglages Régate").Range("G16").Value = ""
            Sheets("Gestion CrewTimer").Select
    MsgBox "Les tirages ont bien été validés et transférés dans la table pour l'export vers CrewTimer !", vbOKOnly + vbInformation, "Tirages Validés"
    Unload Me
  Else
    Exit Sub
  End If
End Sub
Private Sub SupprTirages_Click()
Dim answer2 As Integer
answer2 = MsgBox("Confirmez-vous l'invalidation des tirages ?", vbYesNo + vbExclamation, "Confirmation Invalidation Tirages")
  If answer2 = vbYes Then
    Sheets("Préparation Tirages CT").Select
    Range("A2:K999").Select
    Selection.EntireRow.Delete
    Range("A1").Select
    MsgBox "Les tirages ont bien été invalidés.", vbOKOnly + vbInformation, "Tirages Invalidés"
    Call UserForm_Initialize
  Else
    Exit Sub
  End If
End Sub
Private Sub Quit_Click()
Sheets("Programme des Courses CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Réglages Régate").Select
            Sheets("Réglages Régate").Range("G16").Value = ""
            Sheets("Gestion CrewTimer").Select
 Unload Me
End Sub
