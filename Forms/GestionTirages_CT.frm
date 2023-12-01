VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestionTirages_CT 
   Caption         =   "Gestion des Tirages"
   ClientHeight    =   7740
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   17980
   OleObjectBlob   =   "GestionTirages_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestionTirages_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private alea As Boolean
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
    Dim vidangestock As Long
    For vidangestock = 2 To 2000
    Sheets("Stockage Divers").Cells(vidangestock, 1).value = ""
    If vidangestock = 2000 Then Exit For
    Next vidangestock
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
    Range("B8:B999").Select
    Selection.NumberFormat = "hh:mm:ss"
    Range("B8").Select
    ClearCollection numCollection
    'Trouver la dernière Ligne Utilisée en Colonne A de la Feuille Origine
    LastRow = Sheets("Programme des Courses CT").Cells(Sheets("Programme des Courses CT").Rows.Count, "A").End(xlUp).Row

    'Trouver la dernière ligne non utilisée en colonne A de la Feuille Destinataire
    j = Sheets("Préparation Tirages CT").Cells(Sheets("Préparation Tirages CT").Rows.Count, "A").End(xlUp).Row + 1
    Dim limgoal As Long
    limgoal = Sheets("Import GOAL CT").Cells(Sheets("Import GOAL CT").Rows.Count, "C").End(xlUp).Row + 1
    
    'Coller Chaque Ligne contenant Oui en H
    For i = 1 To LastRow
            If Sheets("Programme des Courses CT").Cells(i, 8).value = "Oui" Then
                partants = 0
                Equipage = ""
                trigramme = ""
                numlignegoal = 2
                'numlignegoal = numlignegoal + 1
                Do While partants < Sheets("Réglages Régate").Range("E14").value
                    Dim lignestockvide As Long
                    lignestockvide = Sheets("Stockage Divers").Cells(Sheets("Stockage Divers").Rows.Count, 2).End(xlUp).Row + 1
                    
                    Sheets("Programme des Courses CT").Rows(i).Copy Destination:=Worksheets("Préparation Tirages CT").Range("A" & j)
                    Sheets("Préparation Tirages CT").Cells(j, 1).value = Sheets("Préparation Tirages CT").Cells(j, 7)
                    Dim A As String
                    A = Sheets("Préparation Tirages CT").Cells(j, 3).value & "_" & Sheets("Préparation Tirages CT").Cells(j, 4).value
                    Dim B As String
                    B = Sheets("Préparation Tirages CT").Cells(j, 6).value & "_" & Sheets("Préparation Tirages CT").Cells(j, 4).value
                    Sheets("Préparation Tirages CT").Cells(j, 3).value = A
                    Sheets("Préparation Tirages CT").Cells(j, 4).value = B
                    Sheets("Préparation Tirages CT").Cells(j, 5).value = A
                    Sheets("Préparation Tirages CT").Cells(j, 6).value = Sheets("Préparation Tirages CT").Cells(j, 9).value
                    Dim wsStockage As Worksheet
                    Set wsStockage = ThisWorkbook.Sheets("Stockage Divers")
                    Do While numlignegoal <= limgoal
                    ' Recherche de la valeur dans la colonne B
                    Dim matchResult As Variant
                    matchResult = Application.Match(numlignegoal, wsStockage.Columns("B"), 0)

                    ' Si une correspondance est trouvée, passez à la valeur suivante
                    If Not IsError(matchResult) Then
                    numlignegoal = numlignegoal + 1
                    Else
                    Exit Do
                    End If
                    Loop
                    Dim u As Integer
                    For u = 10 To 50
                    
                    
                    Dim casegoal As String
                    Dim casetirage As String
                    casegoal = Sheets("Import GOAL CT").Cells(numlignegoal, 3).value
                    casetirage = Sheets("Préparation Tirages CT").Cells(j, u).value
                    If casegoal = casetirage Then Exit For
                    
                    If u = 50 Then numlignegoal = numlignegoal + 1
                    If numlignegoal = limgoal Then partants = partants + 1
                    Next u
                   
                    If casegoal = casetirage Then
                        Equipage = Sheets("Import GOAL CT").Cells(numlignegoal, 5).value & " (" & Sheets("Import GOAL CT").Cells(numlignegoal, 6).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 7).value
                        If Sheets("Import GOAL CT").Cells(numlignegoal, 18).value <> "" Then
                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 18).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 19).value
                            If Sheets("Import GOAL CT").Cells(numlignegoal, 30).value <> "" Then
                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 30).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 31).value
                                If Sheets("Import GOAL CT").Cells(numlignegoal, 42).value <> "" Then
                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 42).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 43).value
                                    If Sheets("Import GOAL CT").Cells(numlignegoal, 54).value <> "" Then
                                        Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 54).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 55).value
                                        If Sheets("Import GOAL CT").Cells(numlignegoal, 66).value <> "" Then
                                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 66).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 67).value
                                            If Sheets("Import GOAL CT").Cells(numlignegoal, 78).value <> "" Then
                                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 78).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 79).value
                                                If Sheets("Import GOAL CT").Cells(numlignegoal, 90).value <> "" Then
                                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 90).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 91).value
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Sheets("Import GOAL CT").Cells(numlignegoal, 102).value <> "" Then
                            Equipage = Equipage & " / Bar : " & Sheets("Import GOAL CT").Cells(numlignegoal, 102).value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 103).value
                        End If
                        Equipage = Equipage & ")"
                        Sheets("Préparation Tirages CT").Cells(j, 7).value = Equipage
                        Sheets("Préparation Tirages CT").Cells(j, 8).value = Sheets("Import GOAL CT").Cells(numlignegoal, 5).value
                        Sheets("Préparation Tirages CT").Cells(j, 9).value = Sheets("Import GOAL CT").Cells(numlignegoal, 3).value
                        Sheets("Préparation Tirages CT").Cells(j, 11).value = Sheets("Import GOAL CT").Cells(numlignegoal, 5).value
                        If Sheets("Réglages Régate").Range("E16").value = "Rivière" Then
                            If Sheets("Réglages Régate").Range("G16").value = "TDR" Then
                            Sheets("Préparation Tirages CT").Cells(j, 10) = Sheets("Import GOAL CT").Cells(numlignegoal, 4).value
                            Else
                            Sheets("Préparation Tirages CT").Cells(j, 10).value = partants + 1
                            End If
                            Sheets("Stockage Divers").Cells(lignestockvide, 2).value = numlignegoal
                            numlignegoal = 2
                            j = j + 1
                            partants = partants + 1
                            casegoal = ""
                            casetirage = ""
                        Else
                            Sheets("Préparation Tirages CT").Cells(j, 10) = Sheets("Import GOAL CT").Cells(numlignegoal, 4).value
                            Sheets("Stockage Divers").Cells(lignestockvide, 2).value = numlignegoal
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
    Call RemplacerCodeCourt
    Range("A1").Select
    Dim LastRow2 As Long
    Dim partants2 As Long
   'Find last used row in a Column A of Sheet1
      LastRow2 = Sheets("Programme des Courses CT").Cells(Sheets("Programme des Courses CT").Rows.Count, "A").End(xlUp).Row

   'Find first row where values should be posted in Sheet2
      k = Sheets("Préparation Tirages CT").Cells(Sheets("Préparation Tirages CT").Rows.Count, "A").End(xlUp).Row + 1
 
   
   'Paste each row that contains "Mavs" in column A of Sheet1 into Sheet2
   For l = 1 To LastRow2
           If Sheets("Programme des Courses CT").Cells(l, 8).value = "Non" Then
           partants2 = 0
           Do While partants2 < Sheets("Réglages Régate").Range("E14").value
               Sheets("Programme des Courses CT").Rows(l).Copy Destination:=Worksheets("Préparation Tirages CT").Range("A" & k)
                Sheets("Préparation Tirages CT").Cells(k, 1).value = Sheets("Préparation Tirages CT").Cells(k, 7)
                Dim C As String
                C = Sheets("Préparation Tirages CT").Cells(k, 3).value & "_" & Sheets("Préparation Tirages CT").Cells(k, 4).value
                Dim D As String
                D = Sheets("Préparation Tirages CT").Cells(k, 6).value & "_" & Sheets("Préparation Tirages CT").Cells(k, 4).value
                Sheets("Préparation Tirages CT").Cells(k, 3).value = C
                Sheets("Préparation Tirages CT").Cells(k, 4).value = D
                Sheets("Préparation Tirages CT").Cells(k, 5).value = C
                Sheets("Préparation Tirages CT").Cells(k, 6).value = Sheets("Préparation Tirages CT").Cells(k, 9).value
                Sheets("Préparation Tirages CT").Cells(k, 7).value = "TBD"
                Sheets("Préparation Tirages CT").Cells(k, 8).value = "TBD"
                Sheets("Préparation Tirages CT").Cells(k, 9).value = "TBD"
                Sheets("Préparation Tirages CT").Cells(k, 10).value = partants2 + 1
                Sheets("Préparation Tirages CT").Cells(k, 11).value = ""
                k = k + 1
                partants2 = partants2 + 1
            Loop
            End If
   Next l
   Dim LastRow3 As Long
   'Find last used row in a Column A of Sheet1
    LastRow3 = Sheets("Préparation Tirages CT").Cells(Sheets("Préparation Tirages CT").Rows.Count, "A").End(xlUp).Row
      
   For w = 1 To LastRow3
           If Sheets("Préparation Tirages CT").Cells(l, 8).value = "Non" Or Sheets("Préparation Tirages CT").Cells(w, 8).value = "Oui" Then
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
        If Sheets("Réglages Régate").Range("G16").value = "TDR" Then
        Sheets("Préparation Tirages CT").Select
    Range("J2:J999").Select
    ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Add2 Key _
        :=Range("J2:J999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
        End If
        If Sheets("Réglages Régate").Range("G16").value = "Rand" Then
        Sheets("Réglages Régate").Select
        Sheets("Réglages Régate").Range("G16").value = ""
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
            Sheets("Réglages Régate").Range("G16").value = ""
    Sheets("Gestion CrewTimer").Select
If MsgBox("Voulez-vous utiliser un tirage aléatoire ?", vbYesNo + vbQuestion, "Tirages Aléatoires ?") = vbYes Then
alea = True
'Mettre créer une colonne random, en ER
    Sheets("Réglages Régate").Select
    Sheets("Réglages Régate").Range("G16").value = "Rand"
    Sheets("Import GOAL CT").Select
    random_method = "Aléatoire"
    Sheets("Import GOAL CT").Select
    Range("ER1").value = "Random"
    Range("ER2").Select
    Dim rand As Long
    rand = 998
    For rand = 1 To rand
    ActiveCell.value = Rnd()
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
    Range("ER1:ER999").value = ""
    Range("A1").Select
    
    ElseIf MsgBox("Voulez-vous utiliser un tirage où le numéro du bateau est l'ordre de départ ? (Tête de Rivière UNIQUEMENT)", vbYesNo + vbQuestion, "Tirages par Numéro de Bateau ?") = vbYes Then
    'Procéder au tirage via l'ordre croissant des numéros de bateau
    random_method = "Par l'ordre croissant des numéros de bateau"
    Sheets("Réglages Régate").Select
    Sheets("Réglages Régate").Range("G16").value = "TDR"
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
            Sheets("Réglages Régate").Range("G16").value = ""
            Range("B8:B999").Select
            Selection.NumberFormat = "hh:mm:ss"
            Range("B8").Select
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
            Sheets("Réglages Régate").Range("G16").value = ""
            Sheets("Gestion CrewTimer").Select
 Unload Me
End Sub
Private Sub RemplacerCodeCourt()
'LR01 (Auvergne-Rhône-Alpes)
    'CD01 (Ain)
    Selection.Replace What:="STE JULIE ASPLA", Replacement:="ASPA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BELLEGARDE AV", Replacement:="ABEL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BELLEY-VIRIGNIN ABHR", Replacement:="ABHR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NANTUA AC", Replacement:="ACN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PARCIEUX CAPSV", Replacement:="CAPS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TREVOUX CN", Replacement:="CNTR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DIVONNE CN", Replacement:="CNDV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SERRIERES CN", Replacement:="CNSE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHAMBOD RCVA", Replacement:="RCVA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD03 (Allier)
    Selection.Replace What:="VICHY CA", Replacement:="CAVI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONTLUCON CMA", Replacement:="CMOA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD07 (Ardèche)
    Selection.Replace What:="VIVIERS AVMC", Replacement:="AVMC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD15 (Cantal)
    Selection.Replace What:="SAINT-FLOUR ANG", Replacement:="ANDG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD26 (Drôme)
    Selection.Replace What:="ROMANS ANRPG", Replacement:="ARPG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ROMANS ARP", Replacement:="ARP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VALENCE AV", Replacement:="AVAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PIERRELATTE CA", Replacement:="CAPI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TAIN-TOURNON SN", Replacement:="SNTT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD38 (Isère)
    Selection.Replace What:="FONTAINE AS AVIRON", Replacement:="ASF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SASSENAGE AC", Replacement:="ACSI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GRENOBLE AV", Replacement:="AGRE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SUD GRESIVAUDAN CA", Replacement:="CASG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LAFFREY CVA", Replacement:="CVAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LAC BLEU AV", Replacement:="ADLB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD42 (Loire)
    Selection.Replace What:="ST PRIEST LA ROCHE ARGL", Replacement:="ARGL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-PIERRE-DE-BOEUF AV", Replacement:="APET", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ROANNE-LE COTEAU AV", Replacement:="ARLC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT ETIENNE AV", Replacement:="ASTE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD63 (Puy de Dôme)
    Selection.Replace What:="CLERMONT FERRAND ACA", Replacement:="ACAY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD69 (Rhône)
    Selection.Replace What:="VAULX-EN-VELIN ASUL AVIRON", Replacement:="ASUL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LYON CALUIRE ACLC", Replacement:="ACLC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ST ROMAIN EN GAL ACPV", Replacement:="ACPV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DECINES AV", Replacement:="ADEC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MEYZIEU AV", Replacement:="AMAJ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LYON AUNL", Replacement:="AUNL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VILLEFRANCHE AUN", Replacement:="AUNV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RONNO BNPALS", Replacement:="BNLS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LYON CA", Replacement:="CALY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CONDRIEU SN", Replacement:="SNCO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-FONS SNS", Replacement:="SNSS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LOIRE SUR RHONE SN", Replacement:="SNLR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD73 (Savoie)
    Selection.Replace What:="AIGUEBELETTE ACL", Replacement:="ACLA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHAMBERY CNCB", Replacement:="CNCB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AIX LES BAINS EN AVIRON", Replacement:="ENAB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD74 (Haute Savoie)
    Selection.Replace What:="SEVRIER ASLA", Replacement:="ASLA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SCIEZ BN", Replacement:="BNS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ANNECY CN", Replacement:="CNAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TALLOIRES CN", Replacement:="CNTA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="THONON CA", Replacement:="CATH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EVIAN CA", Replacement:="CAEV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ANNECY LE VIEUX CS", Replacement:="CSAV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LEMAN AC", Replacement:="LAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR02 (Bourgogne-Franche-Comté)
    'CD21 (Côte d'Or)
    Selection.Replace What:="SEURRE AC", Replacement:="ACSE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DIJON A", Replacement:="ADIJ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AUXONNE AE", Replacement:="AEA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD25 (Doubs)
    Selection.Replace What:="PONTARLIER AV", Replacement:="APON", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-VIT AV", Replacement:="ASTV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BESANCON SN", Replacement:="SNBI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD39 (Jura)
    Selection.Replace What:="DOLE AC", Replacement:="ACDO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ORGELET CAV", Replacement:="CAVO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BOIS D'AMONT CAVJ", Replacement:="CAVJ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD58 (Nièvre)
    Selection.Replace What:="CLAMECY CN", Replacement:="CNCL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD70 (Haute Saône)
    Selection.Replace What:="GRAY SAONE AV", Replacement:="AGS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VESOUL CNHSV", Replacement:="CNHV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD71 (Saône et Loire)
    Selection.Replace What:="LOUHANS CABL", Replacement:="CABL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHALON SUR SAONE CA", Replacement:="CACS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LE CREUSOT CN", Replacement:="CNCR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MACON SR", Replacement:="SRMA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD89 (Yonne)
    Selection.Replace What:="SAINT-FARGEAU A", Replacement:="AFAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="JOIGNY US AVIRON", Replacement:="USJA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VILLENEUVE-SUR-YONNE AV", Replacement:="VSYA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD90 (Territoire de Belfort)
    Selection.Replace What:="BELFORT SR", Replacement:="SPRB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR03 (Bretagne)
    'CD22 (Côtes d'Armor)
    Selection.Replace What:="TREMARGAT APAPP", Replacement:="AAPP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PLERIN ABA", Replacement:="ABDA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GOUET ACGP", Replacement:="ACGP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RANCE AV", Replacement:="ARAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LEZARDRIEUX AT", Replacement:="ATRI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOURNEMINE AA", Replacement:="AARM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DINAN CN", Replacement:="CNDD", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TREGUIER TAR", Replacement:="TRAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD29 (Finistère)
    Selection.Replace What:="PLOUGASTEL-DAOULAS AV", Replacement:="ARM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MORLAIX AB", Replacement:="ABM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BREST AV", Replacement:="ABRE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHATEAULIN AV", Replacement:="ACHA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FOUESNANT AMC", Replacement:="ADMC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PLOUGONVELIN AV MER", Replacement:="AMP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DOUARNENEZ A", Replacement:="ADOA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CONCARNEAU AMC", Replacement:="AMCO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TREGUNC CA", Replacement:="TRCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BREST-IROISE YC", Replacement:="YCBI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT POL YC", Replacement:="YCSP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD35 (Ille et Vilaine)
    Selection.Replace What:="VITRE ANCPV", Replacement:="ANCP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FEINS ALPA", Replacement:="ALPA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="REDON AP", Replacement:="APDR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FORGES LA FORET FA", Replacement:="FAVI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RENNES ECA", Replacement:="RECA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RENNES SR", Replacement:="SRRE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-MALO SNBSM", Replacement:="SBSM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ST LUNAIRE YC", Replacement:="YCSL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD56 (Morbihan)
    Selection.Replace What:="GUER ASAEC", Replacement:="AAEC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AURAY AC", Replacement:="ACDA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SARZEAU ACRH", Replacement:="ACRH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LORIENT A SCORFF", Replacement:="ASCO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LORIENT CN", Replacement:="CNLO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VANNES CA", Replacement:="CAVA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LORIENT PLL AV DU TER", Replacement:="PLLO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CARNAC YC", Replacement:="YCCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR04 (Centre-Val de Loire)
    'CD18 (Cher)
    Selection.Replace What:="BOURGES AC", Replacement:="ACDB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT AMAND CN", Replacement:="CNSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD28 (Eure et Loir)
    Selection.Replace What:="BONNEVAL AV", Replacement:="BOAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD37 (Indre et Loire)
    Selection.Replace What:="BLERE ABVC", Replacement:="ABVC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOURS ATM", Replacement:="ATOM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD41 (Loir et Cher)
    Selection.Replace What:="SAINT LAURENT ASG", Replacement:="ASGS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BLOIS AV", Replacement:="ABLE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ST-AIGNAN ACVC", Replacement:="ACVC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD45 (Loiret)
    Selection.Replace What:="MONTARGIS AC", Replacement:="ACMG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ORLEANS-OLIVET AC", Replacement:="ACOO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GIEN AV", Replacement:="AGIE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="COMBLEUX COC", Replacement:="COCC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR05 (Corse)
    'CD20 (Collectivité de Corse)
    Selection.Replace What:="HAUTE CORSE A", Replacement:="AHCO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AJACCIO IRC", Replacement:="IRCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AJACCIO K", Replacement:="KALL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PIETROSELLA RSAC", Replacement:="RSAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR06 (Grand Est)
    'CD08 (Ardennes)
    Selection.Replace What:="FLIZE AFPS", Replacement:="AFPS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SEDAN AV", Replacement:="AVSE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHARLEVILLE MEZIERES CN", Replacement:="CNCM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GIVET PM", Replacement:="PMGI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD10 (Aube)
    Selection.Replace What:="NOGENT SUR SEINE CA", Replacement:="CANO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TROYES SN", Replacement:="SNT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD51 (Marne)
    Selection.Replace What:="REIMS CNRR", Replacement:="CNRR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHALONS EN CHAMPAGNE AV", Replacement:="LPC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EPERNAY SN", Replacement:="SNEP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD52 (Haute Marne)
    Selection.Replace What:="LANGRES ACK52", Replacement:="LACK", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD54 (Meurthe et Moselle)
    Selection.Replace What:="PONT SAINT VINCENT CNHM", Replacement:="CNHM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LIVERDUN CN", Replacement:="CNLV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PONT A MOUSSON SN", Replacement:="SNPM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NANCY SNN", Replacement:="SNNA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOUL US", Replacement:="USTA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD55 (Meuse)
    Selection.Replace What:="BELLEVILLE 55 AV", Replacement:="B55A", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VERDUN CN", Replacement:="CNVD", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT MIHIEL CN", Replacement:="CNSM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD57 (Moselle)
    Selection.Replace What:="SARREGUEMINES AC", Replacement:="ACDS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BASSE HAM Y", Replacement:="LYH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BOUZONVILLE NCB", Replacement:="NCBO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="METZ SR", Replacement:="SRME", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MITTERSHEIM US", Replacement:="USMI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD67 (Bas Rhin)
    Selection.Replace What:="ERSTEIN ACP", Replacement:="ACPE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="STRASBOURG AV", Replacement:="AS81", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="STRASBOURG CA", Replacement:="CAST", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="STRASBOURG CNI", Replacement:="CNIS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="STRASBOURG ISS", Replacement:="ISST", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="STRASBOURG RC", Replacement:="RCST", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD68 (Haut Rhin)
    Selection.Replace What:="COLMAR ACR", Replacement:="ACRC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MULHOUSE MA", Replacement:="MULA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MULHOUSE RC", Replacement:="RCMU", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="UNION REGIO A", Replacement:="URA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD88 (Vosges)
    Selection.Replace What:="GERARDMER AS", Replacement:="ASG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EPINAL AC", Replacement:="ACE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR07 (Hauts de France)
    'CD02 (Aisne)
    Selection.Replace What:="CHATEAU THIERRY AV", Replacement:="ACT2", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-QUENTIN ASQ", Replacement:="ASTQ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SOISSONS SN", Replacement:="SNSO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD59 (Nord)
    Selection.Replace What:="DOUAI ASE", Replacement:="ADSE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LILLE AUN", Replacement:="UNLI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="HAUBOURDINOIS CN", Replacement:="CNHA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ARMENTIERES CLL", Replacement:="CLLA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="COUDEKERQUE EN", Replacement:="ENCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GRAVELINES GA", Replacement:="GRAV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RIEULAY NC", Replacement:="NCDR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DUNKERQUE SP", Replacement:="SPDU", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CAMBRAI UN", Replacement:="UNCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VALENCIENNES UC", Replacement:="VUC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD60 (Oise)
    Selection.Replace What:="CREIL ENO", Replacement:="ENO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="COMPIEGNE SN", Replacement:="SPNC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD62 (Pas de Calais)
    Selection.Replace What:="SAINT-OMER AV", Replacement:="AAUD", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BETHUNE AA", Replacement:="ABA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BOULOGNE AV", Replacement:="ABOU", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CALAIS CA", Replacement:="CADC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD80 (Somme)
    Selection.Replace What:="ABBEVILLE SN", Replacement:="SNAB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AMIENS SN", Replacement:="SNAM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR08 (Ile de France)
    'CD75 (Paris)
    Selection.Replace What:="CLUB FFA", Replacement:="FFA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PARIS US METRO", Replacement:="USMT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD77 (Seine et Marne)
    Selection.Replace What:="FONTAINEBLEAU APF", Replacement:="APF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VAIRES TORCY AV.", Replacement:="AVT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FERTE SOUS JOUARRE CA", Replacement:="CAFJ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TRILPORT CA", Replacement:="CATR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MEAUX CN", Replacement:="CNME", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MELUN CN", Replacement:="SNME", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LAGNY SN", Replacement:="SNLA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD78 (Yvelines)
    Selection.Replace What:="MANTES AS", Replacement:="ASM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VILLENNES - POISSY AC", Replacement:="ACVP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MEULAN LES MUREAUX AMMH", Replacement:="AMMH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ANDRESY CA CONFLUENT", Replacement:="CACO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MAISONS MESNIL CERAMM", Replacement:="CRMM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VERSAILLES CN", Replacement:="CNVS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PORT-MARLY RC", Replacement:="RCPM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD91 (Essonne)
    Selection.Replace What:="CORBEIL ASCE 91", Replacement:="ASCE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="COUDRAY MONTCEAUX A", Replacement:="ADCM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SOISY SUR SEINE CN", Replacement:="CNSO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="POLYTECHNIQUE CSE", Replacement:="CSEP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EVRY SCA", Replacement:="SCAE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="HAUTE SEINE SN", Replacement:="SNHS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RIS ORANGIS USRO", Replacement:="USRO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD92 (Hauts de Seine)
    Selection.Replace What:="BOULOGNE 92", Replacement:="B92", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FRANCE CN", Replacement:="CNF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BASSE SEINE SNBS", Replacement:="SNBS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SEVRES VSN", Replacement:="VSN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD93 (Seine Saint Denis)
    Selection.Replace What:="NOISY LE GRAND ASLM", Replacement:="ASLM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ROWING CLUB SRP", Replacement:="ROWC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD94 (Val de Marne)
    Selection.Replace What:="ACI 94", Replacement:="AC94", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="JOINVILLE AMJ", Replacement:="AMJ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ABLON SUR SEINE CN", Replacement:="CNAS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NOGENT SUR MARNE CN", Replacement:="CNNO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHAMPIGNY SUR MARNE RSCC", Replacement:="RSCC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-MAUR SACSM", Replacement:="SACS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ENCOURAGEMENT - SESN", Replacement:="SESN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PERREUX SN", Replacement:="SNP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD95 (Val d'Oise)
    Selection.Replace What:="JOUY LE MOUTIER ALSO", Replacement:="ALSO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BEAUMONT AV", Replacement:="BEAA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CERGY ESSEC", Replacement:="ESSE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ARGENTEUIL COMA", Replacement:="COMA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SN OISE", Replacement:="SNO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ENGHIEN SN", Replacement:="SNE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BUTRY SUR OISE VOA", Replacement:="VOA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR09 (Normandie)
    'CD14 (Calvados)
    Selection.Replace What:="BLAINVILLE ASRT", Replacement:="ASRT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="HEROUVILLE CHARM", Replacement:="CHAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CAEN CALVADOS SNCC", Replacement:="SNCC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD27 (Eure)
    Selection.Replace What:="AONES LE VAUDREUIL", Replacement:="AONE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOSNY ACAT", Replacement:="ACAT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VERNON EN", Replacement:="ENVE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD50 (Manche)
    Selection.Replace What:="GRANVILLE AV", Replacement:="AGRA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHERBOURG AV", Replacement:="CCAM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CARTERET CAMBC", Replacement:="CAMB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD76 (Seine Maritime)
    Selection.Replace What:="CAUDEBEC EN CAUX ACVS", Replacement:="ACVS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-AUBIN AVCA", Replacement:="AVCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CANTELEU CROISSET CN", Replacement:="CNCC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BELBEUF CN", Replacement:="CNBE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DIEPPE CN", Replacement:="CNDP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ROUEN CNAR", Replacement:="CNAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LE HAVRE SHA", Replacement:="SHA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR10 (Nouvelle Aquitaine)
    'CD16 (Charente)
    Selection.Replace What:="ANGOULEME AC", Replacement:="ACAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="JARNAC AV", Replacement:="AJAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="COGNAC YRC", Replacement:="CYRC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD17 (Charente Maritime)
    Selection.Replace What:="MARANS AV", Replacement:="AVIM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="OLERONAIS CA", Replacement:="CAOL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT SAVINIEN CN", Replacement:="CNSV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA ROCHELLE CAM", Replacement:="CAMR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINTES CA", Replacement:="CASA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD19 (Corrèze)
    Selection.Replace What:="SECHEMAILLES ADSL", Replacement:="ADSL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BRIVE CSN", Replacement:="CSNB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD24 (Dordogne)
    Selection.Replace What:="ROUFFIAC AC", Replacement:="ROAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BERGERAC SN", Replacement:="SNBE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TRELISSAC SN", Replacement:="SPNT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD33 (Gironde)
    Selection.Replace What:="ARCACHON AV", Replacement:="AARC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BEGLES ACBB", Replacement:="ACBB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA REOLE AS", Replacement:="AESR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CALME", Replacement:="CALA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CLAOUEY CN", Replacement:="CNDC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LIBOURNE CN", Replacement:="CNLB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINTE FOY CN", Replacement:="CNFO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BORDEAUX EN", Replacement:="ENB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ENPC ST CIERS", Replacement:="ENPC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CASTILLON RC", Replacement:="RCCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LANGON SN", Replacement:="SNLG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD40 (Landes)
    Selection.Replace What:="SOUSTONS AC", Replacement:="ACSO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AVIRON LANDES", Replacement:="ALAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PEYREHORADE CCG", Replacement:="CCDG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MIMIZAN CN", Replacement:="CNMI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD47 (Lot et Garonne)
    Selection.Replace What:="AGEN AV", Replacement:="AAGE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CLAIRAC AV", Replacement:="AVCL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARMANDE AV", Replacement:="AMAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINTE LIVRADE ASL", Replacement:="ASTL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VILLENEUVE AV", Replacement:="AVIL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT SYLVESTRE CNPSS", Replacement:="CNSS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LIBOS FUMEL CNMLF", Replacement:="CNLF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD64 (Pyrénées Atlantiques)
    Selection.Replace What:="BAYONNE AV", Replacement:="ABAY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="HENDAYE EAE", Replacement:="ENAE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ANGLET IBAIALDE", Replacement:="IBAI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BAYONNE SN", Replacement:="SNBA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT-JEAN-DE-LUZ UR IKARA", Replacement:="URIK", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT JEAN DE LUZ UYAT UYAE", Replacement:="UYAT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD79 (Deux Sèvres)
    Selection.Replace What:="NIORT AC", Replacement:="NIAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD86 (Vienne)
    Selection.Replace What:="CHATELLERAULT SN", Replacement:="SNCH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD87 (Haute Vienne)
    Selection.Replace What:="PALAIS SUR VIENNE AC", Replacement:="ACPA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LIMOGES CN", Replacement:="CNLM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR11 (Occitanie)
    'CD11 (Aude)
    Selection.Replace What:="CARCASSONNE AV", Replacement:="ACAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CASTELNAUDARY AL", Replacement:="ALAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PORT LA NOUVELLE A", Replacement:="ANOU", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NARBONNE AC", Replacement:="NAAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD12 (Aveyron)
    Selection.Replace What:="BOUILLAC AV", Replacement:="BOAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ARVIEU PARELOUP CAAP", Replacement:="CAAP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD30 (Gard)
    Selection.Replace What:="BEAUCAIRE AV", Replacement:="ABEA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT GILLES AC", Replacement:="ACSG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GRAU DU ROI ATC", Replacement:="ATCG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD31 (Haute Garonne)
    Selection.Replace What:="VILLEMUR AS", Replacement:="ASVI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FENOUILLET AB", Replacement:="ADBO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LAUNAC AV", Replacement:="ALAU", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULOUSE A", Replacement:="ATOU", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="GRENADE CN", Replacement:="CNGC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CARBONNE CO", Replacement:="COCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULOUSE EN", Replacement:="ENTO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONTPITOL ACLL", Replacement:="MACL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RAMONVILLE PSA", Replacement:="PSAR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULOUSE ASL", Replacement:="TASL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULOUSE PPR", Replacement:="TPPR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULOUSE UC", Replacement:="TUC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RIEUX-VOLVESTRE USRVA", Replacement:="USRV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RENNEVILLE VRAC DU LAURAGAIS", Replacement:="VRAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD32 (Gers)
    Selection.Replace What:="CAZAUBON AAC", Replacement:="AAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD34 (Hérault)
    Selection.Replace What:="AGDE AV", Replacement:="AAGA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BALARUC A", Replacement:="ABAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BEZIERS AC", Replacement:="ACBI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SETE ACBT", Replacement:="ACBT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA GRANDE MOTTE AC PONANT", Replacement:="ACPO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="OCTON ACS", Replacement:="ACS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MÈZE A", Replacement:="AMEZ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SETE AV", Replacement:="ASET", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CARNON CAMC", Replacement:="CAMC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONTPELLIER AUC", Replacement:="MAUC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONTPELLIER PI", Replacement:="PINT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD46 (Lot)
    Selection.Replace What:="CAHORS AV", Replacement:="ACAD", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CAJARC AC", Replacement:="ACCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DOUELLE AV", Replacement:="ADOE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SENAILLAC LATRONQUIERE CNHS", Replacement:="CNHS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD66 (Pyrénées Orientales)
    Selection.Replace What:="COLLIOURE ASA", Replacement:="ACSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SALSES LE CHATEAU ANC", Replacement:="ANCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BANYULS AV", Replacement:="ABAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LE BARCARES ASS", Replacement:="BAIV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ARGELES G", Replacement:="GRAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PERPIGNAN AV 66", Replacement:="PA66", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PORT-VENDRES SNCVA", Replacement:="SNCV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD81 (Tarn)
    Selection.Replace What:="ALBI AC", Replacement:="ACAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CASTRES SN", Replacement:="CASN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ALMAYRAC SN", Replacement:="SNAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD82 (Tarn et Garonne)
    Selection.Replace What:="GRISOLLES AV", Replacement:="ACG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MOISSAC AC", Replacement:="ACMO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BRESSOLS AC", Replacement:="BRAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONTAUBAN UN", Replacement:="UNMO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR12 (Pays de la Loire)
    'CD44 (Loire Atlantique)
    Selection.Replace What:="COUERON ALO", Replacement:="ALO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NANTES CA", Replacement:="CAN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="INDRE CN", Replacement:="CNI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NANTES CLL AV", Replacement:="CALL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PORNIC CN", Replacement:="CNPO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FROSSAY CNM AVIRON", Replacement:="CNMA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SUCE SUR ERDRE RC", Replacement:="RCSE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT NAZAIRE SNOS", Replacement:="SNOS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NANTES UNIVERSITE", Replacement:="UNA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD49 (Maine et Loire)
    Selection.Replace What:="ANGERS NA", Replacement:="ANA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CHOLET AS", Replacement:="ASCH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAUMUR LAL", Replacement:="LAL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAUMUR SN", Replacement:="SNSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD53 (Mayenne)
    Selection.Replace What:="CHATEAU GONTIER CNCG", Replacement:="CNCG", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LAVAL CN AVIRON", Replacement:="CNLA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD72 (Sarthe)
    Selection.Replace What:="LE MANS SA", Replacement:="LMSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SABLE NA", Replacement:="SANA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD85 (Vendée)
    Selection.Replace What:="LA ROCHE SUR YON AV 85", Replacement:="LRSY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BOCAGE AC", Replacement:="ACBV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NOIRMOUTIER DNN", Replacement:="DNNO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR13 (SUD PACA)
    'CD04 (Alpes de Haute Provence)
    Selection.Replace What:="SAINTE CROIX AVN 04", Replacement:="AVN4", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MANOSQUE AC", Replacement:="ACDM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ESPARRON DE VERDON CN", Replacement:="CNEV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD05 (Hautes Alpes)
    Selection.Replace What:="SAVINES LE LAC ASP", Replacement:="ASP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EMBRUN CA", Replacement:="CAEM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD06 (Alpes Maritimes)
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
    'CD13 (Bouches du Rhône)
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
    'CD83 (Var)
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
    'CD84 (Vaucluse)
    Selection.Replace What:="AVIGNON SN", Replacement:="SNAV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CADEROUSSE SN", Replacement:="SNCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'CD98 (Monaco)
    Selection.Replace What:="MONACO SN", Replacement:="SNMO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR14 (Guadeloupe)
    'CD9A (Guadeloupe)
    Selection.Replace What:="FERRY ASC", Replacement:="ASCF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LAMENTIN ANCL", Replacement:="ANCL", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PETITBOURG CAC", Replacement:="CTBA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BASSE-TERRE CN", Replacement:="CNBT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR15 (Martinique)
    'CD9B (Martinique)
    Selection.Replace What:="ASCN LA FREGATE", Replacement:="ACNF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DLO FERE ANSES D'ARLET", Replacement:="ADF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="BELLEFONTAINE AMB", Replacement:="ALMB", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SCHOELCHER AC 233", Replacement:="A233", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LE ROBERT AC", Replacement:="ACR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="STE LUCE AV", Replacement:="ALUC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SCHOELCHER CN", Replacement:="CNSC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ETOILE DU NORD CA", Replacement:="CAEN", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RIVIERE-PILOTE CN", Replacement:="CNRP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TROIS ILETS CN", Replacement:="CNTI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FORT DE FRANCE H2O", Replacement:="H2O", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="RIVIERE SALEE LNT", Replacement:="LNT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LE FRANCOIS NSEYTD", Replacement:="NSEY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR16 (Guyane)
    'CD9C (Guyane)
    Selection.Replace What:="MONTSINERY TONNEGRANDE ENMT", Replacement:="ENMT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR17 (La Réunion)
    'CD9D (La Réunion)
    Selection.Replace What:="SAINT GILLES LES BAINS BNO", Replacement:="BNO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LE PORT BN", Replacement:="BNMA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ST PAUL ACR", Replacement:="SACR", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    'LR23 (Polynésie Française)
    'CD9E (Polynésie Française)
    Selection.Replace What:="ASCUP", Replacement:="ASUP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CAP MARARA TAHITI", Replacement:="CAPM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PIRAE FHIP", Replacement:="FHIP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PAPEETE FCH", Replacement:="FCH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MOOREA LPPA", Replacement:="LPPA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FAAA MH", Replacement:="MHOE", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NUKU HIVA NHRC", Replacement:="NHRC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PUNAAUIA SDAPT", Replacement:="SDAP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub


