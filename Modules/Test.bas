Attribute VB_Name = "Test"
Sub SupprConn()
Attribute SupprConn.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopieTirages Macro
'

'
   Dim conConnect As WorkbookConnection
    For Each conConnect In ThisWorkbook.Connections
        With conConnect
                conConnect.Delete
        End With
    Next conConnect
End Sub
Sub SupprTirages()
Attribute SupprTirages.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SupprTirages Macro
'

'
    Sheets("Préparation Tirages").Select
    Range("A2:K29").Select
    Selection.EntireRow.Delete
    Range("A1").Select
End Sub
Sub FiltreCourse()
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$28").AutoFilter Field:=1, Criteria1:="=C01-H1", _
        Operator:=xlOr, Criteria2:="=C01-H3"
    ActiveWorkbook.Save
End Sub
Sub Test_Filtre()
Dim course1 As String, course2 As String, course3 As String, course4 As String, course5 As String
Dim course6 As String, course7 As String, course8 As String, course9 As String, course10 As String
Dim course11 As String, course12 As String, course13 As String, course14 As String, course15 As String
With Sheets("Stockage Impressions")
        course1 = .Range("A1").value
        course2 = .Range("B1").value
        course3 = .Range("C1").value
        course4 = .Range("D1").value
        course5 = .Range("E1").value
        course6 = .Range("F1").value
        course7 = .Range("G1").value
        course8 = .Range("H1").value
        course9 = .Range("I1").value
        course10 = .Range("J1").value
        course11 = .Range("K1").value
        course12 = .Range("L1").value
        course13 = .Range("M1").value
        course14 = .Range("N1").value
        course15 = .Range("O1").value
    End With

    With Sheets("Import Tirages")
        .AutoFilterMode = False
        .Range("$A$1:$EA$999").AutoFilter Field:=1, Criteria1:=Array(course1, course2, course3, course4, course5, _
            course6, course7, course8, course9, course10, course11, course12, course13, course14, course15), _
            Operator:=xlFilterValues
        .Range("A1").Select
    End With
    'Sheets("Impressions CT").Select
End Sub

Sub TestCopieTirages()
    Sheets("Import Tirages").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Feuille CrewTimer").Select
    Range("A7:K35").Select
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Import Tirages").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("A:A").EntireColumn.AutoFit
    Range("I1").Select
End Sub
Sub ImpTirages()
'
' ImpTirages Macro
'

'
    Sheets("Import Tirages").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Feuille CrewTimer").Select
    Range("A7:K999").Select
    Selection.Copy
    Sheets("Import Tirages").Select
    Range("A1").Select
    ActiveSheet.Paste
    'insérer filtre
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2:I999").Select
    Selection.Copy
    Sheets("Impressions Tirages CT ").Select
    Range("A13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Unload Me
End Sub
Sub formatResultat()
'
' formatResultat Macro
'

'
    Sheets("Import Resultats").Select
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("A1:H999").Select
    Selection.Copy
    Sheets("Impressions Résultats CT").Select
    Range("A13:H999").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A13").Select
    Sheets("Import Resultats").Select
    Range("A1:H999").Select
    Selection.Copy
    Sheets("Impressions Résultats CT").Select
    Range("A13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub EffImpr()
'
' EffImpr Macro
'

'
    Sheets("Impressions Tirages CT").Select
    Range("A13:H420").Select
    Selection.ClearContents
    Range("A13").Select
    Sheets("Import Resultats CT").Select
End Sub
Sub testtri()
'
' testtri Macro
'

'
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key _
        :=Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Préparation Tirages CT").Select
End Sub
Sub IfAleaBeforeNon()
'
' IfAleaBeforeNon Macro
'

'
    Sheets("Préparation Tirages CT").Select
    Columns("J:J").Select
    Selection.Cut
    Columns("M:M").Select
    ActiveSheet.Paste
    Range("G2:L999").Select
    Range("L2").Activate
    ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort.SortFields.Add2 Key _
        :=Range("L2:L999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Préparation Tirages CT").Sort
        .SetRange Range("G2:L999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("M:M").Select
    Selection.Cut
    Columns("J:J").Select
    ActiveSheet.Paste
    Range("A2").Select
End Sub
Sub EnleverLigneInutiles()
'
' EnleverLigneInutiles Macro
'

'
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
End Sub
Sub RandomTirages()
'
' RandomTirages Macro
'

'
    Sheets("Feuille CrewTimer").Select
    Range("M7").Select
    ActiveCell.FormulaR1C1 = "Random"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "=RAND()"
    Range("M8").Select
    Selection.AutoFill Destination:=Range("M8:M1000"), Type:=xlFillDefault
    Range("M8:M1000").Select
    Rows("8:1048576").Select
    ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort.SortFields.Add2 Key:= _
        Range("A8:A1000"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort.SortFields.Add2 Key:= _
        Range("B8:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort.SortFields.Add2 Key:= _
        Range("M8:M1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort
        .SetRange Range("A7:N1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub

Sub RandTirages()
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

End Sub
Sub DistanceEntreSautsDePageHorizontal()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Impressions Résultats CT") ' Ajustez selon votre feuille
    ThisWorkbook.Sheets("Impressions Résultats CT").Select
    ws.Activate
    
    If ws.HPageBreaks.Count >= 2 Then
        ' Obtenez les positions des deux sauts de page horizontaux
        Dim firstBreak As Double
        Dim secondBreak As Double
        firstBreak = ws.HPageBreaks(1).Location
        secondBreak = ws.HPageBreaks(2).Location
        
        ' Calculez la distance entre les deux sauts de page en points
        Dim distanceEnPoints As Double
        distanceEnPoints = secondBreak - firstBreak
        
        ' Affichez la distance en points dans la fenêtre de l'éditeur VBA
        Debug.Print "Premier Saut de Page Horizontal : " & firstBreak
        Debug.Print "Deuxième Saut de Page Horizontal : " & secondBreak
        Debug.Print "Distance entre les sauts de page horizontaux : " & distanceEnPoints & " points"
    Else
        MsgBox "Il n'y a pas suffisamment de sauts de page horizontaux pour calculer la distance.", vbExclamation
    End If
End Sub

Sub CompterValeursUniques()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim uniqueValue As Variant

    ' Remplacez "Feuil1" par le nom de votre feuille
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ' Remplacez "A" par la colonne que vous souhaitez traiter
    Set rng = ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' Utiliser un dictionnaire pour stocker les valeurs uniques et leur compte
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parcourir chaque cellule dans la plage
    For Each cell In rng
        If cell.value <> "" Then
            ' Incrémenter le compteur pour la valeur unique dans le dictionnaire
            If dict.Exists(cell.value) Then
                dict(cell.value) = dict(cell.value) + 1
            Else
                ' Ajouter la valeur au dictionnaire avec un compteur initial de 1
                dict.Add cell.value, 1
            End If
        End If
    Next cell

    ' Écrire les résultats deux colonnes à droite des données d'origine
    For Each uniqueValue In dict.Keys
        ' Trouver la première cellule vide deux colonnes à droite
        ' Trouver la première cellule vide deux colonnes à droite
        Dim emptyCell As Range
        Set emptyCell = ws.Cells(ws.Rows.Count, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 2).End(xlUp).Offset(1, 0)


        ' Écrire la valeur unique
        emptyCell.value = uniqueValue
        ' Écrire le compteur correspondant
        emptyCell.Offset(0, 1).value = dict(uniqueValue)
    Next uniqueValue
End Sub



