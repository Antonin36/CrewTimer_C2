Attribute VB_Name = "Module7"
Sub RandomTirages()
Attribute RandomTirages.VB_ProcData.VB_Invoke_Func = " \n14"
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
