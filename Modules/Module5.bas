Attribute VB_Name = "Module5"
Sub testtri()
Attribute testtri.VB_ProcData.VB_Invoke_Func = " \n14"
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
Attribute IfAleaBeforeNon.VB_ProcData.VB_Invoke_Func = " \n14"
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
