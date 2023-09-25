Attribute VB_Name = "Module1"
Sub FiltreCourse()
Attribute FiltreCourse.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FiltreCourse Macro
'

'
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$28").AutoFilter Field:=1, Criteria1:="=C01-H1", _
        Operator:=xlOr, Criteria2:="=C01-H3"
    ActiveWorkbook.Save
End Sub
Sub Test_Filtre()
Attribute Test_Filtre.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Test_Filtre Macro
'

'
Dim course1 As String, course2 As String, course3 As String, course4 As String, course5 As String
Dim course6 As String, course7 As String, course8 As String, course9 As String, course10 As String
Dim course11 As String, course12 As String, course13 As String, course14 As String, course15 As String
With Sheets("Stockage Impressions")
        course1 = .Range("A1").Value
        course2 = .Range("B1").Value
        course3 = .Range("C1").Value
        course4 = .Range("D1").Value
        course5 = .Range("E1").Value
        course6 = .Range("F1").Value
        course7 = .Range("G1").Value
        course8 = .Range("H1").Value
        course9 = .Range("I1").Value
        course10 = .Range("J1").Value
        course11 = .Range("K1").Value
        course12 = .Range("L1").Value
        course13 = .Range("M1").Value
        course14 = .Range("N1").Value
        course15 = .Range("O1").Value
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
Attribute TestCopieTirages.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TestCopieTirages Macro
'

'
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
