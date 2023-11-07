Attribute VB_Name = "Module6"
Sub EnleverLigneInutiles()
Attribute EnleverLigneInutiles.VB_ProcData.VB_Invoke_Func = " \n14"
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
