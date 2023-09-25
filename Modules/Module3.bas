Attribute VB_Name = "Module3"
Sub formatResultat()
Attribute formatResultat.VB_ProcData.VB_Invoke_Func = " \n14"
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
