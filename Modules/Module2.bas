Attribute VB_Name = "Module2"
Sub ImpTirages()
Attribute ImpTirages.VB_ProcData.VB_Invoke_Func = " \n14"
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
