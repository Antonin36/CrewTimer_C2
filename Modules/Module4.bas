Attribute VB_Name = "Module4"
Sub EffImpr()
Attribute EffImpr.VB_ProcData.VB_Invoke_Func = " \n14"
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
