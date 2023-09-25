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
