VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Change(ByVal Destination As Range)
Dim rngDropdown As Range
Dim oldValue As String
Dim newValue As String
Dim DelimiterType As String
DelimiterType = " / "
Dim DelimiterCount As Integer
Dim TargetType As Integer
Dim i As Integer
Dim arr() As String
 
If Destination.Count > 1 Then Exit Sub
On Error Resume Next
 
Set rngDropdown = Cells.SpecialCells(xlCellTypeAllValidation)
On Error GoTo exitError
 
If rngDropdown Is Nothing Then GoTo exitError
If Not Destination.Column = 5 Then GoTo exitError
 
TargetType = 0
    TargetType = Destination.Validation.Type
    If TargetType = 3 Then  ' is validation type is "list"
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        newValue = Destination.value
        Application.Undo
        oldValue = Destination.value
        Destination.value = newValue
        If oldValue <> "" Then
            If newValue <> "" Then
                If oldValue = newValue Or oldValue = newValue & Replace(DelimiterType, " ", "") Or oldValue = newValue & DelimiterType Then ' leave the value if there is only one in the list
                    oldValue = Replace(oldValue, DelimiterType, "")
                    oldValue = Replace(oldValue, Replace(DelimiterType, " ", ""), "")
                    Destination.value = oldValue
                ElseIf InStr(1, oldValue, DelimiterType & newValue) Or InStr(1, oldValue, " " & newValue & DelimiterType) Then
                    arr = Split(oldValue, DelimiterType)
                If Not IsError(Application.Match(newValue, arr, 0)) = 0 Then
                    Destination.value = oldValue & DelimiterType & newValue
                        Else:
                    Destination.value = ""
                    For i = 0 To UBound(arr)
                    If arr(i) <> newValue Then
                        Destination.value = Destination.value & arr(i) & DelimiterType
                    End If
                    Next i
                Destination.value = Left(Destination.value, Len(Destination.value) - Len(DelimiterType))
                End If
                ElseIf InStr(1, oldValue, newValue & Replace(DelimiterType, " ", "")) Then
                    oldValue = Replace(oldValue, newValue, "")
                    Destination.value = oldValue
                Else
                    Destination.value = oldValue & DelimiterType & newValue
                End If
                Destination.value = Replace(Destination.value, Replace(DelimiterType, " ", "") & Replace(DelimiterType, " ", ""), Replace(DelimiterType, " ", "")) ' remove extra commas and spaces
                Destination.value = Replace(Destination.value, DelimiterType & Replace(DelimiterType, " ", ""), Replace(DelimiterType, " ", ""))
                If Destination.value <> "" Then
                    If Right(Destination.value, 2) = DelimiterType Then  ' remove delimiter at the end
                        Destination.value = Left(Destination.value, Len(Destination.value) - 2)
                    End If
                End If
                If InStr(1, Destination.value, DelimiterType) = 1 Then ' remove delimiter as first characters
                    Destination.value = Replace(Destination.value, DelimiterType, "", 1, 1)
                End If
                If InStr(1, Destination.value, Replace(DelimiterType, " ", "")) = 1 Then
                    Destination.value = Replace(Destination.value, Replace(DelimiterType, " ", ""), "", 1, 1)
                End If
                DelimiterCount = 0
                For i = 1 To Len(Destination.value)
                    If InStr(i, Destination.value, Replace(DelimiterType, " ", "")) Then
                        DelimiterCount = DelimiterCount + 1
                    End If
                Next i
                If DelimiterCount = 1 Then ' remove delimiter if last character
                    Destination.value = Replace(Destination.value, DelimiterType, "")
                    Destination.value = Replace(Destination.value, Replace(DelimiterType, " ", ""), "")
                End If
            End If
        End If
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
 
exitError:
  Application.EnableEvents = True
End Sub
 
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
 
End Sub
