VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImpInscrits_CT 
   Caption         =   "Impression des Inscrits"
   ClientHeight    =   5640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7980
   OleObjectBlob   =   "ImpInscrits_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImpInscrits_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Annuler_Click()
    Unload Me
End Sub

Private Sub Imprimer_Click()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim listBox As MSForms.listBox
    Dim selectedValues() As Variant
    Dim lastRowSource As Long
    Dim lastRowDestination As Long
    Dim i As Long, j As Long
    Dim Equipage As String
    Dim LigneImp As Long

    ' Remplacez "Feuil1" par le nom de votre feuille source
    Set wsSource = ThisWorkbook.Sheets("Import GOAL CT")
    ' Remplacez "Feuil2" par le nom de votre feuille destination
    Set wsDestination = ThisWorkbook.Sheets("Impressions Inscrits CT")
    ' Remplacez "ListBox1" par le nom de votre ListBox
    Set listBox = Me.TableauCourses

    ' Trouver la dernière ligne avec des données dans la feuille source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' Obtenir les valeurs sélectionnées dans la ListBox
    ReDim selectedValues(0 To listBox.ListCount - 1)
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            selectedValues(j) = listBox.List(i)
            j = j + 1
        End If
    Next i
    
    ThisWorkbook.Sheets("Impressions Inscrits CT").Select
    LigneImp = 13
    ' Parcourir les lignes de la feuille source
    For i = 2 To lastRowSource
        ' Vérifier si la valeur dans la colonne C fait partie des valeurs sélectionnées
        If IsInArray(wsSource.Cells(i, 3).value, selectedValues) Then
            ' Trouver la dernière ligne avec des données dans la feuille destination
            lastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
            ' Copier la ligne de la feuille source à la feuille destination
            Sheets("Impressions Inscrits CT").Cells(LigneImp, 1).value = Sheets("Import GOAL CT").Cells(i, 3).value
            Equipage = Sheets("Import GOAL CT").Cells(i, 5).value & " (" & Sheets("Import GOAL CT").Cells(i, 6).value & " " & Sheets("Import GOAL CT").Cells(i, 7).value
                        If Sheets("Import GOAL CT").Cells(i, 18).value <> "" Then
                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 18).value & " " & Sheets("Import GOAL CT").Cells(i, 19).value
                            If Sheets("Import GOAL CT").Cells(i, 30).value <> "" Then
                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 30).value & " " & Sheets("Import GOAL CT").Cells(i, 31).value
                                If Sheets("Import GOAL CT").Cells(i, 42).value <> "" Then
                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 42).value & " " & Sheets("Import GOAL CT").Cells(i, 43).value
                                    If Sheets("Import GOAL CT").Cells(i, 54).value <> "" Then
                                        Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 54).value & " " & Sheets("Import GOAL CT").Cells(i, 55).value
                                        If Sheets("Import GOAL CT").Cells(i, 66).value <> "" Then
                                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 66).value & " " & Sheets("Import GOAL CT").Cells(i, 67).value
                                            If Sheets("Import GOAL CT").Cells(i, 78).value <> "" Then
                                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 78).value & " " & Sheets("Import GOAL CT").Cells(i, 79).value
                                                If Sheets("Import GOAL CT").Cells(i, 90).value <> "" Then
                                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(i, 90).value & " " & Sheets("Import GOAL CT").Cells(i, 91).value
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Sheets("Import GOAL CT").Cells(i, 102).value <> "" Then
                            Equipage = Equipage & " / Bar : " & Sheets("Import GOAL CT").Cells(i, 102).value & " " & Sheets("Import GOAL CT").Cells(i, 103).value
                        End If
                        Equipage = Equipage & ")"
            Sheets("Impressions Inscrits CT").Cells(LigneImp, 5).value = Equipage
            LigneImp = LigneImp + 1
        End If
    Next i
    Sheets("Réglages Régate").Range("K30").value = "Ferm"
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim uniqueValues() As Variant
    Dim ws As Worksheet
    Dim listBox As MSForms.listBox
    Dim value As Variant
    Dim i As Long
    Dim dict As Object
    Dim cell As Range

    ' Remplacez "Feuil1" par le nom de votre feuille
    Set ws = ThisWorkbook.Sheets("Import GOAL CT")
    ' Remplacez "A" par la colonne que vous souhaitez utiliser
    Set listBox = Me.TableauCourses

    ' Effacer les éléments existants dans la ListBox
    listBox.Clear

    ' Utiliser un dictionnaire pour stocker les valeurs uniques
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parcourir chaque cellule dans la plage
    For Each cell In ws.Range("C1:C" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        If cell.value <> "" Then
            ' Utiliser un dictionnaire pour stocker les valeurs uniques
            dict(cell.value) = 1
        End If
    Next cell

    ' Transférer les clés du dictionnaire dans un tableau
    uniqueValues = Application.Transpose(dict.Keys)

    ' Ajouter les valeurs uniques à la ListBox
    For i = LBound(uniqueValues) To UBound(uniqueValues)
        listBox.AddItem uniqueValues(i, 1)
    Next i
End Sub
Function IsInArray(value As Variant, arr As Variant) As Boolean
    ' Vérifier si une valeur est dans un tableau
    Dim element As Variant
    For Each element In arr
        If element = value Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
