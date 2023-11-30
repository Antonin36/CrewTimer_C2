VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImpTiragesCateg_CT 
   Caption         =   "Impression des Tirages"
   ClientHeight    =   5640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7980
   OleObjectBlob   =   "ImpTiragesCateg_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImpTiragesCateg_CT"
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
    Dim LigneImp As Long

    ' Remplacez "Feuil1" par le nom de votre feuille source
    Set wsSource = ThisWorkbook.Sheets("Feuille CrewTimer")
    ' Remplacez "Feuil2" par le nom de votre feuille destination
    Set wsDestination = ThisWorkbook.Sheets("Impressions Tirages CT")
    ' Remplacez "ListBox1" par le nom de votre ListBox
    Set listBox = Me.TableauCourses

    ' Trouver la dernière ligne avec des données dans la feuille source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Obtenir les valeurs sélectionnées dans la ListBox
    ReDim selectedValues(0 To listBox.ListCount - 1)
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            selectedValues(j) = listBox.List(i)
            j = j + 1
        End If
    Next i
    
    ThisWorkbook.Sheets("Impressions Tirages CT").Select
    LigneImp = 13
    ' Parcourir les lignes de la feuille source
    For i = 8 To lastRowSource
        ' Vérifier si la valeur dans la colonne C fait partie des valeurs sélectionnées
        If Sheets("Feuille CrewTimer").Cells(i, 9).value = "" Then
        GoTo ProchaineIteration
        End If
        If IsInArray(wsSource.Cells(i, 9).value, selectedValues) Then
            ' Trouver la dernière ligne avec des données dans la feuille destination
            lastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
            
            ' Copier la ligne de la feuille source à la feuille destination
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Monday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Lundi"
            End If
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Tuesday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Mardi"
            End If
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Wednesday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Mercredi"
            End If
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Thursday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Jeudi"
            End If
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Friday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Vendredi"
            End If
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Saturday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Samedi"
            End If
            If Sheets("Feuille CrewTimer").Cells(i, 1).value = "Sunday" Then
            Sheets("Impressions Tirages CT").Cells(LigneImp, 1).value = "Dimanche"
            End If
            Sheets("Impressions Tirages CT").Cells(LigneImp, 2).value = Sheets("Feuille CrewTimer").Cells(i, 2).value
            Sheets("Impressions Tirages CT").Cells(LigneImp, 3).value = Sheets("Feuille CrewTimer").Cells(i, 3).value
            Sheets("Impressions Tirages CT").Cells(LigneImp, 4).value = Sheets("Feuille CrewTimer").Cells(i, 4).value
            Sheets("Impressions Tirages CT").Cells(LigneImp, 5).value = Sheets("Feuille CrewTimer").Cells(i, 7).value
            Sheets("Impressions Tirages CT").Cells(LigneImp, 6).value = Sheets("Feuille CrewTimer").Cells(i, 8).value
            Sheets("Impressions Tirages CT").Cells(LigneImp, 7).value = Sheets("Feuille CrewTimer").Cells(i, 9).value
            Sheets("Impressions Tirages CT").Cells(LigneImp, 8).value = Sheets("Feuille CrewTimer").Cells(i, 10).value
            LigneImp = LigneImp + 1
        End If
ProchaineIteration:
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
    Set ws = ThisWorkbook.Sheets("Feuille CrewTimer")
    ' Remplacez "A" par la colonne que vous souhaitez utiliser
    Set listBox = Me.TableauCourses

    ' Effacer les éléments existants dans la ListBox
    listBox.Clear

    ' Utiliser un dictionnaire pour stocker les valeurs uniques
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parcourir chaque cellule dans la plage
    For Each cell In ws.Range("I8:I" & ws.Cells(ws.Rows.Count, "I").End(xlUp).Row)
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

