VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImpPesee_C2 
   Caption         =   "Impression des Feuilles de Pes�e"
   ClientHeight    =   5640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7980
   OleObjectBlob   =   "ImpPesee_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImpPesee_C2"
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
    Set wsSource = ThisWorkbook.Sheets("Feuille Concept2")
    ' Remplacez "Feuil2" par le nom de votre feuille destination
    Set wsDestination = ThisWorkbook.Sheets("Impressions Pes�e C2")
    ' Remplacez "ListBox1" par le nom de votre ListBox
    Set listBox = Me.TableauCourses

    ' Trouver la derni�re ligne avec des donn�es dans la feuille source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Obtenir les valeurs s�lectionn�es dans la ListBox
    ReDim selectedValues(0 To listBox.ListCount - 1)
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            selectedValues(j) = listBox.List(i)
            j = j + 1
        End If
    Next i
    
    ThisWorkbook.Sheets("Impressions Pes�e C2").Select
    LigneImp = 13
    ' Parcourir les lignes de la feuille source
    For i = 8 To lastRowSource
        ' V�rifier si la valeur dans la colonne C fait partie des valeurs s�lectionn�es
        If IsInArray(wsSource.Cells(i, 4).value, selectedValues) Then
            ' Trouver la derni�re ligne avec des donn�es dans la feuille destination
            lastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
            
            ' Copier la ligne de la feuille source � la feuille destination
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Monday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Lundi"
            End If
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Tuesday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Mardi"
            End If
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Wednesday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Mercredi"
            End If
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Thursday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Jeudi"
            End If
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Friday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Vendredi"
            End If
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Saturday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Samedi"
            End If
            If Sheets("Feuille Concept2").Cells(i, 1).value = "Sunday" Then
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 1).value = "Dimanche"
            End If
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 2).value = Sheets("Feuille Concept2").Cells(i, 2).value
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 3).value = Sheets("Feuille Concept2").Cells(i, 3).value
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 4).value = Sheets("Feuille Concept2").Cells(i, 4).value
            Sheets("Impressions Pes�e C2").Cells(LigneImp, 5).value = Sheets("Feuille Concept2").Cells(i, 7).value
            LigneImp = LigneImp + 1
        End If
    Next i
    Sheets("R�glages R�gate").Range("K30").value = "Ferm"
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
    Set ws = ThisWorkbook.Sheets("Feuille Concept2")
    ' Remplacez "A" par la colonne que vous souhaitez utiliser
    Set listBox = Me.TableauCourses

    ' Effacer les �l�ments existants dans la ListBox
    listBox.Clear

    ' Utiliser un dictionnaire pour stocker les valeurs uniques
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parcourir chaque cellule dans la plage
    For Each cell In ws.Range("D8:D" & ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)
        If cell.value <> "" Then
            ' Utiliser un dictionnaire pour stocker les valeurs uniques
            dict(cell.value) = 1
        End If
    Next cell

    ' Transf�rer les cl�s du dictionnaire dans un tableau
    uniqueValues = Application.Transpose(dict.Keys)

    ' Ajouter les valeurs uniques � la ListBox
    For i = LBound(uniqueValues) To UBound(uniqueValues)
        listBox.AddItem uniqueValues(i, 1)
    Next i
End Sub
Function IsInArray(value As Variant, arr As Variant) As Boolean
    ' V�rifier si une valeur est dans un tableau
    Dim element As Variant
    For Each element In arr
        If element = value Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function




