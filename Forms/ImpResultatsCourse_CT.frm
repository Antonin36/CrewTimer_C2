VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImpResultatsCourse_CT 
   Caption         =   "Impression des R�sultats"
   ClientHeight    =   5640
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   7980
   OleObjectBlob   =   "ImpResultatsCourse_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImpResultatsCourse_CT"
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
    Set wsSource = ThisWorkbook.Sheets("Import Resultats CT")
    ' Remplacez "Feuil2" par le nom de votre feuille destination
    Set wsDestination = ThisWorkbook.Sheets("Impressions R�sultats CT")
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
    
    ThisWorkbook.Sheets("Impressions R�sultats CT").Select
    LigneImp = 13
    ' Parcourir les lignes de la feuille source
    For i = 8 To lastRowSource
        ' V�rifier si la valeur dans la colonne C fait partie des valeurs s�lectionn�es
        If IsInArray(wsSource.Cells(i, 2).value, selectedValues) Then
            ' Trouver la derni�re ligne avec des donn�es dans la feuille destination
            lastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
            
            ' Copier la ligne de la feuille source � la feuille destination
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 1).value = Sheets("Import Resultats CT").Cells(i, 1).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 2).value = Sheets("Import Resultats CT").Cells(i, 2).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 3).value = Sheets("Import Resultats CT").Cells(i, 3).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 4).value = Sheets("Import Resultats CT").Cells(i, 4).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 5).value = Sheets("Import Resultats CT").Cells(i, 5).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 6).value = Sheets("Import Resultats CT").Cells(i, 6).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 7).value = Sheets("Import Resultats CT").Cells(i, 7).value
            Sheets("Impressions R�sultats CT").Cells(LigneImp, 8).value = Sheets("Import Resultats CT").Cells(i, 8).value
            LigneImp = LigneImp + 1
        End If
    Next i
    Sheets("Impressions R�sultats CT").Select
    Range("G13:G999").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("H13:H999").Select
    Selection.NumberFormat = "mm:ss.000"
    Range("H13:H999").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("H13:H999").Select
    Selection.NumberFormat = "mm:ss.000"
    Range("A1").Select
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
    Set ws = ThisWorkbook.Sheets("Import Resultats CT")
    ' Remplacez "A" par la colonne que vous souhaitez utiliser
    Set listBox = Me.TableauCourses

    ' Effacer les �l�ments existants dans la ListBox
    listBox.Clear

    ' Utiliser un dictionnaire pour stocker les valeurs uniques
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parcourir chaque cellule dans la plage
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
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




