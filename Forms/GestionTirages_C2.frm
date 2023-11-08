VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestionTirages_C2 
   Caption         =   "Gestion des Tirages"
   ClientHeight    =   7770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   18020
   OleObjectBlob   =   "GestionTirages_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestionTirages_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private alea As Boolean
Private numCollection As New Collection ' D�clarez une collection pour stocker les num�ros de ligne

' Fonction pour ajouter un num�ro de ligne � la collection
Sub AddToCollection(col As Collection, item As Long)
    On Error Resume Next
    col.Add item, CStr(item) ' Utilisez CStr pour convertir le num�ro de ligne en une cl� de cha�ne unique
    On Error GoTo 0
End Sub
Function IsInCollection(col As Collection, val As Long) As Boolean
    On Error Resume Next
    Dim item As Variant
    IsInCollection = False
    For Each item In col
        If item = val Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    On Error GoTo 0
End Function

' Fonction pour vider la collection
Sub ClearCollection(col As Collection)
    Set col = New Collection
End Sub
Private Sub CreationTirages_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim LastRow As Long
    Dim partants As Long
    Dim numlignegoal As Long
    Dim Equipage As String
    Dim rg As Range
    Dim i As Long, j As Long, k As Long, l As Long
    Dim cat As Long
    Dim trigramme As String
    partants = 0
    numlignegoal = 2
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ClearCollection numCollection
    'Trouver la derni�re Ligne Utilis�e en Colonne A de la Feuille Origine
    LastRow = Sheets("Programme des Courses CT").Cells(Sheets("Programme des Courses CT").Rows.Count, "A").End(xlUp).Row

    'Trouver la derni�re ligne non utilis�e en colonne A de la Feuille Destinataire
    j = Sheets("Pr�paration Tirages CT").Cells(Sheets("Pr�paration Tirages CT").Rows.Count, "A").End(xlUp).Row + 1
    Dim limgoal As Long
    limgoal = Sheets("Import GOAL CT").Cells(Sheets("Import GOAL CT").Rows.Count, "C").End(xlUp).Row + 1
    
    'Coller Chaque Ligne contenant Oui en H
    For i = 1 To LastRow
            If Sheets("Programme des Courses CT").Cells(i, 8).Value = "Oui" Then
                partants = 0
                Equipage = ""
                trigramme = ""
                numlignegoal = 2
                'numlignegoal = numlignegoal + 1
                Do While partants < Sheets("R�glages R�gate").Range("E14").Value
                    Sheets("Programme des Courses CT").Rows(i).Copy Destination:=Worksheets("Pr�paration Tirages CT").Range("A" & j)
                    Sheets("Pr�paration Tirages CT").Cells(j, 1).Value = Sheets("Pr�paration Tirages CT").Cells(j, 7)
                    Dim A As String
                    A = Sheets("Pr�paration Tirages CT").Cells(j, 3).Value & "_" & Sheets("Pr�paration Tirages CT").Cells(j, 4).Value
                    Dim B As String
                    B = Sheets("Pr�paration Tirages CT").Cells(j, 6).Value & "_" & Sheets("Pr�paration Tirages CT").Cells(j, 4).Value
                    Sheets("Pr�paration Tirages CT").Cells(j, 3).Value = A
                    Sheets("Pr�paration Tirages CT").Cells(j, 4).Value = B
                    Sheets("Pr�paration Tirages CT").Cells(j, 5).Value = A
                    Sheets("Pr�paration Tirages CT").Cells(j, 6).Value = Sheets("Pr�paration Tirages CT").Cells(j, 9).Value
                    Dim u As Integer
                    For u = 10 To 50
                    Do
                    If Not IsInCollection(numCollection, numlignegoal) Then
                        Exit Do ' Sort de la boucle Do si numlignegoal n'est pas dans la collection
                    Else
                        numlignegoal = numlignegoal + 1 ' Incr�mente numlignegoal
                    End If
                    Loop
                    Dim casegoal As String
                    Dim casetirage As String
                    casegoal = Sheets("Import GOAL CT").Cells(numlignegoal, 3).Value
                    casetirage = Sheets("Pr�paration Tirages CT").Cells(j, u).Value
                    If casegoal = casetirage Then Exit For
                    
                    If u = 50 Then numlignegoal = numlignegoal + 1
                    If numlignegoal = limgoal Then partants = partants + 1
                    Next u
                   
                    If casegoal = casetirage Then
                        Equipage = Sheets("Import GOAL CT").Cells(numlignegoal, 5).Value & " (" & Sheets("Import GOAL CT").Cells(numlignegoal, 6).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 7).Value
                        If Sheets("Import GOAL CT").Cells(numlignegoal, 18).Value <> "" Then
                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 18).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 19).Value
                            If Sheets("Import GOAL CT").Cells(numlignegoal, 30).Value <> "" Then
                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 30).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 31).Value
                                If Sheets("Import GOAL CT").Cells(numlignegoal, 42).Value <> "" Then
                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 42).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 43).Value
                                    If Sheets("Import GOAL CT").Cells(numlignegoal, 54).Value <> "" Then
                                        Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 54).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 55).Value
                                        If Sheets("Import GOAL CT").Cells(numlignegoal, 66).Value <> "" Then
                                            Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 66).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 67).Value
                                            If Sheets("Import GOAL CT").Cells(numlignegoal, 78).Value <> "" Then
                                                Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 78).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 79).Value
                                                If Sheets("Import GOAL CT").Cells(numlignegoal, 90).Value <> "" Then
                                                    Equipage = Equipage & " / " & Sheets("Import GOAL CT").Cells(numlignegoal, 90).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 91).Value
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Sheets("Import GOAL CT").Cells(numlignegoal, 104).Value <> "" Then
                            Equipage = Equipage & " / Bar : " & Sheets("Import GOAL CT").Cells(numlignegoal, 104).Value & " " & Sheets("Import GOAL CT").Cells(numlignegoal, 105).Value
                        End If
                        Equipage = Equipage & ")"
                        Sheets("Pr�paration Tirages CT").Cells(j, 7).Value = Equipage
                        Sheets("Pr�paration Tirages CT").Cells(j, 8).Value = Sheets("Import GOAL CT").Cells(numlignegoal, 5).Value
                        Sheets("Pr�paration Tirages CT").Cells(j, 9).Value = Sheets("Import GOAL CT").Cells(numlignegoal, 3).Value
                        Sheets("Pr�paration Tirages CT").Cells(j, 11).Value = Sheets("Import GOAL CT").Cells(numlignegoal, 5).Value
                        If Sheets("R�glages R�gate").Range("E16").Value = "Rivi�re" Then
                            If Sheets("R�glages R�gate").Range("G16").Value = "TDR" Then
                            Sheets("Pr�paration Tirages CT").Cells(j, 10) = Sheets("Import GOAL CT").Cells(numlignegoal, 4).Value
                            Else
                            Sheets("Pr�paration Tirages CT").Cells(j, 10).Value = partants + 1
                            End If
                            numCollection.Add numlignegoal
                            numlignegoal = 2
                            j = j + 1
                            partants = partants + 1
                            casegoal = ""
                            casetirage = ""
                        Else
                            Sheets("Pr�paration Tirages CT").Cells(j, 10) = Sheets("Import GOAL CT").Cells(numlignegoal, 4).Value
                            numCollection.Add numlignegoal
                            numlignegoal = 2
                            j = j + 1
                            partants = partants + 1
                            casegoal = ""
                            casetirage = ""
                        End If
                    End If
                Loop
                
            End If
    Next i
    
    Sheets("Pr�paration Tirages CT").Select
    Columns("H:H").Select
    Selection.Replace What:="SAINTE CROIX AVN 04", Replacement:="AVN4", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MANOSQUE AC", Replacement:="ACDM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ESPARRON DE VERDON CN", Replacement:="CNEV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAVINES LE LAC ASP", Replacement:="ASP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EMBRUN CA", Replacement:="CAEM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="NICE CN", Replacement:="CNNI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CANNES MANDELIEU RCCM", Replacement:="RCCM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MENTON SCA", Replacement:="SCAM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VILLEFRANCHE SN", Replacement:="SNVI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE ASPTT", Replacement:="AMSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE AAS", Replacement:="AAS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ISTRES ANO", Replacement:="ANOI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CASSIS AC", Replacement:="ACDC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE CA", Replacement:="CAM", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT CHAMAS CASC", Replacement:="CASC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PEYROLLES CNPA", Replacement:="CNPA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARIGNANE CMS", Replacement:="CMSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARTIGUES AC", Replacement:="MAAC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MARSEILLE RC", Replacement:="RCMA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA CIOTAT SN", Replacement:="SNLC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VITROLLES SA", Replacement:="VSA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="HYERES ASPTT", Replacement:="ASAH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SIX OURS ACSF", Replacement:="ACSF", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="VERDON AC", Replacement:="ACV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINT CASSIEN ASC", Replacement:="ASTC", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="LA SEYNE SUR MER AV", Replacement:="ASEY", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="TOULON AV", Replacement:="ATON", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SAINTE MAXIME CA", Replacement:="CAMX", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="SALETTES CN", Replacement:="CNDS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="AVIGNON SN", Replacement:="SNAV", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="CADEROUSSE SN", Replacement:="SNCA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="MONACO SN", Replacement:="SNMO", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("A1").Select
    Dim LastRow2 As Long
    Dim partants2 As Long
   'Find last used row in a Column A of Sheet1
      LastRow2 = Sheets("Programme des Courses CT").Cells(Sheets("Programme des Courses CT").Rows.Count, "A").End(xlUp).Row

   'Find first row where values should be posted in Sheet2
      k = Sheets("Pr�paration Tirages CT").Cells(Sheets("Pr�paration Tirages CT").Rows.Count, "A").End(xlUp).Row + 1
 
   
   'Paste each row that contains "Mavs" in column A of Sheet1 into Sheet2
   For l = 1 To LastRow2
           If Sheets("Programme des Courses CT").Cells(l, 8).Value = "Non" Then
           partants2 = 0
           Do While partants2 < Sheets("R�glages R�gate").Range("E14").Value
               Sheets("Programme des Courses CT").Rows(l).Copy Destination:=Worksheets("Pr�paration Tirages CT").Range("A" & k)
                Sheets("Pr�paration Tirages CT").Cells(k, 1).Value = Sheets("Pr�paration Tirages CT").Cells(k, 7)
                Dim C As String
                C = Sheets("Pr�paration Tirages CT").Cells(k, 3).Value & "_" & Sheets("Pr�paration Tirages CT").Cells(k, 4).Value
                Dim D As String
                D = Sheets("Pr�paration Tirages CT").Cells(k, 6).Value & "_" & Sheets("Pr�paration Tirages CT").Cells(k, 4).Value
                Sheets("Pr�paration Tirages CT").Cells(k, 3).Value = C
                Sheets("Pr�paration Tirages CT").Cells(k, 4).Value = D
                Sheets("Pr�paration Tirages CT").Cells(k, 5).Value = C
                Sheets("Pr�paration Tirages CT").Cells(k, 6).Value = Sheets("Pr�paration Tirages CT").Cells(k, 9).Value
                Sheets("Pr�paration Tirages CT").Cells(k, 7).Value = "TBD"
                Sheets("Pr�paration Tirages CT").Cells(k, 8).Value = "TBD"
                Sheets("Pr�paration Tirages CT").Cells(k, 9).Value = "TBD"
                Sheets("Pr�paration Tirages CT").Cells(k, 10).Value = partants2 + 1
                Sheets("Pr�paration Tirages CT").Cells(k, 11).Value = ""
                k = k + 1
                partants2 = partants2 + 1
            Loop
            End If
   Next l
   Dim LastRow3 As Long
   'Find last used row in a Column A of Sheet1
    LastRow3 = Sheets("Pr�paration Tirages CT").Cells(Sheets("Pr�paration Tirages CT").Rows.Count, "A").End(xlUp).Row
      
   For w = 1 To LastRow3
           If Sheets("Pr�paration Tirages CT").Cells(l, 8).Value = "Non" Or Sheets("Pr�paration Tirages CT").Cells(w, 8).Value = "Oui" Then
           Sheets("Pr�paration Tirages CT").Rows(w).EntireRow.Delete
           End If
   Next w
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
       MsgBox "Les tirages ont �t� cr��s avec succ�s !", vbOKOnly + vbInformation, "Tirages Cr��s"
       
       Sheets("Pr�paration Tirages CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Pr�paration Tirages CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Pr�paration Tirages CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Pr�paration Tirages CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Pr�paration Tirages CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Gestion CrewTimer").Select
            Sheets("Programme des Courses CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Pr�paration Tirages CT").Select
        Rows("1:1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$Z$403").AutoFilter Field:=7, Criteria1:="( )"
        Rows("2:1048576").Select
        Selection.Delete Shift:=xlUp
        Rows("1:1").Select
        ActiveSheet.ShowAllData
        Selection.AutoFilter
        Range("A1").Select
        If Sheets("R�glages R�gate").Range("G16").Value = "Rand" Then
        Sheets("R�glages R�gate").Select
        Sheets("R�glages R�gate").Range("G16").Value = ""
        Sheets("Feuille CrewTimer").Select
        ' G�n�rer des valeurs al�atoires dans la colonne M
        ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("M8:M1000").FormulaR1C1 = "=RAND()"

        ' Tri des donn�es
        With ActiveWorkbook.Worksheets("Feuille CrewTimer").Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("A8:A1000"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
                :="Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", DataOption:=xlSortNormal
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("B8:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("M8:M1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            .SetRange ActiveWorkbook.Worksheets("Feuille CrewTimer").Range("A7:N1000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        ' Suppression de la colonne M
        ActiveWorkbook.Worksheets("Feuille CrewTimer").Columns("M:M").Delete Shift:=xlToLeft
        End If
        Sheets("Gestion CrewTimer").Select
End Sub
Private Sub UserForm_Initialize()
alea = False
Dim random_method As String
random_method = ""
Sheets("R�glages R�gate").Select
            Sheets("R�glages R�gate").Range("G16").Value = ""
    Sheets("Gestion CrewTimer").Select
If MsgBox("Voulez-vous utiliser un tirage al�atoire ?", vbYesNo + vbQuestion, "Tirages Al�atoires ?") = vbYes Then
alea = True
'Mettre cr�er une colonne random, en ER
    Sheets("R�glages R�gate").Select
    Sheets("R�glages R�gate").Range("G16").Value = "Rand"
    Sheets("Import GOAL CT").Select
    random_method = "Al�atoire"
    Sheets("Import GOAL CT").Select
    Range("ER1").Value = "Random"
    Range("ER2").Select
    Dim rand As Long
    rand = 998
    For rand = 1 To rand
    ActiveCell.Value = Rnd()
    ActiveCell.Offset(1, 0).Select
Next rand
'Trier la table
ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "ER2:ER999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL CT").Sort
        .SetRange Range("A1:ER999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Effacer la colonne random
    Range("ER1:ER999").Value = ""
    Range("A1").Select
    
    ElseIf MsgBox("Voulez-vous utiliser un tirage o� le num�ro du bateau est l'ordre de d�part ? (T�te de Rivi�re UNIQUEMENT)", vbYesNo + vbQuestion, "Tirages par Num�ro de Bateau ?") = vbYes Then
    'Proc�der au tirage via l'ordre croissant des num�ros de bateau
    random_method = "Par l'ordre croissant des num�ros de bateau"
    Sheets("R�glages R�gate").Select
    Sheets("R�glages R�gate").Range("G16").Value = "TDR"
        Sheets("Import GOAL CT").Select
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "D2:D999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL CT").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With
    
    Else
    'Tirage Ordre Alphab�tique Nom Court
    MsgBox "Le Tirage va �tre effectu� dans l'ordre alphab�tique des noms courts des clubs.", vbOKOnly + vbInformation, "Tirage Normal"
    random_method = "Par l'ordre alphab�tique des noms courts des clubs"
    Sheets("Import GOAL CT").Select
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL CT").Sort.SortFields.Add2 Key:=Range( _
        "E2:E999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL CT").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Prg Courses selon Ordre Alphab�tique Cat�g
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    If MsgBox("Le mode de tirage d�fini est : " + random_method + ". Confirmez-vous ce choix ?", vbYesNo + vbInformation, "Confirmation Mode de Tirage") = vbYes Then
    ' Feuille � S�lectionner
    Sheets("Pr�paration Tirages CT").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
    Sheets("Gestion CrewTimer").Select
    Exit Sub
    Else
    Call UserForm_Initialize
    End If
    Sheets("Programme des Courses CT").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Feuille � S�lectionner
    Sheets("Pr�paration Tirages CT").Select
    ' Champs � Afficher (Ne pas oublier de d�clarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
        Sheets("Gestion CrewTimer").Select
End Sub

Private Sub ValidTirages_Click()
Dim answer1 As Integer
answer1 = MsgBox("Confirmez-vous la validation des tirages ?", vbYesNo + vbExclamation, "Confirmation Validation Tirages")
  If answer1 = vbYes Then
  Sheets("Pr�paration Tirages CT").Select
    Range("A2:K999").Select
    Selection.Copy
    Sheets("Feuille CrewTimer").Select
    Range("A8").Select
    ActiveSheet.Paste
    Sheets("Gestion CrewTimer").Select
    Sheets("Programme des Courses CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("R�glages R�gate").Select
            Sheets("R�glages R�gate").Range("G16").Value = ""
            Sheets("Gestion CrewTimer").Select
    MsgBox "Les tirages ont bien �t� valid�s et transf�r�s dans la table pour l'export vers CrewTimer !", vbOKOnly + vbInformation, "Tirages Valid�s"
    Unload Me
  Else
    Exit Sub
  End If
End Sub
Private Sub SupprTirages_Click()
Dim answer2 As Integer
answer2 = MsgBox("Confirmez-vous l'invalidation des tirages ?", vbYesNo + vbExclamation, "Confirmation Invalidation Tirages")
  If answer2 = vbYes Then
    Sheets("Pr�paration Tirages CT").Select
    Range("A2:K999").Select
    Selection.EntireRow.Delete
    Range("A1").Select
    MsgBox "Les tirages ont bien �t� invalid�s.", vbOKOnly + vbInformation, "Tirages Invalid�s"
    Call UserForm_Initialize
  Else
    Exit Sub
  End If
End Sub
Private Sub Quit_Click()
Sheets("Programme des Courses CT").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses CT").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses CT").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("R�glages R�gate").Select
            Sheets("R�glages R�gate").Range("G16").Value = ""
            Sheets("Gestion CrewTimer").Select
 Unload Me
End Sub

