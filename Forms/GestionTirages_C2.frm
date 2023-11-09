VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestionTirages_C2 
   Caption         =   "Gestion des Tirages"
   ClientHeight    =   7740
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17980
   OleObjectBlob   =   "GestionTirages_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestionTirages_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private alea As Boolean
Private numCollection As New Collection ' Déclarez une collection pour stocker les numéros de ligne

' Fonction pour ajouter un numéro de ligne à la collection
Sub AddToCollection(col As Collection, Item As Long)
    On Error Resume Next
    col.Add Item, CStr(Item) ' Utilisez CStr pour convertir le numéro de ligne en une clé de chaîne unique
    On Error GoTo 0
End Sub
Function IsInCollection(col As Collection, val As Long) As Boolean
    On Error Resume Next
    Dim Item As Variant
    IsInCollection = False
    For Each Item In col
        If Item = val Then
            IsInCollection = True
            Exit Function
        End If
    Next Item
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
    Dim jsonCell As String
    Dim rameur1 As String
    Dim rameur2 As String
    Dim rameur3 As String
    Dim rameur4 As String
    Dim rameur5 As String
    Dim rameur6 As String
    Dim rameur7 As String
    Dim rameur8 As String
    Dim barreur As String
    partants = 0
    numlignegoal = 2
    Sheets("Programme des Courses C2").Select
    'Columns("F:F").Select
    'ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
        'Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        'xlSortNormal
   ' With ActiveWorkbook.Worksheets("Programme des Courses C2").Sort
        '.SetRange Range("A1:AW999")
        '.Header = xlYes
        '.MatchCase = False
        '.Orientation = xlTopToBottom
        '.SortMethod = xlPinYin
        '.Apply
    'End With
    ClearCollection numCollection
    'Trouver la dernière Ligne Utilisée en Colonne A de la Feuille Origine
    LastRow = Sheets("Programme des Courses C2").Cells(Sheets("Programme des Courses C2").Rows.Count, "A").End(xlUp).Row

    'Trouver la dernière ligne non utilisée en colonne A de la Feuille Destinataire
    j = Sheets("Préparation Tirages C2").Cells(Sheets("Préparation Tirages C2").Rows.Count, "A").End(xlUp).Row + 1
    Dim limgoal As Long
    limgoal = Sheets("Import GOAL C2").Cells(Sheets("Import GOAL C2").Rows.Count, "C").End(xlUp).Row + 1
    
    'Coller Chaque Ligne contenant Oui en H
    For i = 1 To LastRow
    jsonCell = ""
            If Sheets("Programme des Courses C2").Cells(i, 8).Value = "Oui" Then
                partants = 0
                Equipage = ""
                trigramme = ""
                numlignegoal = 2
                If partants = 0 Then
                'Code Intro JSON Var jsoncell
                jsonCell = "{""race_definition"":{""boats"": ["
                End If
                
                'numlignegoal = numlignegoal + 1
                Do While partants < Sheets("Réglages Régate").Range("E14").Value
                    rameur1 = ""
                    rameur2 = ""
                    rameur3 = ""
                    rameur4 = ""
                    rameur5 = ""
                    rameur6 = ""
                    rameur7 = ""
                    rameur8 = ""
                    barreur = ""
                    Sheets("Programme des Courses C2").Rows(i).Copy Destination:=Worksheets("Préparation Tirages C2").Range("A" & j)
                    Dim A As String
                    A = Sheets("Préparation Tirages C2").Cells(j, 3).Value & "_" & Sheets("Préparation Tirages C2").Cells(j, 4).Value
                    Dim B As String
                    B = Sheets("Préparation Tirages C2").Cells(j, 6).Value & "_" & Sheets("Préparation Tirages C2").Cells(j, 4).Value
                    Sheets("Préparation Tirages C2").Cells(j, 3).Value = A
                    Sheets("Préparation Tirages C2").Cells(j, 4).Value = B
                    Sheets("Préparation Tirages C2").Cells(j, 5).Value = A
                    Sheets("Préparation Tirages C2").Cells(j, 6).Value = Sheets("Préparation Tirages C2").Cells(j, 9).Value
                    Sheets("Préparation Tirages C2").Cells(j, 13).Value = Sheets("Préparation Tirages C2").Cells(j, 50).Value
                    Sheets("Préparation Tirages C2").Cells(j, 14).Value = Sheets("Préparation Tirages C2").Cells(j, 51).Value
                    Sheets("Préparation Tirages C2").Cells(j, 15).Value = Sheets("Préparation Tirages C2").Cells(j, 52).Value
                    Sheets("Préparation Tirages C2").Cells(j, 16).Value = Sheets("Préparation Tirages C2").Cells(j, 53).Value
                    Sheets("Préparation Tirages C2").Cells(j, 17).Value = Sheets("Préparation Tirages C2").Cells(j, 54).Value
                    Sheets("Préparation Tirages C2").Cells(j, 18).Value = Sheets("Préparation Tirages C2").Cells(j, 56).Value
                    Sheets("Préparation Tirages C2").Cells(j, 50).Value = ""
                    Sheets("Préparation Tirages C2").Cells(j, 51).Value = ""
                    Sheets("Préparation Tirages C2").Cells(j, 52).Value = ""
                    Sheets("Préparation Tirages C2").Cells(j, 53).Value = ""
                    Sheets("Préparation Tirages C2").Cells(j, 54).Value = ""
                    Sheets("Préparation Tirages C2").Cells(j, 56).Value = ""
                    Dim u As Integer
                    For u = 10 To 50
                    Do
                    If Not IsInCollection(numCollection, numlignegoal) Then
                        Exit Do ' Sort de la boucle Do si numlignegoal n'est pas dans la collection
                    Else
                        numlignegoal = numlignegoal + 1 ' Incrémente numlignegoal
                    End If
                    Loop
                    Dim casegoal As String
                    Dim casetirage As String
                    casegoal = Sheets("Import GOAL C2").Cells(numlignegoal, 3).Value
                    casetirage = Sheets("Préparation Tirages C2").Cells(j, u).Value
                    If casegoal = casetirage Then Exit For
                    
                    If u = 50 Then numlignegoal = numlignegoal + 1
                    If numlignegoal = limgoal Then partants = partants + 1
                    Next u
                   
                    If casegoal = casetirage Then
                        Sheets("Préparation Tirages C2").Cells(j, 8).Value = Sheets("Import GOAL C2").Cells(numlignegoal, 5).Value
                        Sheets("Préparation Tirages C2").Select
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
                        Dim cell As Range
                        Dim tempStr As String
                        Dim carequipage As Integer

                        ' Spécifiez la cellule à partir de laquelle vous souhaitez enlever les nombres
                        Set cell = Sheets("Préparation Tirages C2").Cells(j, 8)

                        tempStr = cell.Value
                        cell.ClearContents ' Efface le contenu de la cellule

                        ' Parcourez chaque caractère de la chaîne
                        For carequipage = 1 To Len(tempStr)
                        If Not (IsNumeric(Mid(tempStr, carequipage, 1)) Or Mid(tempStr, carequipage, 1) = " ") Then
                        ' Si le caractère n'est pas numérique, ajoutez-le à la cellule
                        cell.Value = cell.Value & Mid(tempStr, carequipage, 1)
                        End If
                        Next carequipage
                        
                        jsonCell = jsonCell & "{""affiliation"": """ & Sheets("Préparation Tirages C2").Cells(j, 8).Value & """,""class_name"": """ & Sheets("Import GOAL C2").Cells(numlignegoal, 3).Value & """,""lane_number"": " & partants + 1 & ","
                        Equipage = Sheets("Import GOAL C2").Cells(numlignegoal, 5).Value & " (" & Sheets("Import GOAL C2").Cells(numlignegoal, 6).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 7).Value
                        rameur1 = Sheets("Import GOAL C2").Cells(numlignegoal, 6).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 7).Value
                        If Sheets("Import GOAL C2").Cells(numlignegoal, 18).Value <> "" Then
                            Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 18).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 19).Value
                            rameur2 = Sheets("Import GOAL C2").Cells(numlignegoal, 18).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 19).Value
                            If Sheets("Import GOAL C2").Cells(numlignegoal, 30).Value <> "" Then
                                Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 30).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 31).Value
                                rameur3 = Sheets("Import GOAL C2").Cells(numlignegoal, 30).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 31).Value
                                If Sheets("Import GOAL C2").Cells(numlignegoal, 42).Value <> "" Then
                                    Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 42).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 43).Value
                                    rameur4 = Sheets("Import GOAL C2").Cells(numlignegoal, 42).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 43).Value
                                    If Sheets("Import GOAL C2").Cells(numlignegoal, 54).Value <> "" Then
                                        Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 54).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 55).Value
                                        rameur5 = Sheets("Import GOAL C2").Cells(numlignegoal, 54).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 55).Value
                                        If Sheets("Import GOAL C2").Cells(numlignegoal, 66).Value <> "" Then
                                            Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 66).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 67).Value
                                            rameur6 = Sheets("Import GOAL C2").Cells(numlignegoal, 66).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 67)
                                            If Sheets("Import GOAL C2").Cells(numlignegoal, 78).Value <> "" Then
                                                Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 78).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 79).Value
                                                rameur7 = Sheets("Import GOAL C2").Cells(numlignegoal, 78).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 79).Value
                                                If Sheets("Import GOAL C2").Cells(numlignegoal, 90).Value <> "" Then
                                                    Equipage = Equipage & " / " & Sheets("Import GOAL C2").Cells(numlignegoal, 90).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 91).Value
                                                    rameur8 = Sheets("Import GOAL C2").Cells(numlignegoal, 90).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 91).Value
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Sheets("Import GOAL C2").Cells(numlignegoal, 104).Value <> "" Then
                            Equipage = Equipage & " / Bar : " & Sheets("Import GOAL C2").Cells(numlignegoal, 104).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 105).Value
                            barreur = Sheets("Import GOAL C2").Cells(numlignegoal, 104).Value & " " & Sheets("Import GOAL C2").Cells(numlignegoal, 105).Value
                        End If
                        Equipage = Equipage & ")"
                        Sheets("Préparation Tirages C2").Cells(j, 7).Value = Equipage
                        jsonCell = jsonCell & """name"": """ & Sheets("Préparation Tirages C2").Cells(j, 7).Value & """,""participants"": ["
                        If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" Then
                            If rameur1 <> "" Then
                            jsonCell = jsonCell & "{""name"": """ & rameur1 & """},"
                                If rameur2 <> "" Then
                                jsonCell = jsonCell & "{""name"": """ & rameur2 & """},"
                                    If rameur3 <> "" Then
                                    jsonCell = jsonCell & "{""name"": """ & rameur3 & """},"
                                        If rameur4 <> "" Then
                                        jsonCell = jsonCell & "{""name"": """ & rameur4 & """},"
                                            If rameur5 <> "" Then
                                            jsonCell = jsonCell & "{""name"": """ & rameur5 & """},"
                                                If rameur6 <> "" Then
                                                jsonCell = jsonCell & "{""name"": """ & rameur6 & """},"
                                                    If rameur7 <> "" Then
                                                    jsonCell = jsonCell & "{""name"": """ & rameur7 & """},"
                                                        If rameur8 <> "" Then
                                                        jsonCell = jsonCell & "{""name"": """ & rameur8 & """},"
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        If Len(jsonCell) > 0 Then
                        jsonCell = Left(jsonCell, Len(jsonCell) - 1)
                        End If
                        jsonCell = jsonCell & "]},"
                        End If
                        If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Relais" Then
                        jsonCell = jsonCell & "{""name"": """
                        If rameur1 <> "" Then
                            jsonCell = jsonCell & rameur1
                                If rameur2 <> "" Then
                                jsonCell = jsonCell & " / " & rameur2
                                    If rameur3 <> "" Then
                                    jsonCell = jsonCell & " / " & rameur3
                                        If rameur4 <> "" Then
                                        jsonCell = jsonCell & " / " & rameur4
                                            If rameur5 <> "" Then
                                            jsonCell = jsonCell & " / " & rameur5
                                                If rameur6 <> "" Then
                                                jsonCell = jsonCell & " / " & rameur6
                                                    If rameur7 <> "" Then
                                                    jsonCell = jsonCell & " / " & rameur7
                                                        If rameur8 <> "" Then
                                                        jsonCell = jsonCell & " / " & rameur8
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        jsonCell = jsonCell & """}]},"
                        End If
                        If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Indiv" Then
                        jsonCell = jsonCell & "{""name"": """ & Sheets("Préparation Tirages C2").Cells(j, 7).Value & """}]},"
                        End If
                        Sheets("Préparation Tirages C2").Cells(j, 9).Value = Sheets("Import GOAL C2").Cells(numlignegoal, 3).Value
                        Sheets("Préparation Tirages C2").Cells(j, 11).Value = Sheets("Import GOAL C2").Cells(numlignegoal, 5).Value
                        Sheets("Préparation Tirages C2").Cells(j, 10).Value = partants + 1
                            numCollection.Add numlignegoal
                            numlignegoal = 2
                            j = j + 1
                            partants = partants + 1
                            casegoal = ""
                            casetirage = ""
                            Sheets("Préparation Tirages C2").Select
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
                    End If
                Loop
                If partants = Sheets("Réglages Régate").Range("E14").Value Then
                'Code Fin JSON Var jsoncell
                If Len(jsonCell) > 0 Then
                jsonCell = Left(jsonCell, Len(jsonCell) - 1)
                End If
                jsonCell = jsonCell & "],""c2_race_id"": """","
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Indiv" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Distance" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""meters"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""individual"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Indiv" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Max de Distance" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""individual"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Indiv" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Max de Calories" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""individual calorie score"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Indiv" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Calories" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""calories"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""individual"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Distance" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Moyenne" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""meters"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""avg"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Max de Distance" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Moyenne" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""avg"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Max de Calories" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Moyenne" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team calorie score"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""avg"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Calories" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Moyenne" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""calories"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""avg"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Distance" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Somme" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""meters"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""sum"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Max de Distance" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Somme" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""sum"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Max de Calories" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Somme" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team calorie score"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""sum"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Equipe" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Calories" And Sheets("Programme des Courses C2").Cells(i, 54).Value = "Somme" Then
                jsonCell = jsonCell & """duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""calories"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""team"",""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_scoring"": ""sum"",""team_size"": " & Sheets("Programme des Courses C2").Cells(i, 56).Value & ",""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Relais" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Distance" Then
                jsonCell = jsonCell & """display_prompt_at_splits"": true,""duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""meters"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""relay"",""sound_horn_at_splits"": false,""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Relais" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Temps" Then
                jsonCell = jsonCell & """display_prompt_at_splits"": true,""duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""time"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""relay"",""sound_horn_at_splits"": false,""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                If Sheets("Programme des Courses C2").Cells(i, 52).Value = "Relais" And Sheets("Programme des Courses C2").Cells(i, 53).Value = "Calories" Then
                jsonCell = jsonCell & """display_prompt_at_splits"": true,""duration"": " & Sheets("Programme des Courses C2").Cells(i, 50).Value & ",""duration_type"": ""calories"",""event_name"": """ & Sheets("Réglages Régate").Range("D4").Value & """,""name_long"": """ & Sheets("Préparation Tirages C2").Cells(j - 1, 4).Value & """,""name_short"": ""short name"",""race_id"": """",""race_type"": ""relay"",""sound_horn_at_splits"": false,""split_value"": " & Sheets("Programme des Courses C2").Cells(i, 51).Value & ",""team_size"": 1,""time_cap"": 0}}"
                End If
                
                End If
            End If
            'Implémenter JSON avant reset
    Sheets("Programme des Courses C2").Cells(i, 55).Value = jsonCell
    Next i
    
    Dim LastRow2 As Long
    Dim partants2 As Long
   'Find last used row in a Column A of Sheet1
      LastRow2 = Sheets("Programme des Courses C2").Cells(Sheets("Programme des Courses C2").Rows.Count, "A").End(xlUp).Row

   'Find first row where values should be posted in Sheet2
      k = Sheets("Préparation Tirages C2").Cells(Sheets("Préparation Tirages C2").Rows.Count, "A").End(xlUp).Row + 1
 
   
   'Paste each row that contains "Mavs" in column A of Sheet1 into Sheet2
   For l = 1 To LastRow2
           If Sheets("Programme des Courses C2").Cells(l, 8).Value = "Non" Then
           partants2 = 0
           Do While partants2 < Sheets("Réglages Régate").Range("E14").Value
               Sheets("Programme des Courses C2").Rows(l).Copy Destination:=Worksheets("Préparation Tirages C2").Range("A" & k)
                Dim C As String
                C = Sheets("Préparation Tirages C2").Cells(k, 3).Value & "_" & Sheets("Préparation Tirages C2").Cells(k, 4).Value
                Dim D As String
                D = Sheets("Préparation Tirages C2").Cells(k, 6).Value & "_" & Sheets("Préparation Tirages C2").Cells(k, 4).Value
                Sheets("Préparation Tirages C2").Cells(k, 3).Value = C
                Sheets("Préparation Tirages C2").Cells(k, 4).Value = D
                Sheets("Préparation Tirages C2").Cells(k, 5).Value = C
                Sheets("Préparation Tirages C2").Cells(k, 6).Value = Sheets("Préparation Tirages C2").Cells(k, 9).Value
                Sheets("Préparation Tirages C2").Cells(k, 7).Value = "A Déterminer"
                Sheets("Préparation Tirages C2").Cells(k, 8).Value = "A Déterminer"
                Sheets("Préparation Tirages C2").Cells(k, 9).Value = "A Déterminer"
                Sheets("Préparation Tirages C2").Cells(k, 10).Value = partants2 + 1
                Sheets("Préparation Tirages C2").Cells(k, 11).Value = ""
                k = k + 1
                partants2 = partants2 + 1
            Loop
            End If
   Next l
   Dim LastRow3 As Long
   'Find last used row in a Column A of Sheet1
    LastRow3 = Sheets("Préparation Tirages C2").Cells(Sheets("Préparation Tirages C2").Rows.Count, "A").End(xlUp).Row
      
   For w = 1 To LastRow3
           If Sheets("Préparation Tirages C2").Cells(l, 8).Value = "Non" Or Sheets("Préparation Tirages C2").Cells(w, 8).Value = "Oui" Then
           Sheets("Préparation Tirages C2").Rows(w).EntireRow.Delete
           End If
   Next w
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
       MsgBox "Les tirages ont été créés avec succès !", vbOKOnly + vbInformation, "Tirages Créés"
       
       Sheets("Préparation Tirages C2").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Préparation Tirages C2").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Préparation Tirages C2").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Préparation Tirages C2").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Préparation Tirages C2").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Gestion Concept2").Select
            Sheets("Programme des Courses C2").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses C2").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Préparation Tirages C2").Select
        Rows("1:1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$Z$403").AutoFilter Field:=7, Criteria1:="( )"
        Rows("2:1048576").Select
        Selection.Delete Shift:=xlUp
        Rows("1:1").Select
        ActiveSheet.ShowAllData
        Selection.AutoFilter
        Range("A1").Select
        If Sheets("Réglages Régate").Range("G16").Value = "Rand" Then
        Sheets("Réglages Régate").Select
        Sheets("Réglages Régate").Range("G16").Value = ""
        Sheets("Feuille Concept2").Select
        ' Générer des valeurs aléatoires dans la colonne M
        ActiveWorkbook.Worksheets("Feuille Concept2").Range("M8:M1000").FormulaR1C1 = "=RAND()"

        ' Tri des données
        With ActiveWorkbook.Worksheets("Feuille Concept2").Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille Concept2").Range("A8:A1000"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
                :="Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", DataOption:=xlSortNormal
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille Concept2").Range("B8:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            .SortFields.Add2 Key:=ActiveWorkbook.Worksheets("Feuille Concept2").Range("M8:M1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            .SetRange ActiveWorkbook.Worksheets("Feuille Concept2").Range("A7:N1000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        ' Suppression de la colonne M
        ActiveWorkbook.Worksheets("Feuille Concept2").Columns("M:M").Delete Shift:=xlToLeft
        End If
        Sheets("Gestion Concept2").Select
End Sub
Private Sub UserForm_Initialize()
alea = False
Dim random_method As String
random_method = ""
Sheets("Réglages Régate").Select
Sheets("Réglages Régate").Range("G16").Value = ""
Sheets("Réglages Régate").Range("H16").Value = ""
Sheets("Gestion Concept2").Select
If MsgBox("Voulez-vous utiliser un tirage aléatoire ?", vbYesNo + vbQuestion, "Tirages Aléatoires ?") = vbYes Then
alea = True
'Mettre créer une colonne random, en ER
    Sheets("Réglages Régate").Select
    Sheets("Réglages Régate").Range("G16").Value = "Rand"
    Sheets("Import GOAL C2").Select
    random_method = "Aléatoire"
    Sheets("Import GOAL C2").Select
    Range("ER1").Value = "Random"
    Range("ER2").Select
    Dim rand As Long
    rand = Sheets("Import GOAL C2").Cells(Sheets("Import GOAL C2").Rows.Count, "C").End(xlUp).Row - 1
    For rand = 1 To rand
    ActiveCell.Value = Rnd()
    ActiveCell.Offset(1, 0).Select
Next rand
'Trier la table
ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Add2 Key:=Range( _
        "ER2:ER999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL C2").Sort
        .SetRange Range("A1:ER999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Effacer la colonne random
    Range("ER1:ER999").Select
    Range("ER1:ER999").Value = ""
    Range("A1").Select
    
    ElseIf MsgBox("Voulez-vous utiliser un tirage où le numéro du bateau défini dans GOAL est l'ordre de départ ?", vbYesNo + vbQuestion, "Tirages par Numéro de Bateau ?") = vbYes Then
    'Procéder au tirage via l'ordre croissant des numéros de bateau
    random_method = "Par l'ordre croissant des numéros de bateau"
    Sheets("Réglages Régate").Select
    Sheets("Réglages Régate").Range("H16").Value = "Num"
        Sheets("Import GOAL C2").Select
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Add2 Key:=Range( _
        "D2:D999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL C2").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With
    
    Else
    'Tirage Ordre Alphabétique Nom Court
    MsgBox "Le Tirage va être effectué dans l'ordre alphabétique des noms courts des clubs.", vbOKOnly + vbInformation, "Tirage Normal"
    random_method = "Par l'ordre alphabétique des noms courts des clubs"
    Sheets("Import GOAL C2").Select
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Add2 Key:=Range( _
        "C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Import GOAL C2").Sort.SortFields.Add2 Key:=Range( _
        "E2:E999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Import GOAL C2").Sort
        .SetRange Range("A1:EQ999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Prg Courses selon Ordre Alphabétique Catég
    Sheets("Programme des Courses C2").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses C2").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    If MsgBox("Le mode de tirage défini est : " + random_method + ". Confirmez-vous ce choix ?", vbYesNo + vbInformation, "Confirmation Mode de Tirage") = vbYes Then
    ' Feuille à Sélectionner
    Sheets("Préparation Tirages C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;150;500;1000;50;50;80;200"
    Sheets("Gestion Concept2").Select
    Exit Sub
    Else
    Call UserForm_Initialize
    End If
    Sheets("Programme des Courses C2").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
        Range("F1:F999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Programme des Courses C2").Sort
        .SetRange Range("A1:AW999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Feuille à Sélectionner
    Sheets("Préparation Tirages C2").Select
    ' Champs à Afficher (Ne pas oublier de déclarer le nbre de colonnes dans Properties.
    TableauTirages.RowSource = "A1:K999"
    TableauTirages.ColumnWidths = "50;80;150;400;0;500;1000;50;50;80;200"
        Sheets("Gestion Concept2").Select
End Sub

Private Sub ValidTirages_Click()
Dim answer1 As Integer
answer1 = MsgBox("Confirmez-vous la validation des tirages ?", vbYesNo + vbExclamation, "Confirmation Validation Tirages")
  If answer1 = vbYes Then
  Sheets("Préparation Tirages C2").Select
    Range("A2:R999").Select
    Selection.Copy
    Sheets("Feuille Concept2").Select
    Range("A8").Select
    ActiveSheet.Paste
    Sheets("Gestion Concept2").Select
    Sheets("Programme des Courses C2").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses C2").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Réglages Régate").Select
            Sheets("Réglages Régate").Range("G16").Value = ""
            Sheets("Réglages Régate").Range("H16").Value = ""
            Sheets("Gestion Concept2").Select
    MsgBox "Les tirages ont bien été validés et transférés dans la table pour l'export vers Concept2 !", vbOKOnly + vbInformation, "Tirages Validés"
    Unload Me
  Else
    Exit Sub
  End If
End Sub
Private Sub SupprTirages_Click()
Dim answer2 As Integer
answer2 = MsgBox("Confirmez-vous l'invalidation des tirages ?", vbYesNo + vbExclamation, "Confirmation Invalidation Tirages")
  If answer2 = vbYes Then
    Sheets("Préparation Tirages C2").Select
    Range("A2:K999").Select
    Selection.EntireRow.Delete
    Range("A1").Select
    ActiveWorkbook.Worksheets("Programme des Courses C2").Columns("BC").ClearContents
    MsgBox "Les tirages ont bien été invalidés.", vbOKOnly + vbInformation, "Tirages Invalidés"
    Call UserForm_Initialize
  Else
    Exit Sub
  End If
End Sub
Private Sub Quit_Click()
Sheets("Programme des Courses C2").Select
            Cells.Select
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
            Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "Lundi,Mardi,Mercredi,Jeudi,Vendredi,Samedi,Dimanche", DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Programme des Courses C2").Sort.SortFields.Add2 Key:= _
            Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Programme des Courses C2").Sort
                .SetRange Range("A1:AW999")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("A1").Select
            Sheets("Réglages Régate").Select
            Sheets("Réglages Régate").Range("G16").Value = ""
            Sheets("Gestion Concept2").Select
 Unload Me
End Sub



