VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifCourse_CT 
   Caption         =   "Modification d'une Course"
   ClientHeight    =   6040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7980
   OleObjectBlob   =   "ModifCourse_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifCourse_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CourseARemplacer_CT As Long
Private Sub Sauvegarder_Click()
Dim CategSel As String
            Dim i As Long
            For i = 0 To Categ.ListCount - 1
            If Categ.Selected(i) Then
                Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 10 + i).value = Categ.List(i)
                CategSel = CategSel & Categ.List(i) & " / "
            End If
            Next i
            CategSel = Left(CategSel, Len(CategSel) - 3)
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "A").value = Jour.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "B").value = Heure.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "C").value = IDCourse.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "D").value = TypeCourse.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "E").value = TypeCourse.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "F").value = CategSel
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "G").value = Jour.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "H").value = Tirage.Text
            Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, "I").value = InfoSysProg.Text
            Sheets("Programme des Courses CT").Select
            Columns("G:G").Select
            Selection.Replace What:="Lundi", Replacement:="Monday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Mardi", Replacement:="Tuesday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Mercredi", Replacement:="Wednesday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Jeudi", Replacement:="Thursday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Vendredi", Replacement:="Friday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Samedi", Replacement:="Saturday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Dimanche", Replacement:="Sunday", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Columns("D:D").Select
            Selection.Replace What:="Série 1", Replacement:="H1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 2", Replacement:="H2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 3", Replacement:="H3", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 4", Replacement:="H4", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 5", Replacement:="H5", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 6", Replacement:="H6", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 7", Replacement:="H7", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Série 8", Replacement:="H8", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale A-D 1", Replacement:="QAD1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale A-D 2", Replacement:="QAD2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale A-D 3", Replacement:="QAD3", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale A-D 4", Replacement:="QAD4", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale E-H 1", Replacement:="QEH1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale E-H 2", Replacement:="QEH2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale E-H 3", Replacement:="QEH3", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Quart de Finale E-H 4", Replacement:="QEH4", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale A-B 1", Replacement:="SAB1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale A-B 2", Replacement:="SAB2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale C-D 1", Replacement:="SCD1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale C-D 2", Replacement:="SCD2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale E-F 1", Replacement:="SEF1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale E-F 2", Replacement:="SEF2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale G-H 1", Replacement:="SGH1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Demi-Finale G-H 2", Replacement:="SGH2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale A", Replacement:="FA", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale B", Replacement:="FB", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale C", Replacement:="FC", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale D", Replacement:="FD", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale E", Replacement:="FE", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale F", Replacement:="FF", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale G", Replacement:="FG", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale H", Replacement:="FH", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série Unique", Replacement:="TT", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 1", Replacement:="TT1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 2", Replacement:="TT2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 3", Replacement:="TT3", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 4", Replacement:="TT4", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 5", Replacement:="TT5", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 6", Replacement:="TT6", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 7", Replacement:="TT7", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre Série 8", Replacement:="TT8", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale A Directe (Pas de Série)", Replacement:="Final", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Autre", Replacement:="Unspecified", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
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
            Sheets("Gestion CrewTimer").Select
          MsgBox "La course à été modifiée avec succès !", vbOKOnly + vbInformation, "Course Modifiée"
      Unload Me
End Sub
Private Sub UserForm_Initialize()
    ' Cette méthode est appelée pour initialiser l'UserForm avec la ligne spécifique à modifier
    CourseARemplacer_CT = Sheets("Réglages Régate").Cells(26, "B").value
    IDCourse.Text = "C00"
Heure.Text = "00:00"
Categ.Clear
Categ.ListIndex = -1
Dim UniqueList()    As String
    Dim x               As Long
    Dim Rng1            As Range
    Dim C               As Range
    Dim Unique          As Boolean
    Dim y               As Long
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
     
    Set Rng1 = Sheets("Import Goal CT").Range("C2:C999")
    y = 1
     
    ReDim UniqueList(1 To Rng1.Rows.Count)
     
    For Each C In Rng1
        If Not C.value = vbNullString Then
            Unique = True
            For x = 1 To y
                If UniqueList(x) = C.Text Then
                    Unique = False
                End If
            Next
            If Unique Then
                y = y + 1
                Me.Categ.AddItem (C.Text)
                UniqueList(y) = C.Text
            End If
        End If
    Next
    
    With Categ
        For i = 0 To .ListCount - 2
            For j = i + 1 To .ListCount - 1
                If .List(i) > .List(j) Then
                    Temp = .List(j)
                    .List(j) = .List(i)
                    .List(i) = Temp
                End If
            Next j
        Next i
    End With
    
    Dim UniqueList2()    As String
    Dim A               As Long
    Dim Rng2            As Range
    Dim D               As Range
    Dim Unique2          As Boolean
    Dim w               As Long
     
    Set Rng2 = Sheets("Référentiel Progression CT").Range("B2:B999")
    w = 1
     
    ReDim UniqueList2(1 To Rng2.Rows.Count)
     
    For Each D In Rng2
        If Not D.value = vbNullString Then
            Unique2 = True
            For A = 1 To w
                If UniqueList2(A) = D.Text Then
                    Unique2 = False
                End If
            Next
            If Unique2 Then
                w = w + 1
                Me.TypeCourse.AddItem (D.Text)
                UniqueList2(w) = D.Text
            End If
        End If
    Next
    Me.Tirage.AddItem ("Oui")
    Me.Tirage.AddItem ("Non")
    Me.Jour.AddItem ("Lundi")
    Me.Jour.AddItem ("Mardi")
    Me.Jour.AddItem ("Mercredi")
    Me.Jour.AddItem ("Jeudi")
    Me.Jour.AddItem ("Vendredi")
    Me.Jour.AddItem ("Samedi")
    Me.Jour.AddItem ("Dimanche")
    ' Ajoutez ici le code nécessaire pour charger les données de la ligne spécifiée dans les contrôles de l'UserForm
    ' Par exemple, si vous avez un contrôle TextBoxNomCourse, vous pouvez le remplir comme ceci :
    Jour.Text = Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 1).value
    Heure.Text = Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 2).Text
    IDCourse.Text = Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 3).value
    TypeCourse.Text = Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 5).value
    Tirage.Text = Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 8).value
    InfoSysProg.Text = Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, 9).value
    'Ne Pas Récupérer G et D
    ' Ajoutez d'autres lignes similaires pour les autres contrôles que vous souhaitez initialiser
    For colcateg = 10 To 49
    ' Vérifiez si la cellule spécifiée contient une valeur
    If Not IsEmpty(Sheets("Programme des Courses CT").Cells(CourseARemplacer_CT, colcateg).value) Then
        ' Ajoutez l'option à la ListBox
        Categ.Selected(colcateg - 10) = True ' Vous pouvez utiliser la valeur de la cellule ici si nécessaire
    End If
    Next colcateg
End Sub

Private Sub Annuler_Click()
    CourseModif_CT = 0
    CourseARemplacer_CT = 0
    Unload Me
End Sub
