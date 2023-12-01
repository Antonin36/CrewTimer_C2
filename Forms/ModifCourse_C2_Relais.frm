VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifCourse_C2_Relais 
   Caption         =   "Modification d'une Course en Relais"
   ClientHeight    =   7640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780.001
   OleObjectBlob   =   "ModifCourse_C2_Relais.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifCourse_C2_Relais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CourseARemplacer_C2 As Long
Private Sub Annuler_Click()
    CourseModif_CT = 0
    CourseARemplacer_CT = 0
    Unload Me
End Sub

Private Sub Sauvegarder_Click()
            Dim CategSel As String
            Dim i As Long
            ' Trouver la premi�re cellule vide entre les colonnes 10 et 49
            Dim emptyCell As Range
            Set emptyCell = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 10)

            For i = 0 To Categ.ListCount - 1
            If Categ.Selected(i) Then
            ' Placer la valeur dans la premi�re cellule vide entre les colonnes 10 et 49
            emptyCell.value = Categ.List(i)
            CategSel = CategSel & Categ.List(i) & " / "
            ' Passer � la cellule suivante
            Set emptyCell = emptyCell.Offset(0, 1)
            End If
            Next i
            CategSel = Left(CategSel, Len(CategSel) - 3)
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "A").value = Jour.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "B").value = Heure.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "C").value = IDCourse.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "D").value = EtapeCourse.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "E").value = EtapeCourse.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "F").value = CategSel
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "G").value = Jour.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "H").value = Tirage.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "I").value = InfoSysProg.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "BA").value = TypeCourse.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "AX").value = DureeCourse.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "AY").value = Split.Text
            Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, "AZ").value = "Relais"
            Sheets("Programme des Courses C2").Select
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
            Selection.Replace What:="S�rie 1", Replacement:="H1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 2", Replacement:="H2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 3", Replacement:="H3", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 4", Replacement:="H4", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 5", Replacement:="H5", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 6", Replacement:="H6", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 7", Replacement:="H7", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="S�rie 8", Replacement:="H8", LookAt:=xlWhole, _
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
            Selection.Replace What:="Contre-la-Montre S�rie Unique", Replacement:="TT", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 1", Replacement:="TT1", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 2", Replacement:="TT2", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 3", Replacement:="TT3", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 4", Replacement:="TT4", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 5", Replacement:="TT5", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 6", Replacement:="TT6", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 7", Replacement:="TT7", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Contre-la-Montre S�rie 8", Replacement:="TT8", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Finale A Directe (Pas de S�rie)", Replacement:="Final", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.Replace What:="Autre", Replacement:="Unspecified", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
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
            Sheets("Gestion Concept2").Select
          MsgBox "La course � �t� modifi�e avec succ�s !", vbOKOnly + vbInformation, "Course Modifi�e"
      Unload Me
End Sub

Private Sub UserForm_Initialize()
' Cette m�thode est appel�e pour initialiser l'UserForm avec la ligne sp�cifique � modifier
CourseARemplacer_C2 = Sheets("R�glages R�gate").Cells(27, "B").value
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
     
    Set Rng1 = Sheets("Import GOAL C2").Range("C2:C999")
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
     
    Set Rng2 = Sheets("R�f�rentiel Progression C2").Range("B2:B999")
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
                Me.EtapeCourse.AddItem (D.Text)
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
    Me.TypeCourse.AddItem ("Distance")
    Me.TypeCourse.AddItem ("Temps")
    Me.TypeCourse.AddItem ("Calories")
    ' Ajoutez ici le code n�cessaire pour charger les donn�es de la ligne sp�cifi�e dans les contr�les de l'UserForm
    ' Par exemple, si vous avez un contr�le TextBoxNomCourse, vous pouvez le remplir comme ceci :
    Jour.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 1).value
    Heure.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 2).Text
    IDCourse.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 3).value
    EtapeCourse.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 5).value
    Tirage.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 8).value
    InfoSysProg.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 9).value
    TypeCourse.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 53).value
    DureeCourse.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 50).value
    Split.Text = Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, 51).value
    'Ne Pas R�cup�rer G et D
    ' Ajoutez d'autres lignes similaires pour les autres contr�les que vous souhaitez initialiser
    For colcateg = 10 To 49
    ' V�rifiez si la cellule sp�cifi�e contient une valeur
    If Not IsEmpty(Sheets("Programme des Courses C2").Cells(CourseARemplacer_C2, colcateg).value) Then
        ' Ajoutez l'option � la ListBox
        Categ.Selected(colcateg - 10) = True ' Vous pouvez utiliser la valeur de la cellule ici si n�cessaire
    End If
    Next colcateg
End Sub




