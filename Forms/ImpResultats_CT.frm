VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImpResultats_CT 
   Caption         =   "Impression des Résultats"
   ClientHeight    =   5670
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7956
   OleObjectBlob   =   "ImpResultats_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImpResultats_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Annuler_Click()
    Unload Me
End Sub

Private Sub Imprimer_Click()
            Dim CourseSel As String
            Dim i As Long
            Sheets("Stockage Impressions").Range("1:1").Delete
            For i = 0 To TableauCourses.ListCount - 1
            If TableauCourses.Selected(i) Then
                Sheets("Stockage Impressions").Cells(3, 1 + i).Value = TableauCourses.List(i)
                CourseSel = CourseSel & TableauCourses.List(i) & " / "
            End If
            Next i
            
    Dim course1 As String, course2 As String, course3 As String, course4 As String, course5 As String
    Dim course6 As String, course7 As String, course8 As String, course9 As String, course10 As String
    Dim course11 As String, course12 As String, course13 As String, course14 As String, course15 As String
    Dim course16 As String, course17 As String, course18 As String, course19 As String, course20 As String
    Dim course21 As String, course22 As String, course23 As String, course24 As String, course25 As String
    Dim course26 As String, course27 As String, course28 As String, course29 As String, course30 As String
    Dim course31 As String, course32 As String, course33 As String, course34 As String, course35 As String
    Dim course36 As String, course37 As String, course38 As String, course39 As String, course40 As String
    Dim course41 As String, course42 As String, course43 As String, course44 As String, course45 As String
    Dim course46 As String, course47 As String, course48 As String, course49 As String, course50 As String
    Dim course51 As String, course52 As String, course53 As String, course54 As String, course55 As String
    Dim course56 As String, course57 As String, course58 As String, course59 As String, course60 As String
    Dim course61 As String, course62 As String, course63 As String, course64 As String, course65 As String
    Dim course66 As String, course67 As String, course68 As String, course69 As String, course70 As String
    Dim course71 As String, course72 As String, course73 As String, course74 As String, course75 As String
    Dim course76 As String, course77 As String, course78 As String, course79 As String, course80 As String
    Dim course81 As String, course82 As String, course83 As String, course84 As String, course85 As String
    Dim course86 As String, course87 As String, course88 As String, course89 As String, course90 As String
    Dim course91 As String, course92 As String, course93 As String, course94 As String, course95 As String
    Dim course96 As String, course97 As String, course98 As String, course99 As String, course100 As String
    Dim course101 As String, course102 As String, course103 As String, course104 As String, course105 As String
    Dim course106 As String, course107 As String, course108 As String, course109 As String, course110 As String
    Dim course111 As String, course112 As String, course113 As String, course114 As String, course115 As String
    Dim course116 As String, course117 As String, course118 As String, course119 As String, course120 As String
    Dim course121 As String, course122 As String, course123 As String, course124 As String, course125 As String
    Dim course126 As String, course127 As String, course128 As String, course129 As String, course130 As String
    Dim course131 As String, course132 As String, course133 As String, course134 As String, course135 As String
    Dim course136 As String, course137 As String, course138 As String, course139 As String, course140 As String
    Dim course141 As String, course142 As String, course143 As String, course144 As String, course145 As String
    Dim course146 As String, course147 As String, course148 As String, course149 As String, course150 As String
    Dim course151 As String, course152 As String, course153 As String, course154 As String, course155 As String
    Dim course156 As String, course157 As String, course158 As String, course159 As String, course160 As String
    Dim course161 As String, course162 As String, course163 As String, course164 As String, course165 As String
    Dim course166 As String, course167 As String, course168 As String, course169 As String, course170 As String
    Dim course171 As String, course172 As String, course173 As String, course174 As String, course175 As String
    Dim course176 As String, course177 As String, course178 As String, course179 As String, course180 As String
    Dim course181 As String, course182 As String, course183 As String, course184 As String, course185 As String
    Dim course186 As String, course187 As String, course188 As String, course189 As String, course190 As String
    Dim course191 As String, course192 As String, course193 As String, course194 As String, course195 As String
    Dim course196 As String, course197 As String, course198 As String, course199 As String, course200 As String
    With Sheets("Stockage Impressions")
        course1 = .Range("A3").Value
        course2 = .Range("B3").Value
        course3 = .Range("C3").Value
        course4 = .Range("D3").Value
        course5 = .Range("E3").Value
        course6 = .Range("F3").Value
        course7 = .Range("G3").Value
        course8 = .Range("H3").Value
        course9 = .Range("I3").Value
        course10 = .Range("J3").Value
        course11 = .Range("K3").Value
        course12 = .Range("L3").Value
        course13 = .Range("M3").Value
        course14 = .Range("N3").Value
        course15 = .Range("O3").Value
        course16 = .Range("P3").Value
        course17 = .Range("Q3").Value
        course18 = .Range("R3").Value
        course19 = .Range("S3").Value
        course20 = .Range("T3").Value
        course21 = .Range("U3").Value
        course22 = .Range("V3").Value
        course23 = .Range("W3").Value
        course24 = .Range("X3").Value
        course25 = .Range("Y3").Value
        course26 = .Range("Z3").Value
        course27 = .Range("AA3").Value
        course28 = .Range("AB3").Value
        course29 = .Range("AC3").Value
        course30 = .Range("AD3").Value
        course31 = .Range("AE3").Value
        course32 = .Range("AF3").Value
        course33 = .Range("AG3").Value
        course34 = .Range("AH3").Value
        course35 = .Range("AI3").Value
        course36 = .Range("AJ3").Value
        course37 = .Range("AK3").Value
        course38 = .Range("AL3").Value
        course39 = .Range("AM3").Value
        course40 = .Range("AN3").Value
        course41 = .Range("AO3").Value
        course42 = .Range("AP3").Value
        course43 = .Range("AQ3").Value
        course44 = .Range("AR3").Value
        course45 = .Range("AS3").Value
        course46 = .Range("AT3").Value
        course47 = .Range("AU3").Value
        course48 = .Range("AV3").Value
        course49 = .Range("AW3").Value
        course50 = .Range("AX3").Value
        course51 = .Range("AY3").Value
        course52 = .Range("AZ3").Value
        course53 = .Range("BA3").Value
        course54 = .Range("BB3").Value
        course55 = .Range("BC3").Value
        course56 = .Range("BD3").Value
        course57 = .Range("BE3").Value
        course58 = .Range("BF3").Value
        course59 = .Range("BG3").Value
        course60 = .Range("BH3").Value
        course61 = .Range("BI3").Value
        course62 = .Range("BJ3").Value
        course63 = .Range("BK3").Value
        course64 = .Range("BL3").Value
        course65 = .Range("BM3").Value
        course66 = .Range("BN3").Value
        course67 = .Range("BO3").Value
        course68 = .Range("BP3").Value
        course69 = .Range("BQ3").Value
        course70 = .Range("BR3").Value
        course71 = .Range("BS3").Value
        course72 = .Range("BT3").Value
        course73 = .Range("BU3").Value
        course74 = .Range("BV3").Value
        course75 = .Range("BW3").Value
        course76 = .Range("BX3").Value
        course77 = .Range("BY3").Value
        course78 = .Range("BZ3").Value
        course79 = .Range("CA3").Value
        course80 = .Range("CB3").Value
        course81 = .Range("CC3").Value
        course82 = .Range("CD3").Value
        course83 = .Range("CE3").Value
        course84 = .Range("CF3").Value
        course85 = .Range("CG3").Value
        course86 = .Range("CH3").Value
        course87 = .Range("CI3").Value
        course88 = .Range("CJ3").Value
        course89 = .Range("CK3").Value
        course90 = .Range("CL3").Value
        course91 = .Range("CM3").Value
        course92 = .Range("CN3").Value
        course93 = .Range("CO3").Value
        course94 = .Range("CP3").Value
        course95 = .Range("CQ3").Value
        course96 = .Range("CR3").Value
        course97 = .Range("CS3").Value
        course98 = .Range("CT3").Value
        course99 = .Range("CU3").Value
        course100 = .Range("CV3").Value
        course101 = .Range("CW3").Value
        course102 = .Range("CX3").Value
        course103 = .Range("CY3").Value
        course104 = .Range("CZ3").Value
        course105 = .Range("DA3").Value
        course106 = .Range("DB3").Value
        course107 = .Range("DC3").Value
        course108 = .Range("DD3").Value
        course109 = .Range("DE3").Value
        course110 = .Range("DF3").Value
        course111 = .Range("DG3").Value
        course112 = .Range("DH3").Value
        course113 = .Range("DI3").Value
        course114 = .Range("DJ3").Value
        course115 = .Range("DK3").Value
        course116 = .Range("DL3").Value
        course117 = .Range("DM3").Value
        course118 = .Range("DN3").Value
        course119 = .Range("DO3").Value
        course120 = .Range("DP3").Value
        course121 = .Range("DQ3").Value
        course122 = .Range("DR3").Value
        course123 = .Range("DS3").Value
        course124 = .Range("DT3").Value
        course125 = .Range("DU3").Value
        course126 = .Range("DV3").Value
        course127 = .Range("DW3").Value
        course128 = .Range("DX3").Value
        course129 = .Range("DY3").Value
        course130 = .Range("DZ3").Value
        course131 = .Range("EA3").Value
        course132 = .Range("EB3").Value
        course133 = .Range("EC3").Value
        course134 = .Range("ED3").Value
        course135 = .Range("EE3").Value
        course136 = .Range("EF3").Value
        course137 = .Range("EG3").Value
        course138 = .Range("EH3").Value
        course139 = .Range("EI3").Value
        course140 = .Range("EJ3").Value
        course141 = .Range("EK3").Value
        course142 = .Range("EL3").Value
        course143 = .Range("EM3").Value
        course144 = .Range("EN3").Value
        course145 = .Range("EO3").Value
        course146 = .Range("EP3").Value
        course147 = .Range("EQ3").Value
        course148 = .Range("ER3").Value
        course149 = .Range("ES3").Value
        course150 = .Range("ET3").Value
        course151 = .Range("EU3").Value
        course152 = .Range("EV3").Value
        course153 = .Range("EW3").Value
        course154 = .Range("EX3").Value
        course155 = .Range("EY3").Value
        course156 = .Range("EZ3").Value
        course157 = .Range("FA3").Value
        course158 = .Range("FB3").Value
        course159 = .Range("FC3").Value
        course160 = .Range("FD3").Value
        course161 = .Range("FE3").Value
        course162 = .Range("FF3").Value
        course163 = .Range("FG3").Value
        course164 = .Range("FH3").Value
        course165 = .Range("FI3").Value
        course166 = .Range("FJ3").Value
        course167 = .Range("FK3").Value
        course168 = .Range("FL3").Value
        course169 = .Range("FM3").Value
        course170 = .Range("FN3").Value
        course171 = .Range("FO3").Value
        course172 = .Range("FP3").Value
        course173 = .Range("FQ3").Value
        course174 = .Range("FR3").Value
        course175 = .Range("FS3").Value
        course176 = .Range("FT3").Value
        course177 = .Range("FU3").Value
        course178 = .Range("FV3").Value
        course179 = .Range("FW3").Value
        course180 = .Range("FX3").Value
        course181 = .Range("FY3").Value
        course182 = .Range("FZ3").Value
        course183 = .Range("GA3").Value
        course184 = .Range("GB3").Value
        course185 = .Range("GC3").Value
        course186 = .Range("GD3").Value
        course187 = .Range("GE3").Value
        course188 = .Range("GF3").Value
        course189 = .Range("GG3").Value
        course190 = .Range("GH3").Value
        course191 = .Range("GI3").Value
        course192 = .Range("GJ3").Value
        course193 = .Range("GK3").Value
        course194 = .Range("GL3").Value
        course195 = .Range("GM3").Value
        course196 = .Range("GN3").Value
        course197 = .Range("GO3").Value
        course198 = .Range("GP3").Value
        course199 = .Range("GQ3").Value
        course200 = .Range("GR3").Value
    End With

    With Sheets("Import Resultats")
        .AutoFilterMode = False
        .Range("$A$1:$EA$999").AutoFilter Field:=2, Criteria1:=Array(course1, course2, course3, course4, course5, _
            course6, course7, course8, course9, course10, course11, course12, course13, course14, course15, course16, _
            course17, course18, course19, course20, course21, course22, course23, course24, course25, course26, _
            course27, course28, course29, course30, course31, course32, course33, course34, course35, course36, _
            course37, course38, course39, course40, course41, course42, course43, course44, course45, course46, _
            course47, course48, course49, course50, course51, course52, course53, course54, course55, course56, _
            course57, course58, course59, course60, course61, course62, course63, course64, course65, course66, _
            course67, course68, course69, course70, course71, course72, course73, course74, course75, course76, _
            course77, course78, course79, course80, course81, course82, course83, course84, course85, course86, _
            course87, course88, course89, course90, course91, course92, course93, course94, course95, course96, _
            course97, course98, course99, course100, course101, course102, course103, course104, course105, _
            course106, course107, course108, course109, course110, course111, course112, course113, course114, _
            course115, course116, course117, course118, course119, course120, course121, course122, course123, _
            course124, course125, course126, course127, course128, course129, course130, course131, course132, _
            course133, course134, course135, course136, course137, course138, course139, course140, course141, _
            course142, course143, course144, course145, course146, course147, course148, course149, course150, _
            course151, course152, course153, course154, course155, course156, course157, course158, course159, _
            course160, course161, course162, course163, course164, course165, course166, course167, course168, _
            course169, course170, course171, course172, course173, course174, course175, course176, course177, _
            course178, course179, course180, course181, course182, course183, course184, course185, course186, _
            course187, course188, course189, course190, course191, course192, course193, course194, course195, _
            course196, course197, course198, course199, course200), _
            Operator:=xlFilterValues
    End With
    'Sheets("Impressions CT").Select
    
    Sheets("Import Resultats").Select
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("A2:H999").Select
    Selection.Copy
    Sheets("Impressions Résultats CT").Select
    Range("A13:H999").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A13").Select
    Sheets("Import Resultats").Select
    Range("A2:H999").Select
    Selection.Copy
    Sheets("Impressions Résultats CT").Select
    Range("A13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Unload Me
            
End Sub

Private Sub UserForm_Initialize()
TableauCourses.Clear
TableauCourses.ListIndex = -1
Dim UniqueList()    As String
    Dim x               As Long
    Dim Rng1            As Range
    Dim C               As Range
    Dim Unique          As Boolean
    Dim y               As Long
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
     
    Set Rng1 = Sheets("Import Resultats").Range("B2:B999")
    y = 1
     
    ReDim UniqueList(1 To Rng1.Rows.Count)
     
    For Each C In Rng1
        If Not C.Value = vbNullString Then
            Unique = True
            For x = 1 To y
                If UniqueList(x) = C.Text Then
                    Unique = False
                End If
            Next
            If Unique Then
                y = y + 1
                Me.TableauCourses.AddItem (C.Text)
                UniqueList(y) = C.Text
            End If
        End If
    Next
    
    With TableauCourses
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
End Sub


