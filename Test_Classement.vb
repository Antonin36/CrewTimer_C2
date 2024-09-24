Private Sub btnClassementJour1_Click()
    ' Gestion du classement pour le jour 1
    If MsgBox("Voulez-vous calculer et afficher le classement du premier jour ?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        Dim courseNumJour2 As Integer
        courseNumJour2 = InputBox("Entrez le numéro de la première course du jour 2", "Numéro de Course Jour 2")
        
        ' Calcul du classement du jour 1
        CalculerClassementJour1 courseNumJour2
        
        MsgBox "Classement du premier jour terminé et affiché.", vbInformation, "Terminé"
    End If
End Sub

Private Sub btnClassementFinal_Click()
    ' Gestion du classement des deux jours
    If MsgBox("Voulez-vous calculer et afficher le classement des deux jours ?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        Dim courseNumJour2 As Integer
        courseNumJour2 = InputBox("Entrez le numéro de la première course du jour 2", "Numéro de Course Jour 2")
        
        ' Calcul du classement final
        GestionClassement courseNumJour2
        
        MsgBox "Classement des deux jours terminé et affiché.", vbInformation, "Terminé"
    End If
End Sub

' Sub pour gérer le classement total des deux jours avec tri et affichage
Sub GestionClassement(courseNumJour2 As Integer)
    Dim wsResultats As Worksheet
    Dim wsClassement As Worksheet
    Dim wsJour1 As Worksheet
    Dim i As Integer, place As Integer
    Dim crew As String, ligue As String
    Dim points As Integer
    Dim typeFinale As String
    Dim category As String
    Dim maxRow As Long
    Dim totalBonus As Integer
    Dim bonusCategories As String
    
    ' Feuilles contenant les résultats et la feuille de sortie pour le classement
    Set wsResultats = Worksheets("FeuilleDeResultats")
    Set wsClassement = Worksheets("Classement")
    Set wsJour1 = Worksheets("ClassementJour1")
    
    ' Initialisation de la feuille de classement
    wsClassement.Cells.Clear
    wsClassement.Range("A1:F1").Value = Array("Place", "Nom de la Ligue", "Catégorie", "Type Finale", "Points", "Bonus et Catégories")
    
    ' Trouver la dernière ligne de la feuille de résultats
    maxRow = wsResultats.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Variable pour bonus et catégories
    totalBonus = 0
    bonusCategories = ""
    
    ' Parcourir les résultats
    For i = 2 To maxRow
        ' Récupérer la place
        place = wsResultats.Cells(i, 3).Value
        
        ' Récupérer le nom de l'équipage
        crew = wsResultats.Cells(i, 4).Value
        
        ' Extraire le nom de la ligue (avant le numéro d'équipage)
        ligue = ExtraireLigue(crew)
        
        ' Récupérer le type de finale (FA, FB, etc.) à partir de l'événement
        typeFinale = ExtraireTypeFinale(wsResultats.Cells(i, 2).Value, place)
        
        ' Récupérer la catégorie (par exemple J16H4x, SH8+)
        category = wsResultats.Cells(i, 9).Value
        
        ' Calcul des points en fonction de la place et de la catégorie
        points = CalculerPoints(place, category)
        
        ' Calcul des bonus pour les victoires en finale A sur les deux jours
        If place = 1 And IsFinaleA(typeFinale) Then
            totalBonus = totalBonus + 80
            If bonusCategories <> "" Then bonusCategories = bonusCategories & ", "
            bonusCategories = bonusCategories & category
        End If
        
        ' Remplir les résultats dans la feuille de classement
        wsClassement.Cells(i, 1).Value = place
        wsClassement.Cells(i, 2).Value = ligue
        wsClassement.Cells(i, 3).Value = category
        wsClassement.Cells(i, 4).Value = typeFinale
        wsClassement.Cells(i, 5).Value = points
        wsClassement.Cells(i, 6).Value = totalBonus & " (" & bonusCategories & ")"
    Next i
    
    ' Trier le tableau par la colonne des points (colonne E, soit la 5ème colonne)
    wsClassement.Range("A2:F" & maxRow).Sort Key1:=wsClassement.Range("E2"), Order1:=xlDescending, Header:=xlNo
    
    ' Calculer le classement à la fin du premier jour
    CalculerClassementJour1 wsResultats, wsJour1, courseNumJour2
    
    ' Sélectionner et afficher le tableau trié
    wsClassement.Select
    wsClassement.Range("A1:F" & maxRow).Select
    MsgBox "Le classement trié a été affiché dans la feuille 'Classement'.", vbInformation, "Classement Terminé"
End Sub

' Sub pour calculer le classement du jour 1 avec tri et affichage
Sub CalculerClassementJour1(courseNumJour2 As Integer)
    Dim wsResultats As Worksheet
    Dim wsJour1 As Worksheet
    Dim i As Integer
    Dim ligue As String
    Dim points As Integer
    Dim maxRow As Long
    
    ' Feuilles contenant les résultats et la feuille de sortie pour le jour 1
    Set wsResultats = Worksheets("FeuilleDeResultats")
    Set wsJour1 = Worksheets("ClassementJour1")
    
    ' Initialisation de la feuille pour le jour 1
    wsJour1.Cells.Clear
    wsJour1.Range("A1:C1").Value = Array("Ligue", "Points", "Total Points")
    
    ' Parcourir les résultats jusqu'au numéro de course du jour 2
    For i = 2 To wsResultats.Cells(Rows.Count, 1).End(xlUp).Row
        If wsResultats.Cells(i, 1).Value >= courseNumJour2 Then Exit For
        
        ' Extraire les informations nécessaires
        ligue = ExtraireLigue(wsResultats.Cells(i, 4).Value)
        points = CalculerPoints(wsResultats.Cells(i, 3).Value, wsResultats.Cells(i, 9).Value)
        
        ' Remplir les informations dans la feuille du jour 1
        wsJour1.Cells(i, 1).Value = ligue
        wsJour1.Cells(i, 2).Value = points
        wsJour1.Cells(i, 3).Formula = "=SUMIF(A:A,A" & i & ",B:B)"
    Next i
    
    ' Définir la dernière ligne à trier
    maxRow = wsJour1.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Trier le tableau par la colonne des points (colonne B)
    wsJour1.Range("A2:C" & maxRow).Sort Key1:=wsJour1.Range("B2"), Order1:=xlDescending, Header:=xlNo
    
    ' Sélectionner et afficher le tableau trié
    wsJour1.Select
    wsJour1.Range("A1:C" & maxRow).Select
    MsgBox "Le classement du premier jour trié a été affiché dans la feuille 'ClassementJour1'.", vbInformation, "Classement Jour 1 Terminé"
End Sub

' Fonction pour calculer les points en fonction de la place et de la catégorie
Function CalculerPoints(place As Integer, category As String) As Integer
    Dim points4x() As Integer
    Dim points8plus() As Integer
    
    ' Points pour 4x
    points4x = Array(25, 22, 20, 18, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    
    ' Points pour 8+
    points8plus = Array(40, 34, 30, 26, 22, 20, 18, 16, 14, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    
    ' Calculer les points en fonction de la catégorie et de la place
    If InStr(category, "4x") > 0 Then
        CalculerPoints = IIf(place <= 20, points4x(place - 1), 1)
    ElseIf InStr(category, "8+") > 0 Then
        CalculerPoints = IIf(place <= 20, points8plus(place - 1), 1)
    Else
        CalculerPoints = 0
    End If
End Function

' Fonction pour extraire la ligue à partir du nom de l'équipage
Function ExtraireLigue(crew As String) As String
    ExtraireLigue = Trim(Split(crew, "(")(0))
End Function

' Fonction pour extraire le type de finale (FA, FB, FC, etc.)
Function ExtraireTypeFinale(eventName As String, place As Integer) As String
    If place <= 6 Then
        ExtraireTypeFinale = "FA"
    ElseIf place <= 12 Then
        ExtraireTypeFinale = "FB"
    ElseIf place <= 18 Then
        ExtraireTypeFinale = "FC"
    ElseIf place <= 24 Then
        ExtraireTypeFinale = "FD"
    Else
        ExtraireTypeFinale = "Autre"
    End If
End Function

' Fonction pour vérifier si la finale est une finale A
Function IsFinaleA(typeFinale As String) As Boolean
    IsFinaleA = (typeFinale = "FA")
End Function
