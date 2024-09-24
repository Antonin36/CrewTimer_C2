Sub GestionClassement()
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
    Dim courseNumJour2 As Integer ' Variable pour définir le numéro de course à partir duquel commence le jour 2
    Dim courseNum As Integer ' Numéro de la course extrait
    
    ' Initialisation du numéro de course à partir duquel on commence le jour 2
    courseNumJour2 = 50 ' Vous pouvez modifier cette valeur selon vos besoins
    
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
    
    ' Calculer le classement à la fin du premier jour
    CalculerClassementJour1 wsResultats, wsJour1, courseNumJour2
    
    MsgBox "Classement mis à jour avec succès !"
End Sub

' Fonction pour extraire le nom de la ligue depuis le nom de l'équipage
Function ExtraireLigue(crew As String) As String
    Dim ligue As String
    Dim pos As Integer
    
    ' Trouver la position du dernier espace avant le numéro d'équipage
    pos = InStrRev(crew, " ")
    
    ' Extraire la ligue (tout avant le dernier espace)
    ligue = Left(crew, pos - 1)
    
    ' Retourner le nom de la ligue
    ExtraireLigue = ligue
End Function

' Fonction pour extraire le type de finale (FA, FB, etc.) depuis le nom de l'événement et la place
Function ExtraireTypeFinale(eventName As String, place As Integer) As String
    Dim typeFinale As String
    
    ' Déterminer le type de finale en fonction de la place
    Select Case place
        Case 1 To 6
            typeFinale = "FA"
        Case 7 To 12
            typeFinale = "FB"
        Case 13 To 18
            typeFinale = "FC"
        Case 19 To 24
            typeFinale = "FD"
        Case Else
            typeFinale = "Autre"
    End Select
    
    ' Retourner le type de finale
    ExtraireTypeFinale = typeFinale
End Function

' Fonction pour déterminer si c'est la finale A
Function IsFinaleA(typeFinale As String) As Boolean
    If typeFinale = "FA" Then
        IsFinaleA = True
    Else
        IsFinaleA = False
    End If
End Function

' Fonction pour calculer les points en fonction de la place et de la catégorie
Function CalculerPoints(place As Integer, category As String) As Integer
    Dim points As Integer
    
    ' Points pour la catégorie 4x
    If Right(category, 2) = "4x" Then
        Select Case place
            Case 1: points = 25
            Case 2: points = 22
            Case 3: points = 20
            Case 4: points = 18
            Case 5: points = 16
            Case 6: points = 15
            Case 7: points = 14
            Case 8: points = 13
            Case 9: points = 12
            Case 10: points = 11
            Case 11: points = 10
            Case 12: points = 9
            Case 13: points = 8
            Case 14: points = 7
            Case 15: points = 6
            Case 16: points = 5
            Case 17: points = 4
            Case 18: points = 3
            Case 19: points = 2
            Case 20: points = 1
            Case Else: points = 1
        End Select
        
    ' Points pour la catégorie 8+
    ElseIf Right(category, 2) = "8+" Then
        Select Case place
            Case 1: points = 40
            Case 2: points = 34
            Case 3: points = 30
            Case 4: points = 26
            Case 5: points = 22
            Case 6: points = 20
            Case 7: points = 18
            Case 8: points = 16
            Case 9: points = 14
            Case 10: points = 12
            Case 11: points = 10
            Case 12: points = 9
            Case 13: points = 8
            Case 14: points = 7
            Case 15: points = 6
            Case 16: points = 5
            Case 17: points = 4
            Case 18: points = 3
            Case 19: points = 2
            Case 20: points = 1
            Case Else: points = 1
        End Select
    End If
    
    ' Retourner les points calculés
    CalculerPoints = points
End Function

' Fonction pour calculer le classement à la fin du premier jour
Sub CalculerClassementJour1(wsResultats As Worksheet, wsJour1 As Worksheet, courseNumJour2 As Integer)
    Dim i As Integer
    Dim ligue As String
    Dim points As Integer
    Dim maxRow As Long
    Dim courseNum As Integer
    
    ' Initialisation de la feuille de classement pour le jour 1
    wsJour1.Cells.Clear
    wsJour1.Range("A1:C1").Value = Array("Nom de la Ligue", "Catégorie", "Total Points")
    
    ' Trouver la dernière ligne de la feuille de résultats
    maxRow = wsResultats.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Parcourir les résultats pour le jour 1 uniquement
    For i = 2 To maxRow
        ' Extraire le numéro de la course à partir du format CXX
        courseNum = CInt(Mid(wsResultats.Cells(i, 1).Value, 2))
        
        ' Si la course fait partie du jour 1 (c'est-à-dire que le numéro de course est inférieur au seuil du jour 2)
        If courseNum < courseNumJour2 Then
            ligue = ExtraireLigue(wsResultats.Cells(i, 4).Value)
            points = wsResultats.Cells(i, 5).Value ' Récupérer les points
            
            ' Ajouter les données dans la feuille de classement du jour 1
            wsJour1.Cells(i, 1).Value = ligue
            wsJour1.Cells(i, 2).Value = wsResultats.Cells(i, 9).Value ' Catégorie
            wsJour1.Cells(i, 3).Value = points
        End If
    Next i
End Sub
