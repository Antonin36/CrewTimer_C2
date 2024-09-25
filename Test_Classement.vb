Private Sub ClassementCalcul_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Voulez-vous procéder au calcul du classement ?", vbYesNo + vbQuestion, "Calcul du Classement")
    If response = vbYes Then
        ' Call the calculation subroutine
        Call CalculClassement_CDF
    Else
        ' Unload the form
        Unload Me
    End If
End Sub

Sub CalculClassement_CDF()
    Dim wsResultats As Worksheet
    Dim wsClassement As Worksheet
    Dim wsReglages As Worksheet
    Dim lastRow As Long
    Dim categorie As String
    Dim courseNum As String
    Dim crewName As String
    Dim crewName2 As Variant
    Dim points As Integer
    Dim nbPartants As Integer
    Dim place As Integer
    Dim finaleType As String
    Dim i As Long
    Dim jour2Start As String ' Variable pour stocker le numéro de la première course du jour 2
    Dim dictJour1 As Object
    Dim key As String
    Dim bonus As Integer

    ' Définir les feuilles
    Set wsResultats = ThisWorkbook.Sheets("Import Resultats CT")
    Set wsClassement = ThisWorkbook.Sheets("Stockage Calcul Classement")
    Set wsRecap = ThisWorkbook.Sheets("Impressions Classement CT")
    Set wsReglages = ThisWorkbook.Sheets("Réglages Régate")
    
    ' Effacer la feuille de classement et de récap
    wsClassement.Cells.Clear
    wsRecap.Range("B13:E113").ClearContents
    
    ' Dernière ligne des résultats
    lastRow = wsResultats.Cells(wsResultats.Rows.Count, 1).End(xlUp).Row
    
    ' Copier les résultats dans la feuille de classement
    wsResultats.Range("A1:H" & lastRow).Copy Destination:=wsClassement.Range("A1")
    
    ' Nombre de partants dans la cellule E14 de la feuille Réglages Régate
    nbPartants = wsReglages.Cells(14, 5).value
    
    ' Numéro de la première course du jour 2 (à ajuster selon ta logique)
    jour2Start = "C03" ' À définir selon le numéro de la première course du jour 2
    
    ' Initialiser un dictionnaire pour stocker les gagnants du jour 1
    Set dictJour1 = CreateObject("Scripting.Dictionary")
    Set dictPoints = CreateObject("Scripting.Dictionary")
    Set dictBonus = CreateObject("Scripting.Dictionary")

    ' Analyser chaque ligne des résultats
    For i = 2 To lastRow ' Ligne 1 contient les en-têtes
        categorie = wsClassement.Cells(i, 6).value ' Colonne F pour "Stroke"
        courseNum = wsClassement.Cells(i, 1).value ' Colonne A pour "EventNum"
        place = wsClassement.Cells(i, 3).value ' Colonne C pour "Place"
        
        ' Traiter le nom de l'équipage
        crewName = wsClassement.Cells(i, 4).value
        crewName = ReformaterCrew(crewName)
        wsClassement.Cells(i, 4).value = crewName
        
        ' Détecter le type de finale (FA, FB, FC, etc.)
        finaleType = Mid(courseNum, InStr(courseNum, "_") + 1, 2)
        
        ' Assigner les points en fonction de la catégorie du bateau
        If InStr(categorie, "4x") > 0 Then
            points = GetPoints4x(DeterminPlaceGlobal(place, finaleType, nbPartants))
        ElseIf InStr(categorie, "8+") > 0 Then
            points = GetPoints8(DeterminPlaceGlobal(place, finaleType, nbPartants))
        Else
            ' Autres catégories si nécessaire
            points = 0
        End If
        
        ' Inscrire les points dans la colonne I
        wsClassement.Cells(i, 9).value = points
        
        ' Gérer les bonus uniquement pour les équipages qui ont remporté la finale A
        If finaleType = "FA" And place = 1 Then
            key = crewName & "|" & categorie ' Utiliser une clé combinant le nom de l'équipage et la catégorie
            
            ' Si c'est le jour 1, stocker l'équipage gagnant
            If courseNum < jour2Start Then
                If Not dictJour1.Exists(key) Then
                    dictJour1.Add key, True
                End If
                bonus = 0
                wsClassement.Cells(i, 10).value = bonus ' Pas de bonus pour le jour 1
            Else
                ' Si c'est le jour 2, vérifier si l'équipage a gagné le jour 1
                If dictJour1.Exists(key) Then
                    bonus = 80 ' Attribuer 80 points bonus
                    wsClassement.Cells(i, 10).value = bonus ' Inscrire les bonus en colonne J
                Else
                    bonus = 0
                    wsClassement.Cells(i, 10).value = bonus ' Pas de bonus si l'équipage n'a pas gagné le jour 1
                End If
            End If
        Else
            bonus = 0
            wsClassement.Cells(i, 10).value = bonus ' Pas de bonus pour les autres finales ou places
        End If
        ' Ajouter les points et bonus au dictionnaire correspondant
        If dictPoints.Exists(crewName) Then
            dictPoints(crewName) = dictPoints(crewName) + points
            dictBonus(crewName) = dictBonus(crewName) + bonus
        Else
            dictPoints.Add crewName, points
            dictBonus.Add crewName, bonus
        End If
    Next i
    
    ' Inscrire le récapitulatif dans la feuille Imp_Classement
    recapRow = 13
    wsRecap.Cells(1, 1).value = "Ligue"
    wsRecap.Cells(1, 2).value = "Points"
    wsRecap.Cells(1, 3).value = "Bonus"
    wsRecap.Cells(1, 4).value = "Total"
    For Each crewName2 In dictPoints.Keys
        wsRecap.Cells(recapRow, 2).value = crewName2
        wsRecap.Cells(recapRow, 3).value = dictPoints(crewName2) ' Total des points
        wsRecap.Cells(recapRow, 4).value = dictBonus(crewName2) ' Total des bonus
        wsRecap.Cells(recapRow, 5).value = dictPoints(crewName2) + dictBonus(crewName2) ' Total points + bonus
        recapRow = recapRow + 1
    Next crewName2
    ThisWorkbook.Sheets("Impressions Classement CT").Select
    Range("B13:E113").Select
    Range("E13").Activate
    ActiveWorkbook.Worksheets("Impressions Classement CT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Impressions Classement CT").Sort.SortFields.Add2 _
        key:=Range("E13:E113"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Impressions Classement CT").Sort
        .SetRange Range("B13:E113")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select

    MsgBox "Le calcul du classement est terminé."
    
    Unload Me
End Sub

' Fonction pour obtenir les points pour la catégorie 4x
Function GetPoints4x(place As Integer) As Integer
    Select Case place
        Case 1: GetPoints4x = 25
        Case 2: GetPoints4x = 22
        Case 3: GetPoints4x = 20
        Case 4: GetPoints4x = 18
        Case 5: GetPoints4x = 16
        Case 6: GetPoints4x = 15
        Case 7: GetPoints4x = 14
        Case 8: GetPoints4x = 13
        Case 9: GetPoints4x = 12
        Case 10: GetPoints4x = 11
        Case 11: GetPoints4x = 10
        Case 12: GetPoints4x = 9
        Case 13: GetPoints4x = 8
        Case 14: GetPoints4x = 7
        Case 15: GetPoints4x = 6
        Case 16: GetPoints4x = 5
        Case 17: GetPoints4x = 4
        Case 18: GetPoints4x = 3
        Case 19: GetPoints4x = 2
        Case Else: GetPoints4x = 1
    End Select
End Function

' Fonction pour obtenir les points pour la catégorie 8+
Function GetPoints8(place As Integer) As Integer
    Select Case place
        Case 1: GetPoints8 = 40
        Case 2: GetPoints8 = 34
        Case 3: GetPoints8 = 30
        Case 4: GetPoints8 = 26
        Case 5: GetPoints8 = 22
        Case 6: GetPoints8 = 20
        Case 7: GetPoints8 = 18
        Case 8: GetPoints8 = 16
        Case 9: GetPoints8 = 14
        Case 10: GetPoints8 = 12
        Case 11: GetPoints8 = 10
        Case 12: GetPoints8 = 9
        Case 13: GetPoints8 = 8
        Case 14: GetPoints8 = 7
        Case 15: GetPoints8 = 6
        Case 16: GetPoints8 = 5
        Case 17: GetPoints8 = 4
        Case 18: GetPoints8 = 3
        Case 19: GetPoints8 = 2
        Case Else: GetPoints8 = 1
    End Select
End Function

' Fonction pour déterminer la place globale en fonction de la finale et du nombre de partants
Function DeterminPlaceGlobal(place As Integer, finaleType As String, nbPartants As Integer) As Integer
    Dim startRange As Integer
    Select Case finaleType
        Case "FA"
            startRange = 1
        Case "FB"
            startRange = 7
        Case "FC"
            startRange = 13
        Case "FD"
            startRange = 19
        Case "FE"
            startRange = 25
        Case "FF"
            startRange = 31
        Case "FG"
            startRange = 37
        Case "FH"
            startRange = 43
        Case "FI"
            startRange = 49
        Case "FJ"
            startRange = 55
        Case "FK"
            startRange = 61
        Case "FL"
            startRange = 67
        ' Ajouter d'autres finales si nécessaire
        Case Else
            startRange = 0
    End Select
    
    ' Calculer la place globale en ajoutant la position dans la finale
    DeterminPlaceGlobal = startRange + place - 1
End Function
Function ReformaterCrew(crew As String) As String
    Dim result As String
    Dim openParenPos As Integer
    
    ' Supprimer tous les chiffres du nom
    result = Replace(crew, "0", "")
    result = Replace(result, "1", "")
    result = Replace(result, "2", "")
    result = Replace(result, "3", "")
    result = Replace(result, "4", "")
    result = Replace(result, "5", "")
    result = Replace(result, "6", "")
    result = Replace(result, "7", "")
    result = Replace(result, "8", "")
    result = Replace(result, "9", "")
    
    ' Supprimer les parenthèses et leur contenu
    openParenPos = InStr(result, "(")
    If openParenPos > 0 Then
        result = Trim(Left(result, openParenPos - 1))
    End If
    
    ReformaterCrew = result
End Function
Private Sub RetourAccueil_Click()
    Unload Me
End Sub