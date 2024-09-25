Sub ClassementWE_Click()

    Dim wsResultats As Worksheet
    Dim wsClassement As Worksheet
    Dim wsReglages As Worksheet
    Dim lastRow As Long
    Dim categorie As String
    Dim courseNum As String
    Dim crewName As String
    Dim points As Integer
    Dim nbPartants As Integer
    Dim place As Integer
    Dim finaleType As String
    Dim i As Long

    ' Définir les feuilles
    Set wsResultats = ThisWorkbook.Sheets("Import Resultats CT")
    Set wsClassement = ThisWorkbook.Sheets("Impressions Classement CT")
    Set wsReglages = ThisWorkbook.Sheets("Réglages Régate")
    
    ' Effacer la feuille de classement
    wsClassement.Cells.Clear
    
    ' Dernière ligne des résultats
    lastRow = wsResultats.Cells(wsResultats.Rows.Count, 1).End(xlUp).Row
    
    ' Copier les résultats dans la feuille de classement
    wsResultats.Range("A1:H" & lastRow).Copy Destination:=wsClassement.Range("A1")
    
    ' Nombre de partants dans la cellule E14 de la feuille Réglages Régate
    nbPartants = wsReglages.Cells(14, 5).value

    ' Analyser chaque ligne des résultats
    For i = 2 To lastRow ' Ligne 1 contient les en-têtes
        categorie = wsClassement.Cells(i, 6).value ' Colonne F pour "Stroke"
        courseNum = wsClassement.Cells(i, 1).value ' Colonne A pour "EventNum"
        place = wsClassement.Cells(i, 3).value ' Colonne C pour "Place"
        
        ' Traiter le nom de l'équipage
        crewName = wsClassement.Cells(i, 4).value
        wsClassement.Cells(i, 4).value = ReformaterCrew(crewName)
        
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
        
        ' Inscrire les points dans la colonne J
        wsClassement.Cells(i, 10).value = points
    Next i

    MsgBox "Le calcul du classement et des points est terminé."

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
