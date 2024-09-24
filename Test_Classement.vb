Private Sub ClassementJ1_Click()
    If MsgBox("Êtes-vous sûr de vouloir calculer le classement du Jour 1 ?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        CalculerClassement "J1"
    End If
End Sub

Private Sub ClassementWE_Click()
    If MsgBox("Êtes-vous sûr de vouloir calculer le classement du Week-End ?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        CalculerClassement "WE"
    End If
End Sub

Private Sub Retour_Accueil_Click()
    Unload Me
End Sub

Sub CalculerClassement(jour As String)
    Dim wsResultats As Worksheet
    Dim wsClassement As Worksheet
    Dim nbPartants As Integer
    Dim bonusGagnes As Dictionary
    Dim categorieKey As String
    
    ' Récupérer la feuille de résultats et la feuille de classement
    Set wsResultats = Worksheets("Import Resultats CT")
    Set wsClassement = Worksheets("Impressions Classement CT")
    
    ' Demander ou récupérer le nombre de partants
    nbPartants = DemanderNombrePartants()
    
    ' Initialiser le dictionnaire pour les bonus gagnés
    Set bonusGagnes = New Dictionary
    
    ' Boucler à travers les résultats pour calculer le classement
    Dim i As Integer
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, 1).End(xlUp).Row
        Dim place As Integer
        Dim crew As String
        Dim eventCode As String
        Dim category As String
        Dim ligue As String
        Dim points As Integer
        Dim typeFinale As String
        
        place = wsResultats.Cells(i, 3).Value
        crew = wsResultats.Cells(i, 4).Value
        eventCode = wsResultats.Cells(i, 2).Value
        category = ExtraireCategorie(wsResultats.Cells(i, 2).Value)
        ligue = ExtraireLigue(crew)
        
        ' Calculer les points en fonction de la catégorie et de la place
        points = CalculerPoints(category, place)
        
        ' Déterminer le type de finale (FA, FB, FC, FD)
        typeFinale = ExtraireTypeFinale(eventCode, place, nbPartants)
        
        ' Enregistrer les résultats dans la feuille de classement
        With wsClassement
            .Cells(i, 1).Value = ligue
            .Cells(i, 2).Value = crew
            .Cells(i, 3).Value = place
            .Cells(i, 4).Value = points
            .Cells(i, 5).Value = typeFinale
        End With
        
        ' Gestion des bonus : vérifier si l'équipage a gagné dans la même catégorie d'âge et bateau
        If jour = "WE" And IsFinaleA(typeFinale) And place = 1 Then
            categorieKey = category & "_" & crew
            If bonusGagnes.Exists(categorieKey) Then
                ' Vérifier si cet équipage a déjà gagné la finale A le premier jour
                If bonusGagnes(categorieKey) = "J1" Then
                    ' Ajouter les 80 points de bonus
                    With wsClassement
                        .Cells(i, 6).Value = 80
                        .Cells(i, 7).Value = category
                    End With
                End If
            Else
                ' Enregistrer que l'équipage a gagné la finale A au jour 1
                bonusGagnes.Add categorieKey, jour
            End If
        End If
    Next i
    
    ' Trier le classement par points décroissants
    wsClassement.Sort.SortFields.Clear
    wsClassement.Sort.SortFields.Add Key:=wsClassement.Columns(4), Order:=xlDescending
    wsClassement.Sort.SetRange wsClassement.Range("A1:E" & wsClassement.Cells(wsClassement.Rows.Count, 1).End(xlUp).Row)
    wsClassement.Sort.Header = xlYes
    wsClassement.Sort.Apply
    
    ' Afficher le classement
    MsgBox "Classement calculé et trié !", vbInformation
End Sub

' Fonction pour calculer les points selon la place et la catégorie
Function CalculerPoints(category As String, place As Integer) As Integer
    Dim points4x As Variant
    Dim points8plus As Variant
    
    points4x = Array(25, 22, 20, 18, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    points8plus = Array(40, 34, 30, 26, 22, 20, 18, 16, 14, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    
    If InStr(category, "4x") > 0 Then
        ' 4x
        If place <= 20 Then
            CalculerPoints = points4x(place - 1)
        Else
            CalculerPoints = 1
        End If
    ElseIf InStr(category, "8+") > 0 Then
        ' 8+
        If place <= 20 Then
            CalculerPoints = points8plus(place - 1)
        Else
            CalculerPoints = 1
        End If
    Else
        ' Autres types de courses peuvent être ajoutés à l'avenir
        CalculerPoints = 0
    End If
End Function

' Fonction pour demander le nombre de partants à l'utilisateur ou le récupérer depuis la feuille de réglages
Function DemanderNombrePartants() As Integer
    If MsgBox("Voulez-vous utiliser le nombre de partants stocké ?", vbYesNo + vbQuestion, "Utiliser le nombre de partants") = vbYes Then
        ' Récupérer depuis la feuille de réglages
        DemanderNombrePartants = Worksheets("Réglages Régate").Range("E14").Value
    Else
        ' Demander à l'utilisateur
        DemanderNombrePartants = InputBox("Entrez le nombre de partants", "Nombre de Partants")
    End If
End Function

' Fonction pour extraire le nom de la ligue à partir du nom de l'équipage
Function ExtraireLigue(crew As String) As String
    Dim parts() As String
    parts = Split(crew, " ")
    ExtraireLigue = parts(0) ' Retourne la première partie (nom de la ligue)
End Function

' Fonction pour extraire le type de finale (FA, FB, etc.)
Function ExtraireTypeFinale(eventCode As String, place As Integer, nbPartants As Integer) As String
    If place <= 6 Then
        ExtraireTypeFinale = "FA"
    ElseIf place <= 12 Then
        ExtraireTypeFinale = "FB"
    ElseIf place <= 18 Then
        ExtraireTypeFinale = "FC"
    ElseIf place <= nbPartants Then
        ExtraireTypeFinale = "FD"
    Else
        ExtraireTypeFinale = "Autres"
    End If
End Function

' Fonction pour déterminer si une course est en finale A
Function IsFinaleA(typeFinale As String) As Boolean
    IsFinaleA = (typeFinale = "FA")
End Function
