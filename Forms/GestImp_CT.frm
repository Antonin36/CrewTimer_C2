VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestImp_CT 
   Caption         =   "Gestion des Impressions"
   ClientHeight    =   5140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7780
   OleObjectBlob   =   "GestImp_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestImp_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImpEmargement_Click()
    ImpEmargement_CT.Show
End Sub

Private Sub ImpInscrits_Click()
    ImpInscrits_CT.Show
    If Sheets("Réglages Régate").Range("K30").value = "Ferm" Then
        Sheets("Réglages Régate").Range("K30").value = ""
        Unload Me
    End If
End Sub

Private Sub ImpPesee_Click()
    ImpPesee_CT.Show
    If Sheets("Réglages Régate").Range("K30").value = "Ferm" Then
        Sheets("Réglages Régate").Range("K30").value = ""
        Unload Me
    End If
End Sub

Private Sub ImpResultatsCateg_Click()
    ImpResultatsCateg_CT.Show
    If Sheets("Réglages Régate").Range("K30").value = "Ferm" Then
        Sheets("Réglages Régate").Range("K30").value = ""
        Unload Me
    End If
End Sub

Private Sub ImpResultatsCourse_Click()
    ImpResultatsCourse_CT.Show
    If Sheets("Réglages Régate").Range("K30").value = "Ferm" Then
        Sheets("Réglages Régate").Range("K30").value = ""
        Unload Me
    End If
End Sub

Private Sub ImpTiragesCateg_Click()
    ImpResultatsCateg_CT.Show
    If Sheets("Réglages Régate").Range("K30").value = "Ferm" Then
        Sheets("Réglages Régate").Range("K30").value = ""
        Unload Me
    End If
End Sub

Private Sub ImpTiragesCourse_Click()
    ImpTiragesCourse_CT.Show
    If Sheets("Réglages Régate").Range("K30").value = "Ferm" Then
        Sheets("Réglages Régate").Range("K30").value = ""
        Unload Me
    End If
End Sub

Private Sub RetourAccueil_Click()
    Unload Me
End Sub

Sub UserForm_Initialize()
ImpPesee.Caption = "Impression des" & vbCrLf & "feuilles de Pesée"
ImpEmargement.Caption = "Impression des feuilles" & vbCrLf & "pour l'émargement"
ImpTiragesCateg.Caption = "Impression des Tirages" & vbCrLf & "par Catégorie"
ImpTiragesCourse.Caption = "Impression des Tirages" & vbCrLf & "par Course"
ImpResultatsCateg.Caption = "Impression des Résultats" & vbCrLf & "par Catégorie"
ImpResultatsCourse.Caption = "Impression des Résultats" & vbCrLf & "par Course"
End Sub
