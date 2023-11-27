VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GestIns_CT 
   Caption         =   "Gestion des Inscriptions"
   ClientHeight    =   4440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6380
   OleObjectBlob   =   "GestIns_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GestIns_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GestionEngagements_Click()
    GestEngagements_CT.Show
End Sub

Private Sub GestionEpreuves_Click()
    GestEpreuves_CT.Show
End Sub

Private Sub RetourAccueil_Click()
    Unload Me
End Sub
