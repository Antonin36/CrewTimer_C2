VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReglagesRegate 
   Caption         =   "Réglages de la Régate"
   ClientHeight    =   5660
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9384.001
   OleObjectBlob   =   "ReglagesRegate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReglagesRegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

TitreRegate.Text = Sheets("Réglages Régate").Range("D4").Value
LieuRegate.Text = Sheets("Réglages Régate").Range("D6").Value
ClubOrganisateur.Text = Sheets("Réglages Régate").Range("D8").Value
NBPartants.Text = Sheets("Réglages Régate").Range("E14").Value
Affiliation.Text = Sheets("Réglages Régate").Range("E18").Value
TypeRegate.Text = Sheets("Réglages Régate").Range("E16").Value
DateDebut.Text = Sheets("Réglages Régate").Range("K4").Value
DateFin.Text = Sheets("Réglages Régate").Range("K6").Value


'Remplissage Valeurs Affiliation
With Affiliation
    .AddItem "FFAviron"
    .AddItem "UNSS/FFSU"
    .AddItem "UNSS"
    .AddItem "FFSU"
End With

'Remplissage Valeurs Type Régate
With TypeRegate
    .AddItem "Rivière"
    .AddItem "Mer"
    .AddItem "Indoor"
End With

'Remplissage NB Couloirs/Partants
With NBPartants
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
End With
End Sub

Private Sub ReInit_Click()
If MsgBox("Etes-vous certain de vouloir réinitialiser TOUTE la régate ?", vbYesNo + vbExclamation, "Demande de confirmation") = vbYes Then
        Sheets("Réglages Régate").Range("D4").Value = ""
        Sheets("Réglages Régate").Range("D6").Value = ""
        Sheets("Réglages Régate").Range("D8").Value = ""
        Sheets("Réglages Régate").Range("E14").Value = ""
        Sheets("Réglages Régate").Range("E18").Value = ""
        Sheets("Réglages Régate").Range("E16").Value = ""
        Sheets("Réglages Régate").Range("K4").Value = ""
        Sheets("Réglages Régate").Range("K6").Value = ""
        Sheets("Préparation Tirages").Select
        Range("A2:K999").Select
        Selection.ClearContents
        Sheets("Feuille CrewTimer").Select
        Range("A8:K999").Select
        Selection.ClearContents
        Sheets("Import GOAL").Select
        Range("A1:FA9999").Select
        Sheets("Stockage Impressions").Select
        Range("A1:FA9999").Select
        Selection.ClearContents
        Sheets("Import Tirages").Select
        Range("A1:FA9999").Select
        Selection.ClearContents
        Sheets("Import Resultats").Select
        Range("A1:FA9999").Select
        Selection.ClearContents
        Sheets("Impressions Résultats CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Impressions Tirages CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Programme des Courses").Select
        Range("A2:FA9999").Select
        Selection.ClearContents
        Sheets("Accueil").Select
        Affiliation.Clear
        TypeRegate.Clear
        NBPartants.Clear
        MsgBox "La régate à bien été réinitialisée !", vbInformation + vbOKOnly, "Régate Réinitialisée"
        Call UserForm_Initialize
    End If
    
End Sub
Private Sub Sauvegarder_Click()
'Récup Titre Régate
TitreRegate = TitreRegate.Text

'Inscription Titre Regate
Sheets("Réglages Régate").Range("D4").Value = TitreRegate

'Récup Lieu Régate
LieuRegate = LieuRegate.Text

'Inscription Lieu Regate
Sheets("Réglages Régate").Range("D6").Value = LieuRegate

'Récup Club Orga
ClubOrga = ClubOrganisateur.Text

'Inscription Lieu Regate
Sheets("Réglages Régate").Range("D8").Value = ClubOrga

'Récup Date Debut
DateDebut = DateDebut.Text

'Inscription Date Debut
Sheets("Réglages Régate").Range("K4").Value = DateDebut

'Récup Date Fin
DateFin = DateFin.Text

'Inscription Date Fin
Sheets("Réglages Régate").Range("K6").Value = DateFin

'Récup Type Régate
TypeRegate = TypeRegate.Text

'Inscription Type Régate
Sheets("Réglages Régate").Range("E16").Value = TypeRegate

'Récup Fede
Fede = Affiliation.Text

'Inscription Fede
Sheets("Réglages Régate").Range("E18").Value = Fede

'Récup NB Partants
NBPartants = NBPartants.Text

'Inscription NB Partants
Sheets("Réglages Régate").Range("E14").Value = NBPartants

MsgBox "Les réglages de la régate ont été sauvegardés avec succès !", vbOKOnly + vbInformation, "Réglages Sauvegardés"

Unload Me

End Sub
