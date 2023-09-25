VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReglagesRegate 
   Caption         =   "R�glages de la R�gate"
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

TitreRegate.Text = Sheets("R�glages R�gate").Range("D4").Value
LieuRegate.Text = Sheets("R�glages R�gate").Range("D6").Value
ClubOrganisateur.Text = Sheets("R�glages R�gate").Range("D8").Value
NBPartants.Text = Sheets("R�glages R�gate").Range("E14").Value
Affiliation.Text = Sheets("R�glages R�gate").Range("E18").Value
TypeRegate.Text = Sheets("R�glages R�gate").Range("E16").Value
DateDebut.Text = Sheets("R�glages R�gate").Range("K4").Value
DateFin.Text = Sheets("R�glages R�gate").Range("K6").Value


'Remplissage Valeurs Affiliation
With Affiliation
    .AddItem "FFAviron"
    .AddItem "UNSS/FFSU"
    .AddItem "UNSS"
    .AddItem "FFSU"
End With

'Remplissage Valeurs Type R�gate
With TypeRegate
    .AddItem "Rivi�re"
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
If MsgBox("Etes-vous certain de vouloir r�initialiser TOUTE la r�gate ?", vbYesNo + vbExclamation, "Demande de confirmation") = vbYes Then
        Sheets("R�glages R�gate").Range("D4").Value = ""
        Sheets("R�glages R�gate").Range("D6").Value = ""
        Sheets("R�glages R�gate").Range("D8").Value = ""
        Sheets("R�glages R�gate").Range("E14").Value = ""
        Sheets("R�glages R�gate").Range("E18").Value = ""
        Sheets("R�glages R�gate").Range("E16").Value = ""
        Sheets("R�glages R�gate").Range("K4").Value = ""
        Sheets("R�glages R�gate").Range("K6").Value = ""
        Sheets("Pr�paration Tirages").Select
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
        Sheets("Impressions R�sultats CT").Select
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
        MsgBox "La r�gate � bien �t� r�initialis�e !", vbInformation + vbOKOnly, "R�gate R�initialis�e"
        Call UserForm_Initialize
    End If
    
End Sub
Private Sub Sauvegarder_Click()
'R�cup Titre R�gate
TitreRegate = TitreRegate.Text

'Inscription Titre Regate
Sheets("R�glages R�gate").Range("D4").Value = TitreRegate

'R�cup Lieu R�gate
LieuRegate = LieuRegate.Text

'Inscription Lieu Regate
Sheets("R�glages R�gate").Range("D6").Value = LieuRegate

'R�cup Club Orga
ClubOrga = ClubOrganisateur.Text

'Inscription Lieu Regate
Sheets("R�glages R�gate").Range("D8").Value = ClubOrga

'R�cup Date Debut
DateDebut = DateDebut.Text

'Inscription Date Debut
Sheets("R�glages R�gate").Range("K4").Value = DateDebut

'R�cup Date Fin
DateFin = DateFin.Text

'Inscription Date Fin
Sheets("R�glages R�gate").Range("K6").Value = DateFin

'R�cup Type R�gate
TypeRegate = TypeRegate.Text

'Inscription Type R�gate
Sheets("R�glages R�gate").Range("E16").Value = TypeRegate

'R�cup Fede
Fede = Affiliation.Text

'Inscription Fede
Sheets("R�glages R�gate").Range("E18").Value = Fede

'R�cup NB Partants
NBPartants = NBPartants.Text

'Inscription NB Partants
Sheets("R�glages R�gate").Range("E14").Value = NBPartants

MsgBox "Les r�glages de la r�gate ont �t� sauvegard�s avec succ�s !", vbOKOnly + vbInformation, "R�glages Sauvegard�s"

Unload Me

End Sub
