VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Start 
   Caption         =   "D�marrage du Syst�me"
   ClientHeight    =   5544
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   6500
   OleObjectBlob   =   "Start.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continuer_Click()
    Sheets("Accueil").Select
    Unload Me
End Sub

Private Sub ReInit_Click()
Dim answer1 As Integer
answer1 = MsgBox("Etes-vous certain de vouloir r�initialiser TOUTE la r�gate ?", vbYesNo + vbExclamation, "Demande de confirmation")
  If answer1 = vbYes Then
        Sheets("R�glages R�gate").Range("D4").Value = ""
        Sheets("R�glages R�gate").Range("D6").Value = ""
        Sheets("R�glages R�gate").Range("D8").Value = ""
        Sheets("R�glages R�gate").Range("E14").Value = ""
        Sheets("R�glages R�gate").Range("E18").Value = ""
        Sheets("R�glages R�gate").Range("E16").Value = ""
        Sheets("R�glages R�gate").Range("K4").Value = ""
        Sheets("R�glages R�gate").Range("K6").Value = ""
        Sheets("Pr�paration Tirages CT").Select
        Range("A2:K999").Select
        Selection.ClearContents
        Sheets("Feuille CrewTimer").Select
        Range("A8:K999").Select
        Selection.ClearContents
        Sheets("Import GOAL CT").Select
        Range("A1:FA9999").Select
        Sheets("Stockage Impressions CT").Select
        Range("A1:FA9999").Select
        Selection.ClearContents
        Sheets("Import Tirages CT").Select
        Range("A1:FA9999").Select
        Selection.ClearContents
        Sheets("Import Resultats CT").Select
        Range("A1:FA9999").Select
        Selection.ClearContents
        Sheets("Impressions R�sultats CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Impressions Tirages CT").Select
        Range("A13:H420").Select
        Selection.ClearContents
        Sheets("Programme des Courses CT").Select
        Range("A2:FA9999").Select
        Selection.ClearContents
        Sheets("Accueil").Select
        ReglagesRegate.Show
        Unload Me
  Else
    Exit Sub
  End If
End Sub

Private Sub UserForm_Initialize()
    Sheets("Accueil").Select
    Regate.Caption = Sheets("R�glages R�gate").Range("D4").Value
    Lieu.Caption = Sheets("R�glages R�gate").Range("D6").Value
    Club.Caption = Sheets("R�glages R�gate").Range("D8").Value
End Sub
