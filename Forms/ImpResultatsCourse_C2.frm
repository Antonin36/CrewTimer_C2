VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImpResultatsCourse_C2 
   Caption         =   "Impression des R�sultats"
   ClientHeight    =   5640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7980
   OleObjectBlob   =   "ImpResultatsCourse_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImpResultatsCourse_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
MsgBox "Pour imprimer les r�sultats par course, vous pouvez soit, n'importer que le fichier de la course, soit imprimer depuis ErgRace.", vbOKOnly + vbInformation, "Impression Non Normale"
Unload Me
End Sub






