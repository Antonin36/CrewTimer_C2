VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifEngagement_CT 
   Caption         =   "Modification d'un Engagement"
   ClientHeight    =   11940
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   18680
   OleObjectBlob   =   "ModifEngagement_CT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifEngagement_CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private EngagementAModifier_CT As Long
Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Sauvegarder_Click()
    Dim indexSelectionneCE As Integer
    indexSelectionneCE = 0
    indexSelectionneCE = Code_Epreuve.ListIndex
    indexSelectionneCE = indexSelectionneCE + 2
    ID_Epreuve.Value = Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 1).Value
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "1" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "2" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "3" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value = Sexe_R3.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "4" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").Value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").Value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").Value = Sexe_R4.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "5" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").Value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").Value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").Value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").Value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").Value = Sexe_R5.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "6" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").Value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").Value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").Value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").Value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").Value = Sexe_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").Value = Nom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").Value = Prenom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").Value = Club_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").Value = DN_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").Value = Cat_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").Value = Sexe_R6.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "7" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").Value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").Value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").Value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").Value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").Value = Sexe_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").Value = Nom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").Value = Prenom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").Value = Club_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").Value = DN_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").Value = Cat_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").Value = Sexe_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BZ").Value = Nom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CA").Value = Prenom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CC").Value = Club_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CG").Value = DN_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CD").Value = Cat_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CH").Value = Sexe_R7.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "8" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").Value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").Value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").Value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").Value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").Value = Sexe_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").Value = Nom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").Value = Prenom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").Value = Club_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").Value = DN_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").Value = Cat_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").Value = Sexe_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BZ").Value = Nom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CA").Value = Prenom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CC").Value = Club_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CG").Value = DN_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CD").Value = Cat_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CH").Value = Sexe_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CL").Value = Nom_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CM").Value = Prenom_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CO").Value = Club_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CS").Value = DN_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CP").Value = Cat_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CT").Value = Sexe_R8.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 5).Value = "Oui" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CX").Value = Nom_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CY").Value = Prenom_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DA").Value = Club_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DE").Value = DN_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DB").Value = Cat_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DF").Value = Sexe_B.Text
    End If
    Sheets("Import GOAL CT").Select
    Range("A1").Select
    Sheets("Gestion CrewTimer").Select
    MsgBox "L'engagement � �t� modifi� avec succ�s !", vbOKOnly + vbInformation, "Engagement Modifi�"
      Unload Me
End Sub
Private Sub Code_Epreuve_Change()
    Dim indexSelectionneCE As Integer
    'Masquage
    'R1
    Label_N_R1.Visible = False
    Nom_R1.Visible = False
    Label_P_R1.Visible = False
    Prenom_R1.Visible = False
    Label_CA_R1.Visible = False
    Label_DN_R1.Visible = False
    Club_R1.Visible = False
    DN_R1.Visible = False
    Label_Cat_R1.Visible = False
    Label_S_R1.Visible = False
    Cat_R1.Visible = False
    Sexe_R1.Visible = False
    'R2
    Label_N_R2.Visible = False
    Nom_R2.Visible = False
    Label_P_R2.Visible = False
    Prenom_R2.Visible = False
    Label_CA_R2.Visible = False
    Label_DN_R2.Visible = False
    Club_R2.Visible = False
    DN_R2.Visible = False
    Label_Cat_R2.Visible = False
    Label_S_R2.Visible = False
    Cat_R2.Visible = False
    Sexe_R2.Visible = False
    'R3
    Label_N_R3.Visible = False
    Nom_R3.Visible = False
    Label_P_R3.Visible = False
    Prenom_R3.Visible = False
    Label_CA_R3.Visible = False
    Label_DN_R3.Visible = False
    Club_R3.Visible = False
    DN_R3.Visible = False
    Label_Cat_R3.Visible = False
    Label_S_R3.Visible = False
    Cat_R3.Visible = False
    Sexe_R3.Visible = False
    'R4
    Label_N_R4.Visible = False
    Nom_R4.Visible = False
    Label_P_R4.Visible = False
    Prenom_R4.Visible = False
    Label_CA_R4.Visible = False
    Label_DN_R4.Visible = False
    Club_R4.Visible = False
    DN_R4.Visible = False
    Label_Cat_R4.Visible = False
    Label_S_R4.Visible = False
    Cat_R4.Visible = False
    Sexe_R4.Visible = False
    'R5
    Label_N_R5.Visible = False
    Nom_R5.Visible = False
    Label_P_R5.Visible = False
    Prenom_R5.Visible = False
    Label_CA_R5.Visible = False
    Label_DN_R5.Visible = False
    Club_R5.Visible = False
    DN_R5.Visible = False
    Label_Cat_R5.Visible = False
    Label_S_R5.Visible = False
    Cat_R5.Visible = False
    Sexe_R5.Visible = False
    'R6
    Label_N_R6.Visible = False
    Nom_R6.Visible = False
    Label_P_R6.Visible = False
    Prenom_R6.Visible = False
    Label_CA_R6.Visible = False
    Label_DN_R6.Visible = False
    Club_R6.Visible = False
    DN_R6.Visible = False
    Label_Cat_R6.Visible = False
    Label_S_R6.Visible = False
    Cat_R6.Visible = False
    Sexe_R6.Visible = False
    'R7
    Label_N_R7.Visible = False
    Nom_R7.Visible = False
    Label_P_R7.Visible = False
    Prenom_R7.Visible = False
    Label_CA_R7.Visible = False
    Label_DN_R7.Visible = False
    Club_R7.Visible = False
    DN_R7.Visible = False
    Label_Cat_R7.Visible = False
    Label_S_R7.Visible = False
    Cat_R7.Visible = False
    Sexe_R7.Visible = False
    'R8
    Label_N_R8.Visible = False
    Nom_R8.Visible = False
    Label_P_R8.Visible = False
    Prenom_R8.Visible = False
    Label_CA_R8.Visible = False
    Label_DN_R8.Visible = False
    Club_R8.Visible = False
    DN_R8.Visible = False
    Label_Cat_R8.Visible = False
    Label_S_R8.Visible = False
    Cat_R8.Visible = False
    Sexe_R8.Visible = False
    'B
    Label_N_B.Visible = False
    Nom_B.Visible = False
    Label_P_B.Visible = False
    Prenom_B.Visible = False
    Label_CA_B.Visible = False
    Label_DN_B.Visible = False
    Club_B.Visible = False
    DN_B.Visible = False
    Label_Cat_B.Visible = False
    Label_S_B.Visible = False
    Cat_B.Visible = False
    Sexe_B.Visible = False
    'R�cup Index
    indexSelectionneCE = 0
    indexSelectionneCE = Code_Epreuve.ListIndex
    indexSelectionneCE = indexSelectionneCE + 2
    ID_Epreuve.Value = Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 1).Value
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "1" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "2" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "3" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "4" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "5" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "6" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    Label_N_R6.Visible = True
    Nom_R6.Visible = True
    Label_P_R6.Visible = True
    Prenom_R6.Visible = True
    Label_CA_R6.Visible = True
    Label_DN_R6.Visible = True
    Club_R6.Visible = True
    DN_R6.Visible = True
    Label_Cat_R6.Visible = True
    Label_S_R6.Visible = True
    Cat_R6.Visible = True
    Sexe_R6.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "7" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    Label_N_R6.Visible = True
    Nom_R6.Visible = True
    Label_P_R6.Visible = True
    Prenom_R6.Visible = True
    Label_CA_R6.Visible = True
    Label_DN_R6.Visible = True
    Club_R6.Visible = True
    DN_R6.Visible = True
    Label_Cat_R6.Visible = True
    Label_S_R6.Visible = True
    Cat_R6.Visible = True
    Sexe_R6.Visible = True
    Label_N_R7.Visible = True
    Nom_R7.Visible = True
    Label_P_R7.Visible = True
    Prenom_R7.Visible = True
    Label_CA_R7.Visible = True
    Label_DN_R7.Visible = True
    Club_R7.Visible = True
    DN_R7.Visible = True
    Label_Cat_R7.Visible = True
    Label_S_R7.Visible = True
    Cat_R7.Visible = True
    Sexe_R7.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).Value = "8" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    Label_N_R6.Visible = True
    Nom_R6.Visible = True
    Label_P_R6.Visible = True
    Prenom_R6.Visible = True
    Label_CA_R6.Visible = True
    Label_DN_R6.Visible = True
    Club_R6.Visible = True
    DN_R6.Visible = True
    Label_Cat_R6.Visible = True
    Label_S_R6.Visible = True
    Cat_R6.Visible = True
    Sexe_R6.Visible = True
    Label_N_R7.Visible = True
    Nom_R7.Visible = True
    Label_P_R7.Visible = True
    Prenom_R7.Visible = True
    Label_CA_R7.Visible = True
    Label_DN_R7.Visible = True
    Club_R7.Visible = True
    DN_R7.Visible = True
    Label_Cat_R7.Visible = True
    Label_S_R7.Visible = True
    Cat_R7.Visible = True
    Sexe_R7.Visible = True
    Label_N_R8.Visible = True
    Nom_R8.Visible = True
    Label_P_R8.Visible = True
    Prenom_R8.Visible = True
    Label_CA_R8.Visible = True
    Label_DN_R8.Visible = True
    Club_R8.Visible = True
    DN_R8.Visible = True
    Label_Cat_R8.Visible = True
    Label_S_R8.Visible = True
    Cat_R8.Visible = True
    Sexe_R8.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 5).Value = "Oui" Then
    Label_N_B.Visible = True
    Nom_B.Visible = True
    Label_P_B.Visible = True
    Prenom_B.Visible = True
    Label_CA_B.Visible = True
    Label_DN_B.Visible = True
    Club_B.Visible = True
    DN_B.Visible = True
    Label_Cat_B.Visible = True
    Label_S_B.Visible = True
    Cat_B.Visible = True
    Sexe_B.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 5).Value = "Non" Then
    Label_N_B.Visible = False
    Nom_B.Visible = False
    Label_P_B.Visible = False
    Prenom_B.Visible = False
    Label_CA_B.Visible = False
    Label_DN_B.Visible = False
    Club_B.Visible = False
    DN_B.Visible = False
    Label_Cat_B.Visible = False
    Label_S_B.Visible = False
    Cat_B.Visible = False
    Sexe_B.Visible = False
    End If
End Sub
Private Sub UserForm_Initialize()
EngagementAModifier_CT = Sheets("R�glages R�gate").Cells(30, "D").Value
'R1
Label_N_R1.Visible = False
Nom_R1.Visible = False
Label_P_R1.Visible = False
Prenom_R1.Visible = False
Label_CA_R1.Visible = False
Label_DN_R1.Visible = False
Club_R1.Visible = False
DN_R1.Visible = False
Label_Cat_R1.Visible = False
Label_S_R1.Visible = False
Cat_R1.Visible = False
Sexe_R1.Visible = False
'R2
Label_N_R2.Visible = False
Nom_R2.Visible = False
Label_P_R2.Visible = False
Prenom_R2.Visible = False
Label_CA_R2.Visible = False
Label_DN_R2.Visible = False
Club_R2.Visible = False
DN_R2.Visible = False
Label_Cat_R2.Visible = False
Label_S_R2.Visible = False
Cat_R2.Visible = False
Sexe_R2.Visible = False
'R3
Label_N_R3.Visible = False
Nom_R3.Visible = False
Label_P_R3.Visible = False
Prenom_R3.Visible = False
Label_CA_R3.Visible = False
Label_DN_R3.Visible = False
Club_R3.Visible = False
DN_R3.Visible = False
Label_Cat_R3.Visible = False
Label_S_R3.Visible = False
Cat_R3.Visible = False
Sexe_R3.Visible = False
'R4
Label_N_R4.Visible = False
Nom_R4.Visible = False
Label_P_R4.Visible = False
Prenom_R4.Visible = False
Label_CA_R4.Visible = False
Label_DN_R4.Visible = False
Club_R4.Visible = False
DN_R4.Visible = False
Label_Cat_R4.Visible = False
Label_S_R4.Visible = False
Cat_R4.Visible = False
Sexe_R4.Visible = False
'R5
Label_N_R5.Visible = False
Nom_R5.Visible = False
Label_P_R5.Visible = False
Prenom_R5.Visible = False
Label_CA_R5.Visible = False
Label_DN_R5.Visible = False
Club_R5.Visible = False
DN_R5.Visible = False
Label_Cat_R5.Visible = False
Label_S_R5.Visible = False
Cat_R5.Visible = False
Sexe_R5.Visible = False
'R6
Label_N_R6.Visible = False
Nom_R6.Visible = False
Label_P_R6.Visible = False
Prenom_R6.Visible = False
Label_CA_R6.Visible = False
Label_DN_R6.Visible = False
Club_R6.Visible = False
DN_R6.Visible = False
Label_Cat_R6.Visible = False
Label_S_R6.Visible = False
Cat_R6.Visible = False
Sexe_R6.Visible = False
'R7
Label_N_R7.Visible = False
Nom_R7.Visible = False
Label_P_R7.Visible = False
Prenom_R7.Visible = False
Label_CA_R7.Visible = False
Label_DN_R7.Visible = False
Club_R7.Visible = False
DN_R7.Visible = False
Label_Cat_R7.Visible = False
Label_S_R7.Visible = False
Cat_R7.Visible = False
Sexe_R7.Visible = False
'R8
Label_N_R8.Visible = False
Nom_R8.Visible = False
Label_P_R8.Visible = False
Prenom_R8.Visible = False
Label_CA_R8.Visible = False
Label_DN_R8.Visible = False
Club_R8.Visible = False
DN_R8.Visible = False
Label_Cat_R8.Visible = False
Label_S_R8.Visible = False
Cat_R8.Visible = False
Sexe_R8.Visible = False
'B
Label_N_B.Visible = False
Nom_B.Visible = False
Label_P_B.Visible = False
Prenom_B.Visible = False
Label_CA_B.Visible = False
Label_DN_B.Visible = False
Club_B.Visible = False
DN_B.Visible = False
Label_Cat_B.Visible = False
Label_S_B.Visible = False
Cat_B.Visible = False
Sexe_B.Visible = False
'Initialisation Code Epreuve
Dim feuilleCE As Worksheet
Set feuilleCE = ThisWorkbook.Sheets("Stockage Epreuves CT")

' D�finit la plage de donn�es � partir de la colonne A (de la ligne 2 � la ligne 999).
Dim plageCE As Range
Set plageCE = feuilleCE.Range("A2:A999")

' Parcours les cellules non vides de la plage et les ajoute � la ComboBox.
Dim celluleCE As Range
For Each celluleCE In plageCE
   If Not IsEmpty(celluleCE.Value) Then
       Code_Epreuve.AddItem celluleCE.Value
   End If
Next celluleCE
'Init List R1
Me.Cat_R1.AddItem ("Jeune (J10)")
Me.Cat_R1.AddItem ("Jeune (J11)")
Me.Cat_R1.AddItem ("Jeune (J12)")
Me.Cat_R1.AddItem ("Jeune (J13)")
Me.Cat_R1.AddItem ("Jeune (J14)")
Me.Cat_R1.AddItem ("Junior (J15)")
Me.Cat_R1.AddItem ("Junior (J16)")
Me.Cat_R1.AddItem ("Junior (J17)")
Me.Cat_R1.AddItem ("Junior (J18)")
Me.Cat_R1.AddItem ("S�nior -23")
Me.Cat_R1.AddItem ("S�nior")
Me.Sexe_R1.AddItem ("Homme")
Me.Sexe_R1.AddItem ("Femme")
'Init List R2
Me.Cat_R2.AddItem ("Jeune (J10)")
Me.Cat_R2.AddItem ("Jeune (J11)")
Me.Cat_R2.AddItem ("Jeune (J12)")
Me.Cat_R2.AddItem ("Jeune (J13)")
Me.Cat_R2.AddItem ("Jeune (J14)")
Me.Cat_R2.AddItem ("Junior (J15)")
Me.Cat_R2.AddItem ("Junior (J16)")
Me.Cat_R2.AddItem ("Junior (J17)")
Me.Cat_R2.AddItem ("Junior (J18)")
Me.Cat_R2.AddItem ("S�nior -23")
Me.Cat_R2.AddItem ("S�nior")
Me.Sexe_R2.AddItem ("Homme")
Me.Sexe_R2.AddItem ("Femme")
'Init List R3
Me.Cat_R3.AddItem ("Jeune (J10)")
Me.Cat_R3.AddItem ("Jeune (J11)")
Me.Cat_R3.AddItem ("Jeune (J12)")
Me.Cat_R3.AddItem ("Jeune (J13)")
Me.Cat_R3.AddItem ("Jeune (J14)")
Me.Cat_R3.AddItem ("Junior (J15)")
Me.Cat_R3.AddItem ("Junior (J16)")
Me.Cat_R3.AddItem ("Junior (J17)")
Me.Cat_R3.AddItem ("Junior (J18)")
Me.Cat_R3.AddItem ("S�nior -23")
Me.Cat_R3.AddItem ("S�nior")
Me.Sexe_R3.AddItem ("Homme")
Me.Sexe_R3.AddItem ("Femme")
'Init List R4
Me.Cat_R4.AddItem ("Jeune (J10)")
Me.Cat_R4.AddItem ("Jeune (J11)")
Me.Cat_R4.AddItem ("Jeune (J12)")
Me.Cat_R4.AddItem ("Jeune (J13)")
Me.Cat_R4.AddItem ("Jeune (J14)")
Me.Cat_R4.AddItem ("Junior (J15)")
Me.Cat_R4.AddItem ("Junior (J16)")
Me.Cat_R4.AddItem ("Junior (J17)")
Me.Cat_R4.AddItem ("Junior (J18)")
Me.Cat_R4.AddItem ("S�nior -23")
Me.Cat_R4.AddItem ("S�nior")
Me.Sexe_R4.AddItem ("Homme")
Me.Sexe_R4.AddItem ("Femme")
'Init List R5
Me.Cat_R5.AddItem ("Jeune (J10)")
Me.Cat_R5.AddItem ("Jeune (J11)")
Me.Cat_R5.AddItem ("Jeune (J12)")
Me.Cat_R5.AddItem ("Jeune (J13)")
Me.Cat_R5.AddItem ("Jeune (J14)")
Me.Cat_R5.AddItem ("Junior (J15)")
Me.Cat_R5.AddItem ("Junior (J16)")
Me.Cat_R5.AddItem ("Junior (J17)")
Me.Cat_R5.AddItem ("Junior (J18)")
Me.Cat_R5.AddItem ("S�nior -23")
Me.Cat_R5.AddItem ("S�nior")
Me.Sexe_R5.AddItem ("Homme")
Me.Sexe_R5.AddItem ("Femme")
'Init List R6
Me.Cat_R6.AddItem ("Jeune (J10)")
Me.Cat_R6.AddItem ("Jeune (J11)")
Me.Cat_R6.AddItem ("Jeune (J12)")
Me.Cat_R6.AddItem ("Jeune (J13)")
Me.Cat_R6.AddItem ("Jeune (J14)")
Me.Cat_R6.AddItem ("Junior (J15)")
Me.Cat_R6.AddItem ("Junior (J16)")
Me.Cat_R6.AddItem ("Junior (J17)")
Me.Cat_R6.AddItem ("Junior (J18)")
Me.Cat_R6.AddItem ("S�nior -23")
Me.Cat_R6.AddItem ("S�nior")
Me.Sexe_R6.AddItem ("Homme")
Me.Sexe_R6.AddItem ("Femme")
'Init List R7
Me.Cat_R7.AddItem ("Jeune (J10)")
Me.Cat_R7.AddItem ("Jeune (J11)")
Me.Cat_R7.AddItem ("Jeune (J12)")
Me.Cat_R7.AddItem ("Jeune (J13)")
Me.Cat_R7.AddItem ("Jeune (J14)")
Me.Cat_R7.AddItem ("Junior (J15)")
Me.Cat_R7.AddItem ("Junior (J16)")
Me.Cat_R7.AddItem ("Junior (J17)")
Me.Cat_R7.AddItem ("Junior (J18)")
Me.Cat_R7.AddItem ("S�nior -23")
Me.Cat_R7.AddItem ("S�nior")
Me.Sexe_R7.AddItem ("Homme")
Me.Sexe_R7.AddItem ("Femme")
'Init List R8
Me.Cat_R8.AddItem ("Jeune (J10)")
Me.Cat_R8.AddItem ("Jeune (J11)")
Me.Cat_R8.AddItem ("Jeune (J12)")
Me.Cat_R8.AddItem ("Jeune (J13)")
Me.Cat_R8.AddItem ("Jeune (J14)")
Me.Cat_R8.AddItem ("Junior (J15)")
Me.Cat_R8.AddItem ("Junior (J16)")
Me.Cat_R8.AddItem ("Junior (J17)")
Me.Cat_R8.AddItem ("Junior (J18)")
Me.Cat_R8.AddItem ("S�nior -23")
Me.Cat_R8.AddItem ("S�nior")
Me.Sexe_R8.AddItem ("Homme")
Me.Sexe_R8.AddItem ("Femme")
'Init List B
Me.Cat_B.AddItem ("Jeune (J10)")
Me.Cat_B.AddItem ("Jeune (J11)")
Me.Cat_B.AddItem ("Jeune (J12)")
Me.Cat_B.AddItem ("Jeune (J13)")
Me.Cat_B.AddItem ("Jeune (J14)")
Me.Cat_B.AddItem ("Junior (J15)")
Me.Cat_B.AddItem ("Junior (J16)")
Me.Cat_B.AddItem ("Junior (J17)")
Me.Cat_B.AddItem ("Junior (J18)")
Me.Cat_B.AddItem ("S�nior -23")
Me.Cat_B.AddItem ("S�nior")
Me.Sexe_B.AddItem ("Homme")
Me.Sexe_B.AddItem ("Femme")
ID_Ins.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").Value
ID_Epreuve.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").Value
Code_Epreuve.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value
Num_Bateau.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").Value
Nom_Equipage.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").Value
Nom_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").Value
Prenom_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").Value
Club_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").Value
DN_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").Value
Cat_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").Value
Sexe_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").Value
Nom_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").Value
Prenom_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").Value
Club_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").Value
DN_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").Value
Cat_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").Value
Sexe_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").Value
Nom_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").Value
Prenom_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").Value
Club_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").Value
DN_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").Value
Cat_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").Value
Sexe_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").Value
Nom_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").Value
Prenom_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").Value
Club_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").Value
DN_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").Value
Cat_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").Value
Sexe_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").Value
Nom_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").Value
Prenom_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").Value
Club_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").Value
DN_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").Value
Cat_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").Value
Sexe_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").Value
Nom_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").Value
Prenom_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").Value
Club_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").Value
DN_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").Value
Cat_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").Value
Sexe_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").Value
Nom_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BZ").Value
Prenom_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CA").Value
Club_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CC").Value
DN_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CG").Value
Cat_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CD").Value
Sexe_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CH").Value
Nom_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CL").Value
Prenom_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CM").Value
Club_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CO").Value
DN_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CS").Value
Cat_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CP").Value
Sexe_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CT").Value
Nom_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CX").Value
Prenom_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CY").Value
Club_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DA").Value
DN_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DE").Value
Cat_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DB").Value
Sexe_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DF").Value
'V�rif Code Epreuve
Dim valeurRecherchee As Variant
Dim feuilleRecherchee As Worksheet
Dim plageRecherche As Range
Dim cellule As Range
Dim ligneTrouvee As Long

    ' D�finir les param�tres de recherche
    valeurRecherchee = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").Value
    Set feuilleRecherchee = ThisWorkbook.Sheets("Stockage Epreuves CT")
    Set plageRecherche = feuilleRecherchee.Range("A2:A999") ' Modifier la plage selon vos besoins

    ' Parcourir chaque cellule de la plage de recherche
    For Each cellule In plageRecherche
        ' V�rifier si la valeur recherch�e est trouv�e
        If cellule.Value = valeurRecherchee Then
            ' R�cup�rer le num�ro de ligne de la cellule trouv�e
            ligneTrouvee = cellule.Row
            Exit For
        End If
    Next cellule

    ID_Epreuve.Value = Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 1).Value
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "1" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "2" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "3" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "4" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "5" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "6" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    Label_N_R6.Visible = True
    Nom_R6.Visible = True
    Label_P_R6.Visible = True
    Prenom_R6.Visible = True
    Label_CA_R6.Visible = True
    Label_DN_R6.Visible = True
    Club_R6.Visible = True
    DN_R6.Visible = True
    Label_Cat_R6.Visible = True
    Label_S_R6.Visible = True
    Cat_R6.Visible = True
    Sexe_R6.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "7" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    Label_N_R6.Visible = True
    Nom_R6.Visible = True
    Label_P_R6.Visible = True
    Prenom_R6.Visible = True
    Label_CA_R6.Visible = True
    Label_DN_R6.Visible = True
    Club_R6.Visible = True
    DN_R6.Visible = True
    Label_Cat_R6.Visible = True
    Label_S_R6.Visible = True
    Cat_R6.Visible = True
    Sexe_R6.Visible = True
    Label_N_R7.Visible = True
    Nom_R7.Visible = True
    Label_P_R7.Visible = True
    Prenom_R7.Visible = True
    Label_CA_R7.Visible = True
    Label_DN_R7.Visible = True
    Club_R7.Visible = True
    DN_R7.Visible = True
    Label_Cat_R7.Visible = True
    Label_S_R7.Visible = True
    Cat_R7.Visible = True
    Sexe_R7.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).Value = "8" Then
    Label_N_R1.Visible = True
    Nom_R1.Visible = True
    Label_P_R1.Visible = True
    Prenom_R1.Visible = True
    Label_CA_R1.Visible = True
    Label_DN_R1.Visible = True
    Club_R1.Visible = True
    DN_R1.Visible = True
    Label_Cat_R1.Visible = True
    Label_S_R1.Visible = True
    Cat_R1.Visible = True
    Sexe_R1.Visible = True
    Label_N_R2.Visible = True
    Nom_R2.Visible = True
    Label_P_R2.Visible = True
    Prenom_R2.Visible = True
    Label_CA_R2.Visible = True
    Label_DN_R2.Visible = True
    Club_R2.Visible = True
    DN_R2.Visible = True
    Label_Cat_R2.Visible = True
    Label_S_R2.Visible = True
    Cat_R2.Visible = True
    Sexe_R2.Visible = True
    Label_N_R3.Visible = True
    Nom_R3.Visible = True
    Label_P_R3.Visible = True
    Prenom_R3.Visible = True
    Label_CA_R3.Visible = True
    Label_DN_R3.Visible = True
    Club_R3.Visible = True
    DN_R3.Visible = True
    Label_Cat_R3.Visible = True
    Label_S_R3.Visible = True
    Cat_R3.Visible = True
    Sexe_R3.Visible = True
    Label_N_R4.Visible = True
    Nom_R4.Visible = True
    Label_P_R4.Visible = True
    Prenom_R4.Visible = True
    Label_CA_R4.Visible = True
    Label_DN_R4.Visible = True
    Club_R4.Visible = True
    DN_R4.Visible = True
    Label_Cat_R4.Visible = True
    Label_S_R4.Visible = True
    Cat_R4.Visible = True
    Sexe_R4.Visible = True
    Label_N_R5.Visible = True
    Nom_R5.Visible = True
    Label_P_R5.Visible = True
    Prenom_R5.Visible = True
    Label_CA_R5.Visible = True
    Label_DN_R5.Visible = True
    Club_R5.Visible = True
    DN_R5.Visible = True
    Label_Cat_R5.Visible = True
    Label_S_R5.Visible = True
    Cat_R5.Visible = True
    Sexe_R5.Visible = True
    Label_N_R6.Visible = True
    Nom_R6.Visible = True
    Label_P_R6.Visible = True
    Prenom_R6.Visible = True
    Label_CA_R6.Visible = True
    Label_DN_R6.Visible = True
    Club_R6.Visible = True
    DN_R6.Visible = True
    Label_Cat_R6.Visible = True
    Label_S_R6.Visible = True
    Cat_R6.Visible = True
    Sexe_R6.Visible = True
    Label_N_R7.Visible = True
    Nom_R7.Visible = True
    Label_P_R7.Visible = True
    Prenom_R7.Visible = True
    Label_CA_R7.Visible = True
    Label_DN_R7.Visible = True
    Club_R7.Visible = True
    DN_R7.Visible = True
    Label_Cat_R7.Visible = True
    Label_S_R7.Visible = True
    Cat_R7.Visible = True
    Sexe_R7.Visible = True
    Label_N_R8.Visible = True
    Nom_R8.Visible = True
    Label_P_R8.Visible = True
    Prenom_R8.Visible = True
    Label_CA_R8.Visible = True
    Label_DN_R8.Visible = True
    Club_R8.Visible = True
    DN_R8.Visible = True
    Label_Cat_R8.Visible = True
    Label_S_R8.Visible = True
    Cat_R8.Visible = True
    Sexe_R8.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 5).Value = "Oui" Then
    Label_N_B.Visible = True
    Nom_B.Visible = True
    Label_P_B.Visible = True
    Prenom_B.Visible = True
    Label_CA_B.Visible = True
    Label_DN_B.Visible = True
    Club_B.Visible = True
    DN_B.Visible = True
    Label_Cat_B.Visible = True
    Label_S_B.Visible = True
    Cat_B.Visible = True
    Sexe_B.Visible = True
    End If
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 5).Value = "Non" Then
    Label_N_B.Visible = False
    Nom_B.Visible = False
    Label_P_B.Visible = False
    Prenom_B.Visible = False
    Label_CA_B.Visible = False
    Label_DN_B.Visible = False
    Club_B.Visible = False
    DN_B.Visible = False
    Label_Cat_B.Visible = False
    Label_S_B.Visible = False
    Cat_B.Visible = False
    Sexe_B.Visible = False
    End If
End Sub





