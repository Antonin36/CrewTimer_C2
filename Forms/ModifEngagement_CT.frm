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
    ID_Epreuve.value = Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 1).value
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "1" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "2" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "3" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value = Sexe_R3.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "4" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").value = Sexe_R4.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "5" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").value = Sexe_R5.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "6" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").value = Sexe_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").value = Nom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").value = Prenom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").value = Club_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").value = DN_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").value = Cat_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").value = Sexe_R6.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "7" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").value = Sexe_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").value = Nom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").value = Prenom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").value = Club_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").value = DN_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").value = Cat_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").value = Sexe_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BZ").value = Nom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CA").value = Prenom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CC").value = Club_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CG").value = DN_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CD").value = Cat_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CH").value = Sexe_R7.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "8" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value = ID_Ins.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value = ID_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value = Code_Epreuve.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value = Num_Bateau.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value = Nom_Equipage.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value = Nom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value = Prenom_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value = Club_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value = DN_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value = Cat_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value = Sexe_R1.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value = Nom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value = Prenom_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value = Club_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value = DN_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value = Cat_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value = Sexe_R2.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value = Nom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value = Prenom_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value = Club_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value = DN_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value = Cat_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value = Sexe_R3.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").value = Nom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").value = Prenom_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").value = Club_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").value = DN_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").value = Cat_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").value = Sexe_R4.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").value = Nom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").value = Prenom_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").value = Club_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").value = DN_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").value = Cat_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").value = Sexe_R5.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").value = Nom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").value = Prenom_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").value = Club_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").value = DN_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").value = Cat_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").value = Sexe_R6.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BZ").value = Nom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CA").value = Prenom_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CC").value = Club_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CG").value = DN_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CD").value = Cat_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CH").value = Sexe_R7.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CL").value = Nom_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CM").value = Prenom_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CO").value = Club_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CS").value = DN_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CP").value = Cat_R8.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CT").value = Sexe_R8.Text
    End If
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 5).value = "Oui" Then
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CX").value = Nom_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CY").value = Prenom_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DA").value = Club_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DE").value = DN_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DB").value = Cat_B.Text
    Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DF").value = Sexe_B.Text
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
    ID_Epreuve.value = Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 1).value
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "1" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "2" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "3" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "4" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "5" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "6" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "7" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 4).value = "8" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 5).value = "Oui" Then
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
    If Sheets("Stockage Epreuves CT").Cells(indexSelectionneCE, 5).value = "Non" Then
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
EngagementAModifier_CT = Sheets("R�glages R�gate").Cells(30, "D").value
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
   If Not IsEmpty(celluleCE.value) Then
       Code_Epreuve.AddItem celluleCE.value
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
ID_Ins.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "A").value
ID_Epreuve.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "B").value
Code_Epreuve.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value
Num_Bateau.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "D").value
Nom_Equipage.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "E").value
Nom_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "F").value
Prenom_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "G").value
Club_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "I").value
DN_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "M").value
Cat_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "J").value
Sexe_R1.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "N").value
Nom_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "R").value
Prenom_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "S").value
Club_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "U").value
DN_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Y").value
Cat_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "V").value
Sexe_R2.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "Z").value
Nom_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AD").value
Prenom_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AE").value
Club_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AG").value
DN_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AK").value
Cat_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AH").value
Sexe_R3.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AL").value
Nom_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AP").value
Prenom_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AQ").value
Club_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AS").value
DN_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AW").value
Cat_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AT").value
Sexe_R4.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "AX").value
Nom_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BB").value
Prenom_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BC").value
Club_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BE").value
DN_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BI").value
Cat_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BF").value
Sexe_R5.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BJ").value
Nom_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BN").value
Prenom_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BO").value
Club_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BP").value
DN_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BU").value
Cat_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BR").value
Sexe_R6.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BV").value
Nom_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "BZ").value
Prenom_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CA").value
Club_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CC").value
DN_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CG").value
Cat_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CD").value
Sexe_R7.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CH").value
Nom_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CL").value
Prenom_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CM").value
Club_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CO").value
DN_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CS").value
Cat_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CP").value
Sexe_R8.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CT").value
Nom_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CX").value
Prenom_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "CY").value
Club_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DA").value
DN_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DE").value
Cat_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DB").value
Sexe_B.Text = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "DF").value
'V�rif Code Epreuve
Dim valeurRecherchee As Variant
Dim feuilleRecherchee As Worksheet
Dim plageRecherche As Range
Dim cellule As Range
Dim ligneTrouvee As Long

    ' D�finir les param�tres de recherche
    valeurRecherchee = Sheets("Import GOAL CT").Cells(EngagementAModifier_CT, "C").value
    Set feuilleRecherchee = ThisWorkbook.Sheets("Stockage Epreuves CT")
    Set plageRecherche = feuilleRecherchee.Range("A2:A999") ' Modifier la plage selon vos besoins

    ' Parcourir chaque cellule de la plage de recherche
    For Each cellule In plageRecherche
        ' V�rifier si la valeur recherch�e est trouv�e
        If cellule.value = valeurRecherchee Then
            ' R�cup�rer le num�ro de ligne de la cellule trouv�e
            ligneTrouvee = cellule.Row
            Exit For
        End If
    Next cellule

    ID_Epreuve.value = Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 1).value
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "1" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "2" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "3" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "4" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "5" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "6" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "7" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 4).value = "8" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 5).value = "Oui" Then
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
    If Sheets("Stockage Epreuves CT").Cells(ligneTrouvee, 5).value = "Non" Then
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





