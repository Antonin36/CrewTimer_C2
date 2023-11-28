VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifEngagement_C2 
   Caption         =   "Modification d'un Engagement"
   ClientHeight    =   11940
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   18680
   OleObjectBlob   =   "ModifEngagement_C2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifEngagement_C2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private EngagementAModifier_C2 As Long
Private Sub Annuler_Click()
    Unload Me
End Sub
Private Sub Sauvegarder_Click()
    Dim indexSelectionneCE As Integer
    indexSelectionneCE = 0
    indexSelectionneCE = Code_Epreuve.ListIndex
    indexSelectionneCE = indexSelectionneCE + 2
    ID_Epreuve.Value = Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 1).Value
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "1" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "2" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "3" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value = Club_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value = DN_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value = Sexe_R3.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "4" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value = Club_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value = DN_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AS").Value = Club_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AW").Value = DN_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AX").Value = Sexe_R4.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "5" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value = Club_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value = DN_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AS").Value = Club_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AW").Value = DN_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BE").Value = Club_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BI").Value = DN_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BJ").Value = Sexe_R5.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "6" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value = Club_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value = DN_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AS").Value = Club_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AW").Value = DN_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BE").Value = Club_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BI").Value = DN_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BJ").Value = Sexe_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BN").Value = Nom_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BO").Value = Prenom_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BP").Value = Club_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BU").Value = DN_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BR").Value = Cat_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BV").Value = Sexe_R6.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "7" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value = Club_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value = DN_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AS").Value = Club_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AW").Value = DN_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BE").Value = Club_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BI").Value = DN_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BJ").Value = Sexe_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BN").Value = Nom_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BO").Value = Prenom_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BP").Value = Club_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BU").Value = DN_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BR").Value = Cat_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BV").Value = Sexe_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BZ").Value = Nom_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CA").Value = Prenom_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CC").Value = Club_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CG").Value = DN_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CD").Value = Cat_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CH").Value = Sexe_R7.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "8" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value = ID_Ins.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value = ID_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value = Code_Epreuve.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value = Num_Bateau.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value = Nom_Equipage.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value = Nom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value = Prenom_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value = Club_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value = DN_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value = Cat_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value = Sexe_R1.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value = Nom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value = Prenom_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value = Club_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value = DN_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value = Cat_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value = Sexe_R2.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value = Nom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value = Prenom_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value = Club_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value = DN_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value = Cat_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value = Sexe_R3.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AP").Value = Nom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AQ").Value = Prenom_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AS").Value = Club_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AW").Value = DN_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AT").Value = Cat_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AX").Value = Sexe_R4.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BB").Value = Nom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BC").Value = Prenom_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BE").Value = Club_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BI").Value = DN_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BF").Value = Cat_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BJ").Value = Sexe_R5.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BN").Value = Nom_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BO").Value = Prenom_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BP").Value = Club_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BU").Value = DN_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BR").Value = Cat_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BV").Value = Sexe_R6.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BZ").Value = Nom_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CA").Value = Prenom_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CC").Value = Club_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CG").Value = DN_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CD").Value = Cat_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CH").Value = Sexe_R7.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CL").Value = Nom_R8.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CM").Value = Prenom_R8.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CO").Value = Club_R8.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CS").Value = DN_R8.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CP").Value = Cat_R8.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CT").Value = Sexe_R8.Text
    End If
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 5).Value = "Oui" Then
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CX").Value = Nom_B.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CY").Value = Prenom_B.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DA").Value = Club_B.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DE").Value = DN_B.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DB").Value = Cat_B.Text
    Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DF").Value = Sexe_B.Text
    End If
    Sheets("Import GOAL C2").Select
    Range("A1").Select
    Sheets("Gestion Concept2").Select
    MsgBox "L'engagement à été modifié avec succès !", vbOKOnly + vbInformation, "Engagement Modifié"
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
    'Récup Index
    indexSelectionneCE = 0
    indexSelectionneCE = Code_Epreuve.ListIndex
    indexSelectionneCE = indexSelectionneCE + 2
    ID_Epreuve.Value = Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 1).Value
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "1" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "2" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "3" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "4" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "5" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "6" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "7" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 4).Value = "8" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 5).Value = "Oui" Then
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
    If Sheets("Stockage Epreuves C2").Cells(indexSelectionneCE, 5).Value = "Non" Then
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
EngagementAModifier_C2 = Sheets("Réglages Régate").Cells(31, "D").Value
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
Set feuilleCE = ThisWorkbook.Sheets("Stockage Epreuves C2")

' Définit la plage de données à partir de la colonne A (de la ligne 2 à la ligne 999).
Dim plageCE As Range
Set plageCE = feuilleCE.Range("A2:A999")

' Parcours les cellules non vides de la plage et les ajoute à la ComboBox.
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
Me.Cat_R1.AddItem ("Sénior -23")
Me.Cat_R1.AddItem ("Sénior")
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
Me.Cat_R2.AddItem ("Sénior -23")
Me.Cat_R2.AddItem ("Sénior")
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
Me.Cat_R3.AddItem ("Sénior -23")
Me.Cat_R3.AddItem ("Sénior")
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
Me.Cat_R4.AddItem ("Sénior -23")
Me.Cat_R4.AddItem ("Sénior")
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
Me.Cat_R5.AddItem ("Sénior -23")
Me.Cat_R5.AddItem ("Sénior")
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
Me.Cat_R6.AddItem ("Sénior -23")
Me.Cat_R6.AddItem ("Sénior")
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
Me.Cat_R7.AddItem ("Sénior -23")
Me.Cat_R7.AddItem ("Sénior")
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
Me.Cat_R8.AddItem ("Sénior -23")
Me.Cat_R8.AddItem ("Sénior")
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
Me.Cat_B.AddItem ("Sénior -23")
Me.Cat_B.AddItem ("Sénior")
Me.Sexe_B.AddItem ("Homme")
Me.Sexe_B.AddItem ("Femme")
ID_Ins.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "A").Value
ID_Epreuve.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "B").Value
Code_Epreuve.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value
Num_Bateau.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "D").Value
Nom_Equipage.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "E").Value
Nom_R1.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "F").Value
Prenom_R1.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "G").Value
Club_R1.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "I").Value
DN_R1.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "M").Value
Cat_R1.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "J").Value
Sexe_R1.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "N").Value
Nom_R2.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "R").Value
Prenom_R2.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "S").Value
Club_R2.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "U").Value
DN_R2.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Y").Value
Cat_R2.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "V").Value
Sexe_R2.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "Z").Value
Nom_R3.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AD").Value
Prenom_R3.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AE").Value
Club_R3.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AG").Value
DN_R3.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AK").Value
Cat_R3.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AH").Value
Sexe_R3.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AL").Value
Nom_R4.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AP").Value
Prenom_R4.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AQ").Value
Club_R4.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AS").Value
DN_R4.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AW").Value
Cat_R4.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AT").Value
Sexe_R4.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "AX").Value
Nom_R5.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BB").Value
Prenom_R5.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BC").Value
Club_R5.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BE").Value
DN_R5.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BI").Value
Cat_R5.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BF").Value
Sexe_R5.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BJ").Value
Nom_R6.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BN").Value
Prenom_R6.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BO").Value
Club_R6.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BP").Value
DN_R6.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BU").Value
Cat_R6.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BR").Value
Sexe_R6.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BV").Value
Nom_R7.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "BZ").Value
Prenom_R7.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CA").Value
Club_R7.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CC").Value
DN_R7.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CG").Value
Cat_R7.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CD").Value
Sexe_R7.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CH").Value
Nom_R8.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CL").Value
Prenom_R8.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CM").Value
Club_R8.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CO").Value
DN_R8.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CS").Value
Cat_R8.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CP").Value
Sexe_R8.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CT").Value
Nom_B.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CX").Value
Prenom_B.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "CY").Value
Club_B.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DA").Value
DN_B.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DE").Value
Cat_B.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DB").Value
Sexe_B.Text = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "DF").Value
'Vérif Code Epreuve
Dim valeurRecherchee As Variant
Dim feuilleRecherchee As Worksheet
Dim plageRecherche As Range
Dim cellule As Range
Dim ligneTrouvee As Long

    ' Définir les paramètres de recherche
    valeurRecherchee = Sheets("Import GOAL C2").Cells(EngagementAModifier_C2, "C").Value
    Set feuilleRecherchee = ThisWorkbook.Sheets("Stockage Epreuves C2")
    Set plageRecherche = feuilleRecherchee.Range("A1:A999") ' Modifier la plage selon vos besoins

    ' Parcourir chaque cellule de la plage de recherche
    For Each cellule In plageRecherche
        ' Vérifier si la valeur recherchée est trouvée
        If cellule.Value = valeurRecherchee Then
            ' Récupérer le numéro de ligne de la cellule trouvée
            ligneTrouvee = cellule.Row
            Exit For
        End If
    Next cellule

    ID_Epreuve.Value = Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 1).Value
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "1" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "2" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "3" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "4" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "5" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "6" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "7" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 4).Value = "8" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 5).Value = "Oui" Then
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
    If Sheets("Stockage Epreuves C2").Cells(ligneTrouvee, 5).Value = "Non" Then
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




