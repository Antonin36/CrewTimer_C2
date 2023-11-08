Note pour encoder une string json mettre d'abord tout le fichier sur une seule ligne, le séparer en morceaux, et le générer de manière procédurale. Puis sortir le fichier dans le répertoire courant et enfin reset la string.

Réglage pour toute la régate :
"event_name": "Nom Compétition"

Course Individuelle :
Par Equipage :
"affiliation" : Club Code Court
"class_name" : Code Catégorie
"lane_number" : Numéro Ergo (Idem que partant pour CT)
"name" : Nom Equipage (MONACO SN 1 (Nom Participant 1)) Idem CT
"participants" -> "name" : Nom Participant
Par course :
"duration" : Durée en Mètres
"name_long": "Nom de la course (Idem CT)",
"race_type": "individual",
"split_value": Durée Split,
"team_size": 1,
"duration_type": "meters"

Course Equipes :
Par Equipage :
"affiliation" : Club Code Court
"class_name" : Code Catégorie
"lane_number" : Numéro Ergo (Idem que partant pour CT)
"name" : Nom Equipage (MONACO SN 1 (Nom Participant 1)) Idem CT
"participants" -> "name" : Nom Participant
Par course :
"duration" : Durée en Mètres
"name_long": "Nom de la course (Idem CT)",
"race_type": "team",
"team_scoring": "avg",
"team_size": Taille Equipe (Nb d'Ergos en Simultané),
"split_value": Durée Split,
"duration_type": "meters"

Course Relais :
Par Equipage :
"affiliation" : Club Code Court
"class_name" : Code Catégorie
"lane_number" : Numéro Ergo (Idem que partant pour CT)
"name" : Nom Equipage (MONACO SN 1 (Nom Participant 1)) Idem CT
"participants" -> "name" : Nom Participant
Par course :
"duration" : Durée en Mètres
"name_long": "Nom de la course (Idem CT)",
"race_type": "relay",
"sound_horn_at_splits": false,
"split_value": Durée Split,
"team_size": 1,
"duration_type": "meters"

Classique :
"duration_type": "meters",
"duration" : Durée en M
"race_type": "individual",

Pour course parcourir la plus grande distance en X secondes :
"duration_type": "time",
"duration" : Durée en S
"race_type": "individual",

Pour course parcourir le plus grand nombre de calorie en X secondes Compatible Equipe mais pas Relais :
"race_type": "individual calorie score",
"duration_type": "time",
"duration" : Durée en S

Pour course en calories :
"duration": 60,
"duration_type": "calories",
"race_type": "individual",
