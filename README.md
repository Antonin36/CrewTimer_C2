# CrewTimer_C2
Un système de gestion de course avec la possibilité d'importer directement les fichiers. \
\
Fonctions Implantées : 

Générales : 
- Jusqu'à 100 partants par course, 
- Jusqu'à 40 épreuves sur la même régate, (Augmentation jusqu'à 100-150 prévue.)
- Jusqu'à 200 courses. 

CrewTimer :

- Import depuis GOAL FFSA,
- Gestion des Inscriptions,
- Tirages Automatiques (Aléatoire, Par Numéro de Bateau et Par Ordre Alphabétique des Noms Courts des Clubs, Prise en Charge TDR), 
- Import Résultats depuis CrewTimer, 
- Impressions Tirages et Résultats (Jusqu'à 17 pages, augmentable/diminutible si besoin.), 
- Gestion des Epreuves,
- Gestion du Programme des Courses. 

Concept2 : 

- Import depuis GOAL FFSA,
- Gestion des Inscriptions,
- Tirages Automatiques (Aléatoire, Par Numéro Défini et Par Ordre Alphabétique des Noms Courts des Clubs.) ,
- Gestion de tous les types de courses prévus dans ErgRace (Individuel, Par Equipe, En Relais), 
- Gestion de tous les types de distance prévus dans ErgRace (Calories, Distance, Max de Distance sur un temps donné, etc...), 
- Impressions Tirages (Jusqu'à 17 pages, augmentable/diminutible si besoin.), 
- Gestion des Epreuves,
- Gestion du Programme des Courses, 
- Génération des Fichiers RAC2 pour importation dans ErgRace. \

Fonctions à Implanter : \
\
CrewTimer et Concept2 : \
\
Import depuis OPUSS, \
Import depuis FFSU, \
Transposer dans les tables export (L'ID inscription, l'ID Course, bref tout les paramètres de GOAL...), \
Webservices GOAL (Si Possible...), prévu : Contrôle Validité, Catégorie, Récupération des Infos Manifestations depuis GOAL via ID Manifestation, \
Analyser les requêtes HTTP/HTTPS de GOAL, pour voir le process d'auth, et le token, \
Modifier les fonctions de modif pour effacer les cases des ListBox, \
Simplifier les impressions, voir CR Indoor PACA 2023, \
Ajouter l'impression de la grille horaire, \
Refaire le module pretirages, ajouter un mode manuel via ID Inscription, \
Ajouter un contrôle des valeurs pour les ID. \
Pour tous bugs trouvés, merci d'ouvrir une Issue.
