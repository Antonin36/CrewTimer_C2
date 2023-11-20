# CrewTimer_C2
Un système de gestion de course avec la possibilité d'importer directement les fichiers. \
\
Fonctions Implantées : 

Générales : 
- Jusqu'à 100 partants par course, 
- Jusqu'à 40 catégories sur la même régate, 
- Jusqu'à 200 courses. 

CrewTimer :

- Import depuis GOAL FFSA 
- Tirages Automatiques (Aléatoire, Par Numéro de Bateau et Par Ordre Alphabétique des Noms Courts des Clubs, Prise en Charge TDR), 
- Import Résultats depuis CrewTimer, 
- Impressions Tirages et Résultats (Jusqu'à 17 pages, augmentable/diminutible si besoin.), 
- Gestion du Programme des Courses. 

Concept2 : 

- Import depuis GOAL FFSA, 
- Tirages Automatiques (Aléatoire, Par Numéro Défini et Par Ordre Alphabétique des Noms Courts des Clubs.) 
- Gestion de tous les types de courses prévus dans ErgRace (Individuel, Par Equipe, En Relais), 
- Gestion de tous les types de distance prévus dans ErgRace (Calories, Distance, Max de Distance sur un temps donné, etc...), 
- Impressions Tirages (Jusqu'à 17 pages, augmentable/diminutible si besoin.), 
- Gestion du Programme des Courses, 
- Génération des Fichiers RAC2 pour importation dans ErgRace. \
Note Importante : L'impression des Résultats doit se faire dans ErgRace !

Fonctions à Implanter : \
\
CrewTimer et Concept2 : \
\
Import depuis OPUSS, \
Import depuis FFSU, \
Gestion des Inscriptions, \
Impressions par catégorie et classement automatique, \
Webservices GOAL (Si Possible...), \
Ajouter tous les clubs de France (Pour l'instant uniquement la LR13 est saisie.), \
Optimiser le cherche et remplace des noms de clubs grâce à un code court, via Array au lieu de ligne par ligne.

Bugs à corriger : \
Impressions des résultats CrewTimer (Bug de Colonnes, voir pour supprimer via un Array.)
