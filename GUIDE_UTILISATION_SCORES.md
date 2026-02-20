# Console Scores Eleves -- Guide d'utilisation

## Installation (une seule fois)

1. Créer un nouveau Google Sheet (ou ouvrir celui de ton app de répartition)
2. Aller dans Extensions → Apps Script
3. Effacer le contenu par défaut (function myFunction()...)
4. Coller tout le contenu du fichier ScoresEleves.gs
5. Sauvegarder (Ctrl+S)
6. Fermer l'éditeur Apps Script
7. Recharger la page du Google Sheet
8. Un menu 📊 Scores Élèves apparaît dans la barre de menu

## Première utilisation

1. Cliquer sur 📊 Scores Élèves → 🏗️ Initialiser le classeur
2. Autoriser le script quand Google le demande
3. Les 6 onglets sont créés automatiquement :
   * DATA_ABS — pour les absences
   * DATA_INCIDENTS — pour les incidents/sanctions
   * DATA_PUNITIONS — pour les punitions
   * DATA_NOTES — pour les notes
   * PARAMÈTRES — seuils modifiables
   * SYNTHÈSE — résultat final

## Coller les exports Pronote

Pour chaque module, depuis Pronote :

| Onglet cible | Export Pronote | Ce qu'il faut coller |
|---|---|---|
| DATA_ABS | Absences | Tout le tableau avec les 2 lignes d'en-tête |
| DATA_INCIDENTS | Incidents | Tout le tableau avec les 2 lignes d'en-tête |
| DATA_PUNITIONS | Punitions | Tout le tableau avec la ligne d'en-tête |
| DATA_NOTES | Notes/Moyennes | Tout le tableau avec les 2 lignes d'en-tête |

**Important** : bien coller depuis la cellule A1 de chaque onglet.

## Lancer les calculs

**Option 1 : Module par module**

* 📊 Scores Élèves → 📋 Calculer ABS — calcule le score d'assiduité
* 📊 Scores Élèves → 🚨 Calculer COM — calcule le score de comportement
* 📊 Scores Élèves → 📚 Calculer TRA — calcule le score de travail
* 📊 Scores Élèves → 🗣️ Calculer PART — calcule le score de participation

**Option 2 : Tout d'un coup**

* 📊 Scores Élèves → 🎯 Calculer TOUS les scores

## Modifier les seuils

Dans l'onglet PARAMÈTRES, les cellules en jaune/bleu sont modifiables.

### Score ABS (Absences)

* Formule : Score DJ ×0.6 + Score NJ ×0.4, arrondi supérieur
* Score DJ basé sur les demi-journées bulletin
* Score NJ basé sur les absences non justifiées

### Score COM (Comportement)

* Formule : Punitions ×1 + Incidents (pts gravité) ×3
* La gravité de chaque incident (1/5 à 5/5) donne de 1 à 5 points, multiplié par 3

### Score TRA (Travail)

* Moyenne pondérée par volume horaire
* Les coefficients sont modifiables dans PARAMÈTRES

### Score PART (Participation)

* Moyenne de l'oral Anglais + oral LV2

## Mettre à jour en cours d'année

1. Coller les nouveaux exports dans les onglets DATA (en écrasant les anciens)
2. Relancer les calculs via le menu
3. La synthèse se met à jour automatiquement

## Résultat

L'onglet SYNTHÈSE contient le tableau final :

| Nom | Classe | ABS | COM | TRA | PART |
|---|---|---|---|---|---|
| ADAM Léa | 4E 1 | 3 | 4 | 2 | 3 |
| ... | ... | ... | ... | ... | ... |

Trié par classe puis ordre alphabétique, prêt à être utilisé par l'app de répartition.
