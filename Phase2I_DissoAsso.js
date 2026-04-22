/**
 * ===================================================================
 * PHASE 2I : HELPERS PARTAGÉS (lockAttributes_, moveEleveToClass_)
 * ===================================================================
 * Les fonctions Phase2I_applyDissoAsso_ / applyDisso_ / applyAsso_ /
 * propagateAssoConstraints_ / findClasseWithoutCode_ ont été retirées
 * (remplacées par Phase2I_applyDissoAsso_BASEOPTI_V3 dans
 * Phases_BASEOPTI_V3_COMPLETE.js et Phase2I_applyDissoAsso_LEGACY dans
 * LEGACY_Phase2_DissoAsso.js).
 *
 * Ce fichier conserve uniquement les 2 helpers utilisés par Orchestration_V14I.js.
 */

/**
 * Verrouille certains attributs pour éviter qu'ils soient modifiés
 */
function lockAttributes_(classesState, locks) {
  for (const [niveau, eleves] of Object.entries(classesState)) {
    for (const eleve of eleves) {
      if (!eleve._locks) {
        eleve._locks = {};
      }

      if (locks.options) {
        eleve._locks.ITA = true;
        eleve._locks.CHAV = true;
      }
      if (locks.lv2) {
        eleve._locks.LV2 = true;
      }
      if (locks.disso) {
        eleve._locks.DISSO = true;
      }
      if (locks.asso) {
        eleve._locks.ASSO = true;
      }
      if (locks.parity) {
        eleve._locks.PARITY = true;
      }
    }
  }
}

/**
 * Déplace un élève d'une classe à une autre
 */
function moveEleveToClass_(classesState, eleve, fromClasse, toClasse) {
  // Retirer de la classe source
  const fromEleves = classesState[fromClasse];
  const index = fromEleves.indexOf(eleve);
  if (index > -1) {
    fromEleves.splice(index, 1);
  }

  // Ajouter à la classe cible
  if (!classesState[toClasse]) {
    classesState[toClasse] = [];
  }
  classesState[toClasse].push(eleve);

  // Mettre à jour la propriété Classe de l'élève
  eleve.Classe = toClasse;
}
