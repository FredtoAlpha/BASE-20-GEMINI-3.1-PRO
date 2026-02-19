/**
 * ===================================================================
 * APP.CORE.JS - FONCTIONS UTILITAIRES PURES
 * ===================================================================
 *
 * Module contenant les fonctions utilitaires réutilisables.
 * Toutes les fonctions sont pures (pas d'effets de bord).
 *
 * ARCHITECTURE PHASE 5 - Refactoring progressif
 * Extraction depuis Orchestration_V14I.js (Phase pilote)
 *
 * Date: 25 novembre 2025
 * Version: 1.0.0 (Phase pilote)
 * ===================================================================
 */

// ===================================================================
// UTILITAIRES DE CONSTRUCTION DE NOMS
// ===================================================================

/**
 * Construit le nom d'un onglet à partir du niveau et du suffixe
 *
 * @param {string} niveau - Le niveau (ex: "6°1", "5°2")
 * @param {string} suffix - Le suffixe (ex: "TEST", "CACHE", "FIN")
 * @returns {string} Le nom complet (ex: "6°1TEST")
 *
 * @example
 * buildSheetName_('6°1', 'TEST') // → '6°1TEST'
 * buildSheetName_('5°2', 'CACHE') // → '5°2CACHE'
 *
 * @throws {Error} Si niveau ou suffix est vide
 */
function buildSheetName_(niveau, suffix) {
  const base = String(niveau || '').trim();
  const sfx = String(suffix || '').trim();
  if (!base) throw new Error('buildSheetName_: niveau vide');
  if (!sfx) throw new Error('buildSheetName_: suffix vide');
  return base + sfx;
}

/**
 * Génère une liste de noms d'onglets à partir des niveaux et d'un suffixe
 *
 * @param {Array<string>} niveaux - Liste des niveaux (ex: ["6°1", "6°2", "5°1"])
 * @param {string} suffix - Le suffixe à ajouter (ex: "TEST", "CACHE")
 * @returns {Array<string>} Liste des noms d'onglets (ex: ["6°1TEST", "6°2TEST", "5°1TEST"])
 *
 * @example
 * makeSheetsList_(['6°1', '6°2'], 'TEST') // → ['6°1TEST', '6°2TEST']
 * makeSheetsList_(['5°1', '5°2', '5°3'], 'CACHE') // → ['5°1CACHE', '5°2CACHE', '5°3CACHE']
 */
function makeSheetsList_(niveaux, suffix) {
  return (niveaux || []).map(function(niv) { return buildSheetName_(niv, suffix); });
}

// ===================================================================
// UTILITAIRES DE MARQUAGE ET COLLECTE
// ===================================================================

/**
 * Marque un résultat de phase avec son nom
 *
 * @param {string} name - Le nom de la phase (ex: "Phase1", "Phase2")
 * @param {Object} res - Le résultat de la phase
 * @returns {Object} L'objet résultat avec le nom ajouté
 *
 * @example
 * tagPhase_('Phase1', { success: true, count: 42 })
 * // → { name: 'Phase1', success: true, count: 42 }
 */
function tagPhase_(name, res) {
  return { name, ...res };
}

/**
 * Collecte tous les warnings de toutes les phases
 *
 * @param {Array<Object>} phases - Liste des résultats de phases
 * @returns {Array<string>} Liste plate de tous les warnings
 *
 * @example
 * const phases = [
 *   { name: 'Phase1', warnings: ['Warning 1', 'Warning 2'] },
 *   { name: 'Phase2', warnings: ['Warning 3'] },
 *   { name: 'Phase3', warnings: [] }
 * ];
 * collectWarnings_(phases) // → ['Warning 1', 'Warning 2', 'Warning 3']
 */
function collectWarnings_(phases) {
  return phases.flatMap(p => p.warnings || []);
}

// ===================================================================
// UTILITAIRES GOOGLE SHEETS
// ===================================================================

/**
 * Retourne le spreadsheet actif
 * Point d'entrée unique pour éviter d'écrire dans le mauvais classeur
 *
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Le spreadsheet actif
 *
 * @example
 * const ss = getActiveSS_();
 * const sheet = ss.getSheetByName('_STRUCTURE');
 */
function getActiveSS_() {
  return SpreadsheetApp.getActive();
}

// ===================================================================
// UTILITAIRES DE BASE
// ===================================================================

/**
 * Convertit une valeur en chaîne majuscule nettoyée
 *
 * @param {*} v - Valeur à convertir
 * @returns {string} Chaîne en majuscules sans espaces de début/fin
 *
 * @example
 * _u_('  hello  ') // → 'HELLO'
 * _u_(null) // → ''
 * _u_(undefined) // → ''
 */
function _u_(v) {
  return String(v || '').trim().toUpperCase();
}

/**
 * Convertit une valeur en tableau
 *
 * @param {*} v - Valeur à convertir
 * @returns {Array} Tableau (vide si null/undefined, [v] si valeur simple)
 *
 * @example
 * _arr([1, 2, 3]) // → [1, 2, 3]
 * _arr('hello') // → ['hello']
 * _arr(null) // → []
 * _arr(undefined) // → []
 */
function _arr(v) {
  return Array.isArray(v) ? v : (v == null ? [] : [v]);
}

// ===================================================================
// UTILITAIRES DE PARSING
// ===================================================================

/**
 * Parse les codes ASSO/DISSO d'un objet élève
 *
 * @param {Object} rowObj - Objet élève avec colonnes A/D ou CODES
 * @returns {Object} { A: codeA, D: codeD } en majuscules
 *
 * @example
 * parseCodes_({ A: 'a1', D: 'd2' }) // → { A: 'A1', D: 'D2' }
 * parseCodes_({ CODES: 'A5 D3' }) // → { A: 'A5', D: 'D3' }
 */
function parseCodes_(rowObj) {
  let A = _u_(rowObj.A || rowObj.codeA || '');
  let D = _u_(rowObj.D || rowObj.codeD || '');
  const C = _u_(rowObj.CODES || '');

  if (!A && /A\d+/.test(C)) A = (C.match(/A\d+/) || [])[0];
  if (!D && /D\d+/.test(C)) D = (C.match(/D\d+/) || [])[0];

  return { A: A, D: D };
}

// ===================================================================
// UTILITAIRES DE RECHERCHE
// ===================================================================

/**
 * Trouve un élève d'un genre donné dans une liste
 * Exclut les élèves avec codes ASSO ou DISSO
 *
 * @param {Array<Object>} eleves - Liste des élèves
 * @param {string} genre - 'F' ou 'M'
 * @returns {Object|null} Premier élève trouvé ou null
 *
 * @example
 * findEleveByGenre_([{Genre: 'F'}, {Genre: 'M'}], 'F') // → {Genre: 'F'}
 */
function findEleveByGenre_(eleves, genre) {
  for (const eleve of eleves) {
    const g = eleve.Genre || eleve.Sexe || '';
    const codeA = eleve.ASSO || eleve.A || eleve['Code A'] || '';
    const codeD = eleve.DISSO || eleve.D || eleve['Code D'] || '';

    // Ignorer les élèves avec code ASSO ou DISSO
    if ((codeA && codeA !== '') || (codeD && codeD !== '')) {
      continue;
    }

    if (genre === 'F' && (g === 'F' || g === 'Fille')) {
      return eleve;
    }
    if (genre === 'M' && (g === 'M' || g === 'Garçon' || g === 'G')) {
      return eleve;
    }
  }
  return null;
}

// ===================================================================
// UTILITAIRES DE CALCUL - MÉTRIQUES DE CLASSE
// ===================================================================

/**
 * Calcule une métrique pour une classe (moyenne)
 *
 * @param {Array<Object>} eleves - Liste des élèves
 * @param {string} metric - Nom de la métrique (ex: 'COM', 'TRA')
 * @returns {number} Moyenne de la métrique
 *
 * @example
 * calculateClassMetric_([{COM: 2}, {COM: 3}, {COM: 1}], 'COM') // → 2
 */
function calculateClassMetric_LEGACY_(eleves, metric) {
  let sum = 0;

  for (const eleve of eleves) {
    const value = parseFloat(eleve[metric]) || 0;
    sum = sum + value;
  }

  return eleves.length > 0 ? sum / eleves.length : 0;
}

/**
 * Calcule l'état complet d'une classe (parité, scores)
 *
 * @param {Array<Object>} eleves - Liste des élèves de la classe
 * @returns {Object} État { size, countF, countM, deltaFM, badCOM, sumCOM, sumTRA, sumPART, sumABS }
 *
 * @example
 * computeClassState_LEGACY_([
 *   { SEXE: 'F', COM: 1, TRA: 2, PART: 0, ABS: 1 },
 *   { SEXE: 'M', COM: 2, TRA: 1, PART: 1, ABS: 0 }
 * ]) // → { size: 2, countF: 1, countM: 1, deltaFM: 0, badCOM: 1, sumCOM: 3, ... }
 */
function computeClassState_LEGACY_(eleves) {
  let countF = 0, countM = 0;
  let badCOM = 0, sumCOM = 0;
  let sumTRA = 0, sumPART = 0, sumABS = 0;

  eleves.forEach(function(e) {
    // Genre
    const genre = String(e.SEXE || e.Genre || e.Sexe || '').toUpperCase();
    if (genre === 'F') countF++;
    else if (genre === 'M') countM++;

    // Scores
    const com = parseFloat(e.COM || 0);
    const tra = parseFloat(e.TRA || 0);
    const part = parseFloat(e.PART || 0);
    const abs = parseFloat(e.ABS || 0);

    if (com === 1) badCOM++;
    sumCOM += com;
    sumTRA += tra;
    sumPART += part;
    sumABS += abs;
  });

  return {
    size: eleves.length,
    countF: countF,
    countM: countM,
    deltaFM: countF - countM,
    badCOM: badCOM,
    sumCOM: sumCOM,
    sumTRA: sumTRA,
    sumPART: sumPART,
    sumABS: sumABS
  };
}

/**
 * Simule l'état d'une classe après un swap (enlève out, ajoute in)
 *
 * @param {Object} state - État actuel de la classe
 * @param {Object} out - Élève sortant
 * @param {Object} in_ - Élève entrant
 * @returns {Object} Nouvel état après swap
 *
 * @example
 * const state = { countF: 10, countM: 12, deltaFM: -2, sumCOM: 50, ... };
 * const out = { SEXE: 'M', COM: 2, TRA: 1, PART: 0, ABS: 1 };
 * const in_ = { SEXE: 'F', COM: 1, TRA: 2, PART: 1, ABS: 0 };
 * simulateSwapState_LEGACY_(state, out, in_) // → { countF: 11, countM: 11, deltaFM: 0, ... }
 */
function simulateSwapState_LEGACY_(state, out, in_) {
  const newState = JSON.parse(JSON.stringify(state)); // Clone

  // Retirer out
  const genreOut = String(out.SEXE || out.Genre || out.Sexe || '').toUpperCase();
  if (genreOut === 'F') newState.countF--;
  else if (genreOut === 'M') newState.countM--;

  const comOut = parseFloat(out.COM || 0);
  if (comOut === 1) newState.badCOM--;
  newState.sumCOM -= comOut;
  newState.sumTRA -= parseFloat(out.TRA || 0);
  newState.sumPART -= parseFloat(out.PART || 0);
  newState.sumABS -= parseFloat(out.ABS || 0);

  // Ajouter in
  const genreIn = String(in_.SEXE || in_.Genre || in_.Sexe || '').toUpperCase();
  if (genreIn === 'F') newState.countF++;
  else if (genreIn === 'M') newState.countM++;

  const comIn = parseFloat(in_.COM || 0);
  if (comIn === 1) newState.badCOM++;
  newState.sumCOM += comIn;
  newState.sumTRA += parseFloat(in_.TRA || 0);
  newState.sumPART += parseFloat(in_.PART || 0);
  newState.sumABS += parseFloat(in_.ABS || 0);

  newState.deltaFM = newState.countF - newState.countM;

  return newState;
}

/**
 * Calcule les scores de toutes les classes pour toutes les métriques
 *
 * @param {Object} classesState - État {classe: [eleves]}
 * @param {Array<string>} metrics - Liste des métriques à calculer
 * @returns {Object} Scores {classe: {metric: value}}
 *
 * @example
 * calculateClassScores_LEGACY_(
 *   { '6°1': [{COM: 2}], '6°2': [{COM: 1}] },
 *   ['COM', 'TRA']
 * ) // → { '6°1': {COM: 2, TRA: ...}, '6°2': {COM: 1, TRA: ...} }
 */
function calculateClassScores_LEGACY_(classesState, metrics) {
  const scores = {};

  for (const [niveau, eleves] of Object.entries(classesState)) {
    scores[niveau] = {};

    for (const metric of metrics) {
      scores[niveau][metric] = calculateClassMetric_LEGACY_(eleves, metric);
    }
  }

  return scores;
}

// ===================================================================
// UTILITAIRES DE VALIDATION - LV2/OPT
// ===================================================================

/**
 * Vérifie qu'un élève reste compatible avec l'offre LV2/OPT de la classe cible
 *
 * @param {Object} eleve - L'élève à placer
 * @param {string} targetClass - Classe de destination
 * @param {Object} offers - Offres {classe: {LV2: Set, OPT: Set}}
 * @returns {boolean} true si placement OK
 *
 * @example
 * isPlacementLV2OPTOK_(
 *   { LV2: 'ITA', OPT: 'CHAV' },
 *   '6°1',
 *   { '6°1': { LV2: new Set(['ITA']), OPT: new Set(['CHAV']) } }
 * ) // → true
 */
function isPlacementLV2OPTOK_(eleve, targetClass, offers) {
  const cls = String(targetClass || '').replace(/CACHE|TEST|FIN$/,'');
  const off = offers[cls];
  if (!off) return true; // si pas d'info structure, ne pas bloquer

  const lv2 = String(eleve.lv2 || eleve.LV2 || '').toUpperCase().trim();
  const opt = String(eleve.opt || eleve.OPT || '').toUpperCase().trim();

  // LV2/OPT vides => pas de contrainte
  const lv2OK = !lv2 || off.LV2.has(lv2);
  const optOK = !opt || off.OPT.has(opt);

  return lv2OK && optOK;
}

/**
 * Vérifie si un mouvement d'élève est autorisé (garde-fou universel)
 * Vérifie: FIXE, offre LV2/OPT, quotas, mobilité
 *
 * @param {Object} eleve - L'élève à déplacer
 * @param {string} clsTo - Classe de destination
 * @param {Object} offer - Offre {classe: {LV2: [], OPT: [], quotas: {}}}
 * @param {Object} counts - Compteurs actuels {cls: {LV2:{ITA:n}, OPT:{CHAV:m}}}
 * @param {Object} quotas - Quotas attendus (optionnel)
 * @returns {boolean} true si mouvement autorisé
 */
function isMoveAllowed_(eleve, clsTo, offer, counts, quotas) {
  // 0) Élève FIXE => jamais bouge
  const fixe = String(eleve.FIXE || eleve.fixe || '').trim().toUpperCase();
  if (fixe === '1' || fixe === 'OUI' || fixe === 'X' || fixe === 'FIXE') {
    return false;
  }

  // 1) Offre LV2/OPT de la classe cible
  const off = offer[clsTo] || { LV2: [], OPT: [], quotas: {} };

  const lv2 = String(eleve.LV2 || eleve.lv2 || '').trim().toUpperCase();
  const opt = String(eleve.OPT || eleve.opt || '').trim().toUpperCase();

  // Vérifier que la LV2 est autorisée (sauf ANG qui est partout)
  if (lv2 && lv2 !== 'ANG' && off.LV2.length > 0 && off.LV2.indexOf(lv2) === -1) {
    return false;
  }

  // Vérifier que l'OPT est autorisée
  if (opt && off.OPT.length > 0 && off.OPT.indexOf(opt) === -1) {
    return false;
  }

  // 2) Respect des quotas (si définis)
  const q = off.quotas || {};
  const clsCounts = counts[clsTo] || { LV2: {}, OPT: {} };

  // Compte réalisé actuel dans la classe cible
  const realizedLV2 = lv2 ? (clsCounts.LV2[lv2] || 0) : 0;
  const realizedOPT = opt ? (clsCounts.OPT[opt] || 0) : 0;

  // Cible attendue
  const targetLV2 = lv2 ? (q[lv2] || 0) : 0;
  const targetOPT = opt ? (q[opt] || 0) : 0;

  // Si quota existe (>0), ne pas dépasser
  if (lv2 && targetLV2 > 0 && realizedLV2 >= targetLV2) {
    return false;
  }
  if (opt && targetOPT > 0 && realizedOPT >= targetOPT) {
    return false;
  }

  // 3) Mobilité : PERMUT ou LIBRE uniquement
  const mobi = String(eleve.MOBILITE || eleve.mobilite || '').trim().toUpperCase();
  if (mobi && mobi.indexOf('PERMUT') === -1 && mobi !== 'LIBRE') {
    // Si ce n'est ni PERMUT ni LIBRE, on refuse (sauf si vide = on accepte)
    if (mobi !== '') return false;
  }

  return true;
}

/**
 * Vérifie si un élève est mobile (LIBRE ou PERMUT, hors quotas, hors ASSO)
 *
 * @param {Object} eleve - L'élève à vérifier
 * @param {Object} counts - Compteurs actuels
 * @param {string} currentClass - Classe actuelle de l'élève
 * @param {Object} offer - Offre {classe: {quotas: {}}}
 * @returns {boolean} true si élève mobile
 */
function isEleveMobile_LEGACY_(eleve, counts, currentClass, offer) {
  // 1) Élève FIXE => jamais mobile
  const fixe = String(eleve.FIXE || eleve.fixe || '').trim().toUpperCase();
  if (fixe === '1' || fixe === 'OUI' || fixe === 'X' || fixe === 'FIXE') {
    return false;
  }

  // 2) Élève quota (ITA/CHAV) => FIXE pour préserver les quotas
  const lv2 = String(eleve.LV2 || eleve.lv2 || '').trim().toUpperCase();
  const opt = String(eleve.OPT || eleve.opt || '').trim().toUpperCase();

  // Si l'élève a une LV2 ou OPT avec quota dans sa classe actuelle, il est FIXE
  const classOffer = offer[currentClass] || { quotas: {} };
  const quotas = classOffer.quotas || {};

  if ((lv2 && quotas[lv2] > 0) || (opt && quotas[opt] > 0)) {
    return false; // Élève quota => FIXE
  }

  // 3) Codes ASSO => FIXE (ne pas casser les groupes)
  const codeA = String(eleve.ASSO || eleve.A || eleve.CODE_A || '').trim().toUpperCase();
  if (codeA) {
    return false; // Groupe ASSO => FIXE
  }

  // 4) Sinon => LIBRE (mobile)
  return true;
}

/**
 * @deprecated Utiliser isMoveAllowed_ à la place
 * Conservé pour compatibilité avec le code existant
 */
function eligibleForSwap_LEGACY_(eleve, clsCible, offer) {
  return isMoveAllowed_(eleve, clsCible, offer, {}, {});
}

/**
 * Vérifie si un swap est valide selon les verrous et contraintes
 *
 * @param {Object} eleve1 - Premier élève
 * @param {string} classe1 - Classe du premier élève
 * @param {Object} eleve2 - Second élève
 * @param {string} classe2 - Classe du second élève
 * @param {Object} locks - Verrous {keepDisso: boolean}
 * @param {Object} classesState - État complet des classes
 * @param {Object} offer - Offre LV2/OPT
 * @param {Object} counts - Compteurs actuels
 * @returns {boolean} true si swap valide
 */
function isSwapValid_LEGACY_(eleve1, classe1, eleve2, classe2, locks, classesState, offer, counts) {
  // ✅ Utiliser isMoveAllowed_ avec vérification des quotas
  if (!isMoveAllowed_(eleve1, classe2, offer, counts || {}, {})) {
    return false;
  }
  if (!isMoveAllowed_(eleve2, classe1, offer, counts || {}, {})) {
    return false;
  }

  // Vérifications supplémentaires selon les verrous
  if (locks.keepDisso) {
    // Vérifier que les codes D ne créent pas de conflit
    const d1 = String(eleve1.DISSO || eleve1.D || '').trim().toUpperCase();
    const d2 = String(eleve2.DISSO || eleve2.D || '').trim().toUpperCase();

    if (d1) {
      // Vérifier qu'aucun élève de classe2 n'a le même code D
      const eleves2 = classesState[classe2] || [];
      for (const e of eleves2) {
        const d = String(e.DISSO || e.D || '').trim().toUpperCase();
        if (d === d1 && e !== eleve2) return false;
      }
    }

    if (d2) {
      // Vérifier qu'aucun élève de classe1 n'a le même code D
      const eleves1 = classesState[classe1] || [];
      for (const e of eleves1) {
        const d = String(e.DISSO || e.D || '').trim().toUpperCase();
        if (d === d2 && e !== eleve1) return false;
      }
    }
  }

  return true;
}

// ===================================================================
// UTILITAIRES DE COMPTAGE
// ===================================================================

/**
 * Calcule les compteurs LV2/OPT actuels de toutes les classes
 *
 * @param {Object} classesState - État {classe: [eleves]}
 * @returns {Object} Compteurs {classe: {LV2:{ITA:n}, OPT:{CHAV:m}, total:n}}
 */
function computeCountsFromState_(classesState) {
  const counts = {};

  Object.keys(classesState).forEach(function(cls) {
    const eleves = classesState[cls] || [];
    const LV2 = {};
    const OPT = {};

    eleves.forEach(function(e) {
      const lv2 = String(e.LV2 || e.lv2 || '').trim().toUpperCase();
      const opt = String(e.OPT || e.opt || '').trim().toUpperCase();

      if (lv2 && lv2 !== 'ANG') {
        LV2[lv2] = (LV2[lv2] || 0) + 1;
      }
      if (opt) {
        OPT[opt] = (OPT[opt] || 0) + 1;
      }
    });

    counts[cls] = {
      LV2: LV2,
      OPT: OPT,
      total: eleves.length
    };
  });

  return counts;
}

/**
 * Calcule les statistiques de mobilité (LIBRE vs FIXE)
 *
 * @param {Object} classesState - État {classe: [eleves]}
 * @param {Object} offer - Offre LV2/OPT avec quotas
 * @returns {Object} Stats {libre: n, fixe: n, total: n}
 */
function computeMobilityStats_LEGACY_(classesState, offer) {
  let libre = 0;
  let fixe = 0;

  for (const [classe, eleves] of Object.entries(classesState)) {
    const counts = computeCountsFromState_(classesState);
    eleves.forEach(function(e) {
      if (isEleveMobile_LEGACY_(e, counts, classe, offer)) {
        libre++;
      } else {
        fixe++;
      }
    });
  }

  return {
    libre: libre,
    fixe: fixe,
    total: libre + fixe
  };
}

/**
 * Calcule le score d'amélioration d'un swap (hiérarchisé + pondéré)
 * Priorité 1: Parité (si hors tolérance)
 * Priorité 2: Scores pondérés (COM/TRA/PART/ABS)
 *
 * @param {Object} eleve1 - Premier élève
 * @param {string} classe1 - Classe du premier élève
 * @param {Object} eleve2 - Second élève
 * @param {string} classe2 - Classe du second élève
 * @param {Object} classesState - État complet des classes
 * @param {Object} weights - Poids {parity, com, tra, part, abs}
 * @param {number} parityTol - Tolérance de parité
 * @returns {number} Score d'amélioration (positif = amélioration)
 */
function calculateSwapScore_LEGACY_(eleve1, classe1, eleve2, classe2, classesState, weights, parityTol) {
  // Calculer l'état actuel des classes
  const state1 = computeClassState_LEGACY_(classesState[classe1]);
  const state2 = computeClassState_LEGACY_(classesState[classe2]);

  // Simuler le swap
  const state1After = simulateSwapState_LEGACY_(state1, eleve1, eleve2);
  const state2After = simulateSwapState_LEGACY_(state2, eleve2, eleve1);

  // === NIVEAU 1 : PARITÉ (prioritaire si hors tolérance) ===
  const parityBefore = Math.abs(state1.deltaFM) + Math.abs(state2.deltaFM);
  const parityAfter = Math.abs(state1After.deltaFM) + Math.abs(state2After.deltaFM);
  const parityImprovement = parityBefore - parityAfter;

  // Si une classe est hors tolérance, prioriser la parité
  const parityOutOfTol = (Math.abs(state1.deltaFM) > parityTol) || (Math.abs(state2.deltaFM) > parityTol);

  if (parityOutOfTol && parityImprovement > 0) {
    // Bonus massif pour améliorer la parité hors tolérance
    return 1000 * parityImprovement;
  }

  // === NIVEAU 2 : SCORES PONDÉRÉS (COM/TRA/PART/ABS) ===

  // Dispersion COM=1 (équilibrer les mauvais COM entre classes)
  const allClasses = Object.keys(classesState);
  const totalBadCOMBefore = allClasses.reduce(function(sum, cls) {
    return sum + computeClassState_LEGACY_(classesState[cls]).badCOM;
  }, 0);
  const meanBadCOM = totalBadCOMBefore / allClasses.length;

  const dispersionBefore = Math.abs(state1.badCOM - meanBadCOM) + Math.abs(state2.badCOM - meanBadCOM);
  const dispersionAfter = Math.abs(state1After.badCOM - meanBadCOM) + Math.abs(state2After.badCOM - meanBadCOM);
  const improvementDispersion = dispersionBefore - dispersionAfter;

  // Coût individuel pondéré
  const costBefore =
    weights.com * (state1.sumCOM + state2.sumCOM) +
    weights.tra * (state1.sumTRA + state2.sumTRA) +
    weights.part * (state1.sumPART + state2.sumPART) +
    weights.abs * (state1.sumABS + state2.sumABS);

  const costAfter =
    weights.com * (state1After.sumCOM + state2After.sumCOM) +
    weights.tra * (state1After.sumTRA + state2After.sumTRA) +
    weights.part * (state1After.sumPART + state2After.sumPART) +
    weights.abs * (state1After.sumABS + state2After.sumABS);

  const improvementCost = costBefore - costAfter;

  // Score final pondéré
  // Priorité : dispersion COM=1 (×20) > coût individuel > parité faible
  return 20 * improvementDispersion + improvementCost + 0.1 * parityImprovement;
}

// ===================================================================
// UTILITAIRES DE LOGGING
// ===================================================================

/**
 * Log avec timestamp et niveau
 *
 * @param {string} level - Niveau de log (INFO, WARN, ERROR)
 * @param {string} msg - Message à logger
 *
 * @example
 * logLine('INFO', 'Traitement terminé')
 * logLine('WARN', 'Quota non atteint')
 */
function logLine(level, msg) {
  const stamp = new Date().toLocaleString('fr-FR');
  const prefix = stamp + ' ' + level.padEnd(7);
  Logger.log(prefix + ' ' + msg);
}

// ===================================================================
// EXPORTS (Google Apps Script charge automatiquement)
// ===================================================================

/**
 * Note : Dans Google Apps Script, tous les fichiers .js sont chargés
 * automatiquement dans le scope global. Pas besoin d'export/import.
 *
 * Les fonctions définies ici sont automatiquement disponibles dans
 * tous les autres fichiers du projet.
 */
