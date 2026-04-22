/**
 * ===================================================================
 * ORCHESTRATION V14I - NOUVEAU SYSTÈME
 * ===================================================================
 * Version: 1.1.0 — SAFE: suppression des ~30 fonctions en collision
 *
 * Architecture incrémentale correcte :
 *
 * 1. Lit STRUCTURE + QUOTAS depuis l'interface Optimisation
 * 2. Exécute Phase 1 → 2 → 3 → 4 séquentiellement
 * 3. Après CHAQUE phase : écrit uniquement dans ...CACHE
 * 4. Après CHAQUE phase : affiche les onglets CACHE dans l'UI
 * 5. Phase 4 (swaps) respecte TOUS les verrous
 *
 * FONCTIONS SUPPRIMÉES (définitions canoniques dans App.*.js) :
 *
 * → App.Context.js :
 *   makeCtxFromUI_, readModeFromUI_, readNiveauxFromUI_,
 *   readQuotasFromUI_, readSourceToDestMapping_, readTargetsFromUI_,
 *   readParityToleranceFromUI_, readMaxSwapsFromUI_,
 *   buildClassOffers_, computeAllow_, buildOfferWithQuotas_
 *
 * → App.CacheManager.js :
 *   forceCacheInUIAndReload_, setInterfaceModeCACHE_,
 *   activateFirstCacheTabIfAny_, triggerUIReloadFromCACHE_,
 *   readElevesFromSelectedMode_, readElevesFromCache_, openCacheTabs_
 *
 * → App.Core.js :
 *   logLine, _u_, _arr, parseCodes_, findEleveByGenre_,
 *   isPlacementLV2OPTOK_, calculateClassScores_LEGACY_,
 *   calculateClassMetric_LEGACY_, computeCountsFromState_,
 *   isMoveAllowed_, eligibleForSwap_LEGACY_, isSwapValid_LEGACY_,
 *   computeMobilityStats_LEGACY_, isEleveMobile_LEGACY_,
 *   calculateSwapScore_LEGACY_, computeClassState_LEGACY_,
 *   simulateSwapState_LEGACY_
 *
 * FONCTIONS CONSERVÉES (uniques ou divergentes, Orchestration gagne) :
 *   makeCtxFromSourceSheets_, readClassAuthorizationsFromUI_,
 *   readQuotasFromStructure_, readTargetsFromStructure_,
 *   readElevesFromSheet_, writeAllClassesToCACHE_,
 *   announcePhaseDone_, ensureColumn_, computeMobilityFlags_,
 *   auditCacheAgainstStructure_, buildOffersFromStructure_,
 *   findBestSwap_LEGACY_, applyParityGuardrail_LEGACY_,
 *   swapEleves_, toutes les fonctions Phase*, Wrapper, etc.
 *
 * ===================================================================
 */

// ===================================================================
// FONCTION SPÉCIALE POUR PIPELINE LEGACY INITIAL
// ===================================================================

/**
 * Détecte automatiquement les onglets sources existants (ECOLE1, 6°1, etc.)
 * et crée un contexte pour le pipeline LEGACY initial (Sources → TEST)
 * @return {Object} Contexte prêt pour les 4 phases LEGACY
 */
function makeCtxFromSourceSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  // Détecter les onglets sources (formats multiples supportés)
  const sourceSheets = [];
  // ✅ Pattern universel : 6°1, ALBEXT°7, BONHOURE°2, etc. (toujours avec °)
  const sourcePattern = /^[A-Za-z0-9_-]+°\d+$/;
  // ❌ Exclure les onglets TEST, CACHE, DEF, FIN, etc.
  const excludePattern = /TEST|CACHE|DEF|FIN|SRC|SOURCE|_CONFIG|_STRUCTURE|_LOG/i;

  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (sourcePattern.test(name) && !excludePattern.test(name)) {
      sourceSheets.push(name);
    }
  }

  if (sourceSheets.length === 0) {
    throw new Error(
      '❌ Aucun onglet source trouvé !\n\n' +
      'Formats supportés:\n' +
      '• Classique: 6°1, 5°1, 4°1, 3°1, etc.\n' +
      '• ECOLE: ECOLE1, ECOLE2, etc.\n' +
      '• Personnalisé: GAMARRA°4, NOMECOLE°1, etc.'
    );
  }

  sourceSheets.sort();
  logLine('INFO', `📋 Onglets sources détectés: ${sourceSheets.join(', ')}`);

  // ✅ Lire le mapping CLASSE_ORIGINE → CLASSE_DEST depuis _STRUCTURE
  const sourceToDestMapping = readSourceToDestMapping_();

  // ✅ CORRECTION : Extraire les destinations UNIQUES depuis le MAPPING (pas seulement les sources existantes)
  // Plusieurs sources peuvent mapper vers la même destination (ex: PANASS°5, BUISSON°6, ALBEXT°7 → 6°5)
  const uniqueDestinations = [];
  const seenDest = {};
  const destToSourceMapping = {}; // Mapping inverse pour copier les en-têtes
  const sourceSheetSet = new Set(sourceSheets); // Pour vérifier l'existence rapide

  // D'abord, traiter TOUS les mappings depuis _STRUCTURE
  for (const [sourceName, dest] of Object.entries(sourceToDestMapping)) {
    if (dest && !seenDest[dest]) {
      uniqueDestinations.push(dest);
      seenDest[dest] = true;

      // Trouver la première source qui EXISTE physiquement pour cette destination
      if (!destToSourceMapping[dest]) {
        if (sourceSheetSet.has(sourceName)) {
          destToSourceMapping[dest] = sourceName;
        }
      }
    }
  }

  // Ensuite, traiter les sources détectées qui n'ont PAS de mapping
  for (const sourceName of sourceSheets) {
    if (!sourceToDestMapping[sourceName]) {
      // Pas de mapping → fallback
      let dest;
      const match = sourceName.match(/([3-6]°\d+)/);
      if (match) {
        dest = match[1];
      } else {
        const matchEcole = sourceName.match(/ECOLE(\d+)/);
        if (matchEcole) {
          dest = '6°' + matchEcole[1];
        }
      }

      if (dest && !seenDest[dest]) {
        uniqueDestinations.push(dest);
        seenDest[dest] = true;
        destToSourceMapping[dest] = sourceName;
      }
    }
  }

  // Pour les destinations sans source existante, utiliser la première source du mapping
  for (const dest of uniqueDestinations) {
    if (!destToSourceMapping[dest]) {
      // Trouver n'importe quelle source qui mappe vers cette destination
      for (const [src, d] of Object.entries(sourceToDestMapping)) {
        if (d === dest) {
          destToSourceMapping[dest] = src;
          logLine('WARN', `⚠️ Onglet ${src} introuvable, utilisé comme référence pour ${dest}TEST`);
          break;
        }
      }
    }
  }

  // Générer les noms d'onglets TEST pour les destinations uniques
  const testSheets = uniqueDestinations.map(dest => dest + 'TEST');
  const niveauxDest = uniqueDestinations;

  logLine('INFO', `📋 Onglets TEST à créer: ${testSheets.join(', ')}`);

  // Lire les quotas par classe depuis _STRUCTURE
  const quotas = readQuotasFromUI_();

  // Lire les cibles d'effectifs par classe
  const targets = readTargetsFromUI_();

  // Lire la tolérance de parité
  const tolParite = readParityToleranceFromUI_() || 2;

  // Lire le nombre max de swaps
  const maxSwaps = readMaxSwapsFromUI_() || 500;

  // Lire les autorisations de classes pour options/LV2
  const autorisations = readClassAuthorizationsFromUI_();

  return {
    ss,
    modeSrc: '',  // ✅ FIX: Mode vide pour LEGACY car les sources n'ont pas de suffixe
    writeTarget: 'TEST',
    niveaux: niveauxDest,  // ✅ FIX: Les niveaux de destination (5°1, 5°2, etc.)
    levels: niveauxDest,  // ✅ ALIAS pour compatibilité Phase4_BASEOPTI_V2
    srcSheets: sourceSheets,  // Les onglets sources réels (6°1, 6°2, etc.)
    cacheSheets: testSheets,  // Les onglets TEST à créer (5°1TEST, 5°2TEST, etc.)
    sourceToDestMapping,  // ✅ Mapping source → dest (ex: PREVERT°1 → 6°1)
    destToSourceMapping,  // ✅ Mapping inverse dest → source (ex: 6°1 → PREVERT°1)
    quotas,
    targets,
    tolParite,
    maxSwaps,
    autorisations
  };
}

// ===================================================================
// 0. UTILITAIRES DE GESTION DES ONGLETS
// ===================================================================

// ===================================================================
// ✅ FONCTIONS DÉPLACÉES DANS APP.SHEETSDATA.JS (Phase 5 - Refactoring)
// ===================================================================
// Les fonctions suivantes ont été extraites vers App.Core.js :
// - buildSheetName_(niveau, suffix)
// - makeSheetsList_(niveaux, suffix)
// - getActiveSS_()
//
// Les fonctions suivantes ont été extraites vers App.SheetsData.js :
// - getOrCreateSheet_(name)
// - getOrCreateSheetByExactName_(ss, name)
//
// Ces fonctions sont automatiquement disponibles car Google Apps Script
// charge tous les fichiers .js dans le scope global.
// ===================================================================

// writeAndVerify_() → supprimée (définition canonique dans App.SheetsData.js)

// ===================================================================
// 1. ORCHESTRATEUR PRINCIPAL
// ===================================================================

/**
 * Point d'entrée principal pour l'optimisation V14I
 * @param {Object} options - Options depuis l'interface
 * @returns {Object} Résultat complet avec statut de chaque phase
 */
function runOptimizationV14FullI(options) {
  const startTime = new Date();
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '🚀 LANCEMENT OPTIMISATION V14I - ARCHITECTURE INCRÉMENTALE');
  logLine('INFO', '='.repeat(80));

  try {
    // 1. Construire le contexte depuis _OPTI_CONFIG (Pipeline OPTI)
    const ctx = buildCtx_V2(options);
    logLine('INFO', 'Contexte OPTI créé : Mode=' + ctx.modeSrc + ', Niveaux=' + ctx.niveaux.join(',') + ', Tolérance parité=' + ctx.tolParite);
    logLine('INFO', '  📊 Max swaps: ' + ctx.maxSwaps + ', Runtime: ' + ctx.runtimeSec + 's');
    logLine('INFO', '  📊 Weights: ' + JSON.stringify(ctx.weights));

    const phasesOut = [];
    let ok = true;

    // ===== INIT V3 : VIDER CACHE ET CRÉER _BASEOPTI =====
    logLine('INFO', '\n🔧 INIT V3 : Préparation _BASEOPTI...');
    const initResult = initOptimization_V3(ctx);
    if (!initResult.ok) {
      logLine('ERROR', '❌ INIT V3 échoué');
      return { success: false, error: 'INIT V3 échoué', phases: [] };
    }
    logLine('INFO', '✅ INIT V3 terminé : ' + initResult.total + ' élèves dans _BASEOPTI');

    // ===== PHASE 1 V3 : Options & LV2 (depuis _BASEOPTI) =====
    logLine('INFO', '\n📌 PHASE 1 V3 : Affectation Options & LV2 (depuis _BASEOPTI)...');
    const p1 = Phase1I_dispatchOptionsLV2_BASEOPTI_V3(ctx);

    phasesOut.push(tagPhase_('Phase 1 V3 - Options/LV2', p1));
    announcePhaseDone_('Phase 1 V3 (Options/LV2) écrite dans _BASEOPTI + CACHE');
    forceCacheInUIAndReload_(ctx);
    ok = ok && p1.ok;
    logLine('INFO', '✅ Phase 1 V3 terminée : ' + (p1.counts ? JSON.stringify(p1.counts) : 'OK'));

    // ===== PHASE 2 V3 : DISSO/ASSO (depuis _BASEOPTI) =====
    logLine('INFO', '\n📌 PHASE 2 V3 : Application codes DISSO/ASSO (depuis _BASEOPTI)...');
    const p2 = Phase2I_applyDissoAsso_BASEOPTI_V3(ctx);
    phasesOut.push(tagPhase_('Phase 2 V3 - DISSO/ASSO', p2));
    announcePhaseDone_('Phase 2 V3 (DISSO/ASSO) écrite dans _BASEOPTI + CACHE');
    forceCacheInUIAndReload_(ctx);
    ok = ok && p2.ok;
    logLine('INFO', '✅ Phase 2 V3 terminée : DISSO=' + (p2.disso || 0) + ', ASSO=' + (p2.asso || 0));

    // ===== PHASE 3 V3 : Effectifs + Parité (depuis _BASEOPTI) =====
    logLine('INFO', '\n📌 PHASE 3 V3 : Compléter effectifs & équilibrer parité (depuis _BASEOPTI)...');
    const p3 = Phase3I_completeAndParity_BASEOPTI_V3(ctx);
    phasesOut.push(tagPhase_('Phase 3 V3 - Effectifs/Parité', p3));
    announcePhaseDone_('Phase 3 V3 (Effectifs/Parité) écrite dans _BASEOPTI + CACHE');
    forceCacheInUIAndReload_(ctx);
    ok = ok && p3.ok;
    logLine('INFO', '✅ Phase 3 V3 terminée');

    // ===== CROSS-PHASE LOOP : Phase 3 → Phase 4 avec feedback =====
    // Boucle itérative : si Phase 4 n'améliore pas assez,
    // on re-brasse la pire classe et on relance Phase 3 + Phase 4.
    const crossPhaseLoops = MULTI_RESTART_CONFIG.crossPhaseLoops;
    let p4 = null;
    let previousError = Infinity;

    for (let cpLoop = 0; cpLoop <= crossPhaseLoops; cpLoop++) {
      if (cpLoop > 0) {
        // Re-run Phase 3 : re-brasser pour donner de nouvelles cartes à Phase 4
        logLine('INFO', '\n🔄 CROSS-PHASE boucle ' + cpLoop + '/' + crossPhaseLoops + ' : relance Phase 3 + Phase 4');

        // Réinjecter les élèves de la pire classe dans le pool (désassigner)
        reshuffleWorstClass_V3_(ctx);

        const p3b = Phase3I_completeAndParity_BASEOPTI_V3(ctx);
        phasesOut.push(tagPhase_('Phase 3 V3 - Cross-Phase #' + cpLoop, p3b));
        forceCacheInUIAndReload_(ctx);
      }

      // Phase 4 : Optimisation par swaps (multi-restart intégré)
      logLine('INFO', '\n📌 PHASE 4 V3 : Optimisation par swaps' + (cpLoop > 0 ? ' (cross-phase #' + cpLoop + ')' : '') + '...');
      p4 = Phase4_balanceScoresSwaps_BASEOPTI_V3(ctx);

      const currentError = p4.finalError || Infinity;
      const improvement = previousError > 0 ? (previousError - currentError) / previousError : 0;

      logLine('INFO', '✅ Phase 4 V3 : ' + (p4.swapsApplied || 0) + ' swaps, erreur=' + (currentError === Infinity ? '?' : currentError.toFixed(2)) + ', amélioration=' + (improvement * 100).toFixed(1) + '%');

      if (cpLoop > 0 && improvement < MULTI_RESTART_CONFIG.minImprovementPct) {
        logLine('INFO', '  🛑 Amélioration insuffisante (' + (improvement * 100).toFixed(1) + '% < ' + (MULTI_RESTART_CONFIG.minImprovementPct * 100).toFixed(1) + '%), arrêt cross-phase.');
        break;
      }
      previousError = currentError;
    }

    phasesOut.push(tagPhase_('Phase 4 V3 - Swaps', p4));
    announcePhaseDone_('Phase 4 V3 terminée : ' + (p4.swapsApplied || 0) + ' swaps appliqués. Résultat dans _BASEOPTI + CACHE');
    forceCacheInUIAndReload_(ctx);
    ok = ok && (p4.ok !== false);

    // Basculer l'interface en mode CACHE
    setInterfaceModeCACHE_(ctx);

    const endTime = new Date();
    const durationSec = (endTime - startTime) / 1000;
    const durationLog = durationSec.toFixed(2);

    logLine('INFO', '='.repeat(80));
    logLine('INFO', '✅ OPTIMISATION V14I (PIPELINE OPTI V3) TERMINÉE EN ' + durationLog + 's');
    logLine('INFO', 'Swaps totaux : ' + (p4.swapsApplied || 0));
    logLine('INFO', 'Architecture : _BASEOPTI + _OPTI_CONFIG');
    logLine('INFO', '='.repeat(80));

    // ✅ FORCER L'OUVERTURE DES ONGLETS CACHE AVEC FLUSH STRICT
    logLine('INFO', '📂 Ouverture des onglets CACHE...');
    const openedInfo = openCacheTabs_(ctx);

    // ✅ AUDIT FINAL : Vérifier conformité CACHE vs STRUCTURE
    const cacheAudit = auditCacheAgainstStructure_(ctx);

    // ✅ FINALISATION : Calcul moyennes et mise en forme onglets TEST
    try {
      finalizeTestSheets_(ctx);
    } catch (e) {
      logLine('WARN', '⚠️ Erreur lors de la finalisation des onglets TEST : ' + e.message);
    }

    // ✅ Réponse 100% sérialisable et compatible avec l'UI
    const warningsOut = (collectWarnings_(phasesOut) || []).map(function(w) {
      return String(w);
    });

    const response = {
      success: ok,                              // Contrat UI attend "success"
      ok: ok,                                   // Compatibilité legacy
      nbSwaps: p4.swapsApplied || 0,           // Contrat UI attend "nbSwaps"
      swaps: p4.swapsApplied || 0,             // Compatibilité legacy
      tempsTotalMs: Math.round(durationSec * 1000), // Contrat UI attend "tempsTotalMs"
      durationMs: Math.round(durationSec * 1000),   // Alias explicite pour compatibilité
      duration: durationSec,                    // Durée en secondes (nombre)
      durationSec: durationSec,                 // Alias explicite pour analyse côté client
      warnings: warningsOut,                    // Forcer String[]
      writeSuffix: 'CACHE',
      sourceSuffix: ctx.modeSrc || 'TEST',
      cacheSheets: ctx.cacheSheets.slice(),    // ✅ Liste des onglets CACHE pour l'UI
      quotasLus: ctx.quotas || {},             // ✅ Diagnostic : quotas détectés
      cacheStats: openedInfo.stats || [],      // ✅ Stats détaillées : lignes/colonnes par onglet
      cacheAudit: cacheAudit || {},            // ✅ Audit de conformité par classe
      openedInfo: {                             // ✅ Info sur les onglets activés
        opened: openedInfo.opened || [],
        active: openedInfo.active || null,
        error: openedInfo.error || null
      },
      // ⚠️ NE PAS inclure phasesOut (peut contenir des objets Apps Script)
      message: ok ? 'Optimisation réussie' : 'Optimisation terminée avec warnings'
    };

    // ✅ Garantir la sérialisation JSON (purge undefined, fonctions, objets Apps Script)
    return JSON.parse(JSON.stringify(response));

  } catch (e) {
    logLine('ERROR', '❌ ERREUR FATALE ORCHESTRATION V14I : ' + e.message);
    logLine('ERROR', e.stack);
    throw e;
  }
}

// ===================================================================
// ✅ FONCTIONS DÉPLACÉES DANS APP.CORE.JS (Phase 5 - Refactoring)
// ===================================================================
// - tagPhase_(name, res)
// - collectWarnings_(phases)
// ===================================================================

// ===================================================================
// 2. CONSTRUCTION DU CONTEXTE DEPUIS L'INTERFACE
// ===================================================================

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.Context.js :
//  - makeCtxFromUI_()
//  - readModeFromUI_()
//  - readNiveauxFromUI_()
//  - readQuotasFromUI_()
// ===================================================================

// readQuotasFromStructure_() → supprimée (définition canonique dans App.SheetsData.js)

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.Context.js :
//  - readSourceToDestMapping_()
//  - readTargetsFromUI_()
// ===================================================================

// readTargetsFromStructure_() → supprimée (définition canonique dans App.SheetsData.js)

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.Context.js :
//  - readParityToleranceFromUI_()
//  - readMaxSwapsFromUI_()
// ===================================================================

/**
 * @deprecated Cette fonction est obsolète. Utiliser buildCtx_V2() à la place.
 * @see buildCtx_V2() dans BASEOPTI_Architecture_V3.gs
 * 
 * Lit les autorisations de classes par option (legacy).
 * Format : { ITA: ["6°1", "6°3"], CHAV: ["6°2", "6°3"], ... }
 * Retourne des valeurs codées en dur.
 * 
 * ⚠️ LEGACY : Cette fonction ne lit plus l'interface réelle.
 * Les autorisations sont maintenant calculées depuis _OPTI_CONFIG (colonnes ITA, CHAV, etc.).
 */
function readClassAuthorizationsFromUI_() {
  // ⚠️ LEGACY : Valeurs codées en dur
  // Les autorisations sont maintenant calculées depuis _OPTI_CONFIG
  return {
    ITA: ["6°1"],
    CHAV: ["6°3"],
    ESP: ["6°1", "6°2", "6°3", "6°4", "6°5"]
  };
}

// ===================================================================
// 3. UI : FORCER L'AFFICHAGE DES ONGLETS CACHE
// ===================================================================

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.CacheManager.js :
//  - forceCacheInUIAndReload_()
//  - setInterfaceModeCACHE_()
//  - activateFirstCacheTabIfAny_()
//  - triggerUIReloadFromCACHE_()
// ===================================================================

// announcePhaseDone_() → supprimée (définition canonique dans App.UIBridge.js)

// ===================================================================
// 4. LECTURE / ÉCRITURE DES DONNÉES
// ===================================================================

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.CacheManager.js :
//  - readElevesFromSelectedMode_()
//  - readElevesFromCache_()
// ===================================================================

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.SheetsData.js :
//  - readElevesFromSheet_()
//  - writeAllClassesToCACHE_()
//  - clearSheets_()
//  - writeElevesToSheet_()
//  - writeToCache_()
// ===================================================================

// ===================================================================
// 5. PHASE 1I : AFFECTATION OPTIONS & LV2
// ===================================================================


/**
 * Calcul & écriture des colonnes FIXE/MOBILITE dans tous les ...CACHE
 */
function computeMobilityFlags_(ctx) {
  logLine('INFO', '🔍 Calcul des statuts de mobilité (FIXE/PERMUT/LIBRE)...');

  const ss = ctx.ss;
  const classOffers = buildClassOffers_(ctx); // "6°1" -> {LV2:Set, OPT:Set}

  logLine('INFO', '  Classes offrant LV2/OPT: ' + JSON.stringify(
    Object.keys(classOffers).reduce(function(acc, cl) {
      acc[cl] = {
        LV2: Array.from(classOffers[cl].LV2),
        OPT: Array.from(classOffers[cl].OPT)
      };
      return acc;
    }, {})
  ));

  // 1) Lire tout le CACHE en mémoire + construire groupes A + index D
  const studentsByClass = {}; // "6°1" -> [{row, data, id, ...}]
  const groupsA = {};         // "A7" -> [{class,nameRow,indexRow,...}]
  const Dindex = {};          // "6°1" -> Set(Dx déjà présents)

  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const base = cacheName.replace(/CACHE$/, '');
    const sh = ss.getSheetByName(cacheName);
    if (!sh) return;

    const lr = Math.max(sh.getLastRow(), 1);
    const lc = Math.max(sh.getLastColumn(), 1);
    const values = sh.getRange(1, 1, lr, lc).getDisplayValues();
    const headers = values[0];
    const find = function(name) { return headers.indexOf(name); };

    // Assure colonnes FIXE & MOBILITE
    const colFIXE = ensureColumn_(sh, 'FIXE');
    const colMOB = ensureColumn_(sh, 'MOBILITE');

    // indices de colonnes utiles
    const idxNom = find('NOM');
    const idxPrenom = find('PRENOM');
    const idxSexe = find('SEXE');
    const idxLV2 = find('LV2');
    const idxOPT = find('OPT');
    const idxA = find('A');
    const idxD = find('D');
    const idxCodes = find('CODES');
    const idxAsso = find('ASSO');
    const idxDisso = find('DISSO');

    studentsByClass[base] = [];
    Dindex[base] = new Set();

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const obj = {
        NOM: row[idxNom] || '',
        PRENOM: row[idxPrenom] || '',
        SEXE: row[idxSexe] || '',
        LV2: row[idxLV2] || '',
        OPT: row[idxOPT] || '',
        A: (idxA >= 0 ? row[idxA] : (idxAsso >= 0 ? row[idxAsso] : '')),
        D: (idxD >= 0 ? row[idxD] : (idxDisso >= 0 ? row[idxDisso] : '')),
        CODES: (idxCodes >= 0 ? row[idxCodes] : '')
      };

      const codes = parseCodes_(obj);
      const id = _u_((obj.NOM || '') + '|' + (obj.PRENOM || '') + '|' + base);
      const st = {
        id: id,
        classe: base,
        rowIndex: r + 1,
        data: obj,
        A: codes.A,
        D: codes.D,
        colFIXE: colFIXE,
        colMOB: colMOB,
        sheet: sh
      };

      studentsByClass[base].push(st);

      if (codes.A) {
        if (!groupsA[codes.A]) groupsA[codes.A] = [];
        groupsA[codes.A].push(st);
      }
      if (codes.D) {
        Dindex[base].add(codes.D);
      }
    }
  });

  logLine('INFO', '  Groupes A détectés: ' + Object.keys(groupsA).length);
  logLine('INFO', '  Codes D détectés: ' + JSON.stringify(
    Object.keys(Dindex).reduce(function(acc, cl) {
      acc[cl] = Array.from(Dindex[cl]);
      return acc;
    }, {})
  ));

  // 2) Déterminer FIXE explicite & compute Allow individuels
  const explicitFixed = new Set();
  Object.keys(studentsByClass).forEach(function(cl) {
    studentsByClass[cl].forEach(function(st) {
      const vFIXE = _u_(st.sheet.getRange(st.rowIndex, st.colFIXE + 1).getDisplayValue());
      if (vFIXE === 'FIXE' || vFIXE === 'SPEC' || vFIXE === 'LOCK') {
        explicitFixed.add(st.id);
      }
      st.allow = computeAllow_(st.data, classOffers);
    });
  });

  logLine('INFO', '  Élèves FIXE explicites: ' + explicitFixed.size);

  // 3) Résoudre groupes A
  const groupAllow = {};
  Object.keys(groupsA).forEach(function(codeA) {
    const members = groupsA[codeA];
    let inter = null;
    let anyFixed = false;
    let fixedClass = null;

    members.forEach(function(st) {
      if (explicitFixed.has(st.id)) {
        anyFixed = true;
        fixedClass = st.classe;
      }
      const set = new Set(st.allow);
      inter = (inter === null) ? set : new Set([...inter].filter(function(x) { return set.has(x); }));
    });

    const allowArr = inter ? Array.from(inter) : [];
    let status = null;
    let pin = null;

    if (anyFixed) {
      if (allowArr.includes(fixedClass)) {
        status = 'FIXE';
        pin = fixedClass;
      } else {
        status = 'CONFLIT';
      }
    } else {
      if (allowArr.length === 0) status = 'CONFLIT';
      else if (allowArr.length === 1) {
        status = 'FIXE';
        pin = allowArr[0];
      } else {
        status = 'PERMUT';
      }
    }

    groupAllow[codeA] = { allow: new Set(allowArr), status: status, pin: pin };
  });

  // 4) Statut individuel final
  function statusForStudent(st) {
    // a) FIXE explicite
    if (explicitFixed.has(st.id)) return { fix: true, mob: 'FIXE' };

    // b) groupe A
    if (st.A && groupAllow[st.A]) {
      const g = groupAllow[st.A];
      if (g.status === 'CONFLIT') return { fix: false, mob: 'CONFLIT(A)' };
      if (g.status === 'FIXE') return { fix: true, mob: 'GROUPE_FIXE(' + st.A + '→' + g.pin + ')' };
      if (g.status === 'PERMUT') return { fix: false, mob: 'GROUPE_PERMUT(' + st.A + '→' + Array.from(g.allow).join('/') + ')' };
    }

    // c) LV2+OPT individuellement
    let allow = st.allow.slice();

    // d) filtre D
    if (st.D) {
      allow = allow.filter(function(c) { return !Dindex[c].has(st.D) || c === st.classe; });
    }

    if (allow.length === 0) return { fix: false, mob: 'CONFLIT(LV2/OPT/D)' };
    if (allow.length === 1) return { fix: true, mob: 'FIXE' };
    if (allow.length === 2) return { fix: false, mob: 'PERMUT(' + allow.join(',') + ')' };

    return { fix: false, mob: 'LIBRE' };
  }

  // 5) Écrire en feuille
  let countFIXE = 0;
  let countPERMUT = 0;
  let countLIBRE = 0;
  let countCONFLIT = 0;

  Object.keys(studentsByClass).forEach(function(cl) {
    const arr = studentsByClass[cl];
    arr.forEach(function(st) {
      const s = statusForStudent(st);

      if (s.fix) {
        st.sheet.getRange(st.rowIndex, st.colFIXE + 1).setValue('FIXE');
        countFIXE++;
      } else {
        st.sheet.getRange(st.rowIndex, st.colFIXE + 1).clearContent();
      }

      st.sheet.getRange(st.rowIndex, st.colMOB + 1).setValue(s.mob);

      if (s.mob.includes('PERMUT')) countPERMUT++;
      else if (s.mob === 'LIBRE') countLIBRE++;
      else if (s.mob.includes('CONFLIT')) countCONFLIT++;
    });
  });

  SpreadsheetApp.flush();

  logLine('INFO', '✅ Mobilité calculée: FIXE=' + countFIXE + ', PERMUT=' + countPERMUT + ', LIBRE=' + countLIBRE + ', CONFLIT=' + countCONFLIT);
}

// openCacheTabs_() → supprimée (définition canonique dans App.CacheManager.js)

// ===================================================================
// 9B. AUDIT CACHE CONTRE STRUCTURE
// ===================================================================

// buildOfferWithQuotas_() → supprimée (définition canonique dans App.Context.js)

/**
 * Audite les onglets CACHE contre la structure attendue
 * Retourne un objet { classe: { total, F, M, LV2:{}, OPT:{}, violations:{} } }
 */
function auditCacheAgainstStructure_(ctx) {
  logLine('INFO', '\n🔍 AUDIT: Vérification conformité CACHE vs STRUCTURE...');
  
  const offer = buildOfferWithQuotas_(ctx);
  const audit = {};
  
  // Pour chaque onglet CACHE
  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const cls = cacheName.replace(/CACHE$/, '').trim();
    const sh = ctx.ss.getSheetByName(cacheName);
    
    if (!sh) {
      logLine('WARN', '  ⚠️ Onglet ' + cacheName + ' introuvable');
      return;
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      audit[cls] = { total: 0, F: 0, M: 0, LV2: {}, OPT: {}, FIXE: 0, PERMUT: 0, LIBRE: 0, violations: { LV2: [], OPT: [], D: [], A: [], QUOTAS: [] } };
      return;
    }
    
    const headers = data[0];
    const idxSexe = headers.indexOf('SEXE') || headers.indexOf('Genre');
    const idxLV2 = headers.indexOf('LV2');
    const idxOPT = headers.indexOf('OPT');
    const idxDisso = headers.indexOf('DISSO') || headers.indexOf('D');
    const idxAsso = headers.indexOf('ASSO') || headers.indexOf('A');
    const idxFixe = headers.indexOf('FIXE');
    const idxMob = headers.indexOf('MOBILITE');
    
    // Agrégation
    const agg = {
      total: 0,
      F: 0,
      M: 0,
      LV2: {},
      OPT: {},
      FIXE: 0,
      PERMUT: 0,
      LIBRE: 0,
      violations: {
        LV2: [],
        OPT: [],
        D: [],
        A: [],
        QUOTAS: []
      }
    };
    
    const codesD = new Set();
    const codesA = {};
    
    // Parcourir les élèves
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      agg.total++;
      
      // Sexe
      const sexe = String(row[idxSexe] || '').trim().toUpperCase();
      if (sexe === 'F' || sexe === 'FILLE') agg.F++;
      else if (sexe === 'M' || sexe === 'G' || sexe === 'GARÇON') agg.M++;
      
      // LV2
      if (idxLV2 >= 0) {
        const lv2 = String(row[idxLV2] || '').trim().toUpperCase();
        if (lv2 && lv2 !== 'ANG') {
          agg.LV2[lv2] = (agg.LV2[lv2] || 0) + 1;
        }
      }
      
      // OPT
      if (idxOPT >= 0) {
        const opt = String(row[idxOPT] || '').trim().toUpperCase();
        if (opt) {
          agg.OPT[opt] = (agg.OPT[opt] || 0) + 1;
        }
      }
      
      // FIXE/MOBILITE (corrigé pour compter tous les élèves)
      let estFixe = false;
      let estPermut = false;
      
      if (idxFixe >= 0) {
        const fixe = String(row[idxFixe] || '').trim().toUpperCase();
        if (fixe === 'FIXE' || fixe === 'X') {
          agg.FIXE++;
          estFixe = true;
        }
      }
      
      if (!estFixe && idxMob >= 0) {
        const mob = String(row[idxMob] || '').trim().toUpperCase();
        if (mob.indexOf('PERMUT') >= 0 || mob === 'PERMUT') {
          agg.PERMUT++;
          estPermut = true;
        } else if (mob === 'FIXE') {
          agg.FIXE++;
          estFixe = true;
        }
      }
      
      // Si ni FIXE ni PERMUT, c'est LIBRE par défaut
      if (!estFixe && !estPermut) {
        agg.LIBRE++;
      }
      
      // Codes D
      if (idxDisso >= 0) {
        const d = String(row[idxDisso] || '').trim().toUpperCase();
        if (d) {
          if (codesD.has(d)) {
            agg.violations.D.push('Code D=' + d + ' en double');
          }
          codesD.add(d);
        }
      }
      
      // Codes A
      if (idxAsso >= 0) {
        const a = String(row[idxAsso] || '').trim().toUpperCase();
        if (a) {
          if (!codesA[a]) codesA[a] = [];
          codesA[a].push(i);
        }
      }
    }
    
    // Vérifier violations LV2
    const offLV2 = (offer[cls] && offer[cls].LV2) ? offer[cls].LV2 : [];
    Object.keys(agg.LV2).forEach(function(lv2) {
      if (offLV2.length > 0 && offLV2.indexOf(lv2) === -1) {
        agg.violations.LV2.push(lv2 + ' non autorisée (' + agg.LV2[lv2] + ' élèves)');
      }
    });
    
    // Vérifier violations OPT
    const offOPT = (offer[cls] && offer[cls].OPT) ? offer[cls].OPT : [];
    Object.keys(agg.OPT).forEach(function(opt) {
      if (offOPT.length > 0 && offOPT.indexOf(opt) === -1) {
        agg.violations.OPT.push(opt + ' non autorisée (' + agg.OPT[opt] + ' élèves)');
      }
    });
    
    // Vérifier violations A (groupes éclatés)
    Object.keys(codesA).forEach(function(a) {
      if (codesA[a].length < 2) {
        agg.violations.A.push('Groupe A=' + a + ' incomplet (1 seul élève)');
      }
    });
    
    // ⚖️ Vérification quotas par classe (si présents)
    const q = (offer[cls] && offer[cls].quotas) ? offer[cls].quotas : {};
    const quotaViol = [];
    Object.keys(q).forEach(function(key) {
      const K = key.toUpperCase();
      const target = q[K]; // quota attendu
      // Où chercher le réalisé ?
      const realized =
        (agg.LV2[K] !== undefined ? agg.LV2[K] : 0) +
        (agg.OPT[K] !== undefined ? agg.OPT[K] : 0);
      
      if (target > 0 && realized !== target) {
        quotaViol.push(K + ': attendu=' + target + ', réalisé=' + realized);
      }
    });
    
    if (quotaViol.length) {
      agg.violations.QUOTAS = quotaViol;
    } else {
      agg.violations.QUOTAS = [];
    }
    
    audit[cls] = agg;
    
    // Log par classe
    logLine('INFO', '📦 Classe ' + cls + ' — Total=' + agg.total + ', F=' + agg.F + ', M=' + agg.M + ', |F-M|=' + Math.abs(agg.F - agg.M));
    logLine('INFO', '   Offre attendue: LV2=[' + offLV2.join(',') + '], OPT=[' + offOPT.join(',') + ']');
    logLine('INFO', '   LV2 réalisées: ' + JSON.stringify(agg.LV2));
    logLine('INFO', '   OPT réalisées: ' + JSON.stringify(agg.OPT));
    logLine('INFO', '   Mobilité: FIXE=' + agg.FIXE + ', PERMUT=' + agg.PERMUT + ', LIBRE=' + agg.LIBRE);
    
    if (agg.violations.LV2.length) {
      logLine('WARN', '   ❌ Violations LV2 (' + agg.violations.LV2.length + '): ' + agg.violations.LV2.join(' | '));
    }
    if (agg.violations.OPT.length) {
      logLine('WARN', '   ❌ Violations OPT (' + agg.violations.OPT.length + '): ' + agg.violations.OPT.join(' | '));
    }
    if (agg.violations.D.length) {
      logLine('WARN', '   ❌ Violations D (' + agg.violations.D.length + '): ' + agg.violations.D.join(' | '));
    }
    if (agg.violations.A.length) {
      logLine('WARN', '   ❌ Violations A (' + agg.violations.A.length + '): ' + agg.violations.A.join(' | '));
    }
    if (agg.violations.QUOTAS && agg.violations.QUOTAS.length) {
      logLine('WARN', '   ❌ Violations QUOTAS (' + agg.violations.QUOTAS.length + '): ' + agg.violations.QUOTAS.join(' | '));
    }
  });
  
  logLine('INFO', '✅ Audit terminé pour ' + Object.keys(audit).length + ' classes');
  return audit;
}
