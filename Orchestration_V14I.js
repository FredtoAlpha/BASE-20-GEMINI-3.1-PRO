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
 * Phase 1I : Affecte les options et LV2 selon les quotas UI
 * LIT : onglet sélectionné
 * ÉCRIT : uniquement CACHE
 */
function Phase1I_dispatchOptionsLV2_(ctx) {
  const warnings = [];

  // ✅ OPTIMISATION : Lire depuis l'onglet CONSOLIDATION (regroupe tous les élèves)
  // Plus simple et plus rapide qu'une lecture multi-sources
  const consolidationSheet = ctx.ss.getSheetByName('CONSOLIDATION');
  if (!consolidationSheet) {
    logLine('ERROR', 'Onglet CONSOLIDATION introuvable ! Impossible de continuer.');
    return {
      ok: false,
      warnings: ['CONSOLIDATION introuvable'],
      counts: {}
    };
  }

  const poolGlobal = readElevesFromSheet_(consolidationSheet);
  logLine('INFO', 'Phase 1 : Pool global de ' + poolGlobal.length + ' élèves (depuis CONSOLIDATION)');

  // Créer les classes destination VIDES
  const classesState = {};
  for (const niveau of ctx.niveaux) {
    classesState[niveau] = [];
  }

  // Dispatcher les élèves selon quotas LV2/OPT
  const { stats, warn } = dispatchElevesWithQuotas_(
    poolGlobal,
    classesState,
    ctx.quotas,
    ctx.niveaux
  );

  warnings.push(...(warn || []));

  // Écrire dans CACHE
  writeAllClassesToCACHE_(ctx, classesState);

  return {
    ok: true,
    warnings,
    counts: stats
  };
}

/**
 * ✅ NOUVELLE LOGIQUE : Dispatcher les élèves depuis le pool global vers les classes
 * selon les quotas LV2/OPT définis dans _STRUCTURE
 *
 * RÈGLE : Non-ESP prioritaires (comme dit par le user)
 */
function dispatchElevesWithQuotas_(poolGlobal, classesState, quotas, niveaux) {
  const warn = [];
  const stats = {
    ITA: 0,
    CHAV: 0,
    LATIN: 0,
    ESP: 0,
    ALL: 0,
    PT: 0
  };

  // Marquer tous les élèves comme non assignés
  poolGlobal.forEach(function(e) { e._ASSIGNED = false; });

  // Pour chaque classe, dispatcher les élèves selon les quotas
  for (const niveau of niveaux) {
    const quota = quotas[niveau] || {};

    // 1. Dispatcher ITA (LV2)
    if (quota.ITA && quota.ITA > 0) {
      const elevesITA = poolGlobal.filter(function(e) {
        return !e._ASSIGNED && String(e.LV2 || '').trim().toUpperCase() === 'ITA';
      });

      const assigned = Math.min(elevesITA.length, quota.ITA);
      for (let i = 0; i < assigned; i++) {
        elevesITA[i]._ASSIGNED = true;
        classesState[niveau].push(elevesITA[i]);
      }
      stats.ITA += assigned;

      if (assigned < quota.ITA) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' ITA affectés sur ' + quota.ITA + ' demandés');
      }
      logLine('INFO', '  ✅ ' + niveau + ' : ' + assigned + ' élèves ITA placés');
    }

    // 2. Dispatcher CHAV (OPT)
    if (quota.CHAV && quota.CHAV > 0) {
      const elevesCHAV = poolGlobal.filter(function(e) {
        return !e._ASSIGNED && String(e.OPT || '').trim().toUpperCase() === 'CHAV';
      });

      const assigned = Math.min(elevesCHAV.length, quota.CHAV);
      for (let i = 0; i < assigned; i++) {
        elevesCHAV[i]._ASSIGNED = true;
        classesState[niveau].push(elevesCHAV[i]);
      }
      stats.CHAV += assigned;

      if (assigned < quota.CHAV) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' CHAV affectés sur ' + quota.CHAV + ' demandés');
      }
      logLine('INFO', '  ✅ ' + niveau + ' : ' + assigned + ' élèves CHAV placés');
    }

    // 3. Dispatcher LATIN (OPT)
    if (quota.LATIN && quota.LATIN > 0) {
      const elevesLATIN = poolGlobal.filter(function(e) {
        return !e._ASSIGNED && String(e.OPT || '').trim().toUpperCase() === 'LATIN';
      });

      const assigned = Math.min(elevesLATIN.length, quota.LATIN);
      for (let i = 0; i < assigned; i++) {
        elevesLATIN[i]._ASSIGNED = true;
        classesState[niveau].push(elevesLATIN[i]);
      }
      stats.LATIN += assigned;

      if (assigned < quota.LATIN) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' LATIN affectés sur ' + quota.LATIN + ' demandés');
      }
      logLine('INFO', '  ✅ ' + niveau + ' : ' + assigned + ' élèves LATIN placés');
    }

    // 4. Dispatcher ESP (LV2) si quota défini
    if (quota.ESP && quota.ESP > 0) {
      const elevesESP = poolGlobal.filter(function(e) {
        return !e._ASSIGNED && String(e.LV2 || '').trim().toUpperCase() === 'ESP';
      });

      const assigned = Math.min(elevesESP.length, quota.ESP);
      for (let i = 0; i < assigned; i++) {
        elevesESP[i]._ASSIGNED = true;
        classesState[niveau].push(elevesESP[i]);
      }
      stats.ESP += assigned;
      logLine('INFO', '  ✅ ' + niveau + ' : ' + assigned + ' élèves ESP placés');
    }

    // 5. Dispatcher ALL (LV2) si quota défini
    if (quota.ALL && quota.ALL > 0) {
      const elevesALL = poolGlobal.filter(function(e) {
        return !e._ASSIGNED && String(e.LV2 || '').trim().toUpperCase() === 'ALL';
      });

      const assigned = Math.min(elevesALL.length, quota.ALL);
      for (let i = 0; i < assigned; i++) {
        elevesALL[i]._ASSIGNED = true;
        classesState[niveau].push(elevesALL[i]);
      }
      stats.ALL += assigned;
      logLine('INFO', '  ✅ ' + niveau + ' : ' + assigned + ' élèves ALL placés');
    }

    // 6. Dispatcher PT (LV2) si quota défini
    if (quota.PT && quota.PT > 0) {
      const elevesPT = poolGlobal.filter(function(e) {
        return !e._ASSIGNED && String(e.LV2 || '').trim().toUpperCase() === 'PT';
      });

      const assigned = Math.min(elevesPT.length, quota.PT);
      for (let i = 0; i < assigned; i++) {
        elevesPT[i]._ASSIGNED = true;
        classesState[niveau].push(elevesPT[i]);
      }
      stats.PT += assigned;
      logLine('INFO', '  ✅ ' + niveau + ' : ' + assigned + ' élèves PT placés');
    }
  }

  // ✅ IMPORTANT : Dispatcher les élèves restants (non assignés) de manière équilibrée
  // pour atteindre les effectifs cibles
  const nonAssignes = poolGlobal.filter(function(e) { return !e._ASSIGNED; });
  logLine('INFO', 'Phase 1 : ' + nonAssignes.length + ' élèves restants à placer');

  // Dispatcher les non assignés de manière round-robin pour équilibrer les effectifs
  let idx = 0;
  while (nonAssignes.length > 0 && idx < nonAssignes.length) {
    const eleve = nonAssignes[idx];

    // Trouver la classe avec le moins d'élèves
    let minClasse = niveaux[0];
    let minCount = classesState[niveaux[0]].length;
    for (const niveau of niveaux) {
      if (classesState[niveau].length < minCount) {
        minCount = classesState[niveau].length;
        minClasse = niveau;
      }
    }

    // Ajouter l'élève à cette classe
    classesState[minClasse].push(eleve);
    eleve._ASSIGNED = true;

    // Retirer de la liste des non assignés
    nonAssignes.splice(idx, 1);
  }

  // Nettoyer le flag _ASSIGNED
  poolGlobal.forEach(function(e) { delete e._ASSIGNED; });

  // Logger les effectifs finaux
  for (const niveau of niveaux) {
    logLine('INFO', '  📊 ' + niveau + ' : ' + classesState[niveau].length + ' élèves au total');
  }

  logLine('INFO', 'Phase 1I stats : ITA=' + stats.ITA + ', CHAV=' + stats.CHAV + ', LATIN=' + stats.LATIN + ', ESP=' + stats.ESP + ', ALL=' + stats.ALL + ', PT=' + stats.PT);

  return {
    stats: stats,
    warn: warn
  };
}

/**
 * @deprecated - Remplacée par dispatchElevesWithQuotas_()
 * Conservée pour compatibilité si besoin
 */
function assignOptionsThenLV2_(classesState, quotas, autorisations, niveaux) {
  const warn = [];
  const stats = {
    ITA: 0,
    CHAV: 0,
    LV2_ESP: 0,
    LV2_ALL: 0,
    LV2_PT: 0
  };

  // Pour chaque classe
  for (const [niveau, eleves] of Object.entries(classesState)) {
    const quota = quotas[niveau] || {};

    // 1. Affecter ITA
    if (quota.ITA && quota.ITA > 0) {
      const assigned = assignOptionToClass_(eleves, 'ITA', quota.ITA, niveau);
      stats.ITA = stats.ITA + assigned;
      if (assigned < quota.ITA) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' ITA affectés sur ' + quota.ITA + ' demandés');
      }
    }

    // 2. Affecter CHAV
    if (quota.CHAV && quota.CHAV > 0) {
      const assigned = assignOptionToClass_(eleves, 'CHAV', quota.CHAV, niveau);
      stats.CHAV = stats.CHAV + assigned;
      if (assigned < quota.CHAV) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' CHAV affectés sur ' + quota.CHAV + ' demandés');
      }
    }

    // 3. Affecter LV2 ESP
    if (quota.LV2_ESP && quota.LV2_ESP > 0) {
      const assigned = assignLV2ToClass_(eleves, 'ESP', quota.LV2_ESP, niveau);
      stats.LV2_ESP = stats.LV2_ESP + assigned;
      if (assigned < quota.LV2_ESP) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' ESP affectés sur ' + quota.LV2_ESP + ' demandés');
      }
    }

    // 4. Affecter LV2 ALL
    if (quota.LV2_ALL && quota.LV2_ALL > 0) {
      const assigned = assignLV2ToClass_(eleves, 'ALL', quota.LV2_ALL, niveau);
      stats.LV2_ALL = stats.LV2_ALL + assigned;
      if (assigned < quota.LV2_ALL) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' ALL affectés sur ' + quota.LV2_ALL + ' demandés');
      }
    }

    // 5. Affecter LV2 PT
    if (quota.LV2_PT && quota.LV2_PT > 0) {
      const assigned = assignLV2ToClass_(eleves, 'PT', quota.LV2_PT, niveau);
      stats.LV2_PT = stats.LV2_PT + assigned;
      if (assigned < quota.LV2_PT) {
        warn.push('Classe ' + niveau + ' : Seulement ' + assigned + ' PT affectés sur ' + quota.LV2_PT + ' demandés');
      }
    }
  }

  logLine('INFO', 'Phase 1I stats : ITA=' + stats.ITA + ', CHAV=' + stats.CHAV + ', ESP=' + stats.LV2_ESP + ', ALL=' + stats.LV2_ALL + ', PT=' + stats.LV2_PT);

  return {
    classesState: classesState,
    stats: stats,
    warn: warn
  };
}

/**
 * Affecte une option (ITA ou CHAV) à N élèves d'une classe
 * @returns {number} Nombre d'élèves effectivement affectés
 */
function assignOptionToClass_(eleves, optionName, targetCount, niveau) {
  let assigned = 0;

  // Parcourir les élèves sans cette option
  for (let i = 0; i < eleves.length && assigned < targetCount; i++) {
    const eleve = eleves[i];

    // Vérifier si l'élève a déjà cette option
    const currentValue = eleve[optionName] || '';
    if (currentValue === '' || currentValue === 'NON' || currentValue === '0') {
      eleve[optionName] = 'OUI';
      assigned = assigned + 1;
    }
  }

  return assigned;
}

/**
 * Affecte une LV2 (ESP, ALL, PT) à N élèves d'une classe
 * @returns {number} Nombre d'élèves effectivement affectés
 */
function assignLV2ToClass_(eleves, lv2Code, targetCount, niveau) {
  let assigned = 0;

  // Parcourir les élèves sans LV2 assignée
  for (let i = 0; i < eleves.length && assigned < targetCount; i++) {
    const eleve = eleves[i];

    // Vérifier si l'élève n'a pas encore de LV2
    const currentLV2 = eleve.LV2 || '';
    if (currentLV2 === '' || currentLV2 === 'ANG') {
      eleve.LV2 = lv2Code;
      assigned = assigned + 1;
    }
  }

  return assigned;
}

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans Phase2I_DissoAsso.js :
//  - lockAttributes_()
//  - applyDisso_()
//  - applyAsso_()
//  - findClasseWithoutCode_()
//  - moveEleveToClass_()
// ===================================================================

// ===================================================================
// 7. PHASE 3I : EFFECTIFS & PARITÉ
// ===================================================================

/**
 * Phase 3I : Compléter effectifs et équilibrer parité
 * LIT : onglet sélectionné
 * ÉCRIT : uniquement CACHE
 */
function Phase3I_completeAndParity_(ctx) {
  const warnings = [];

  // ✅ FIX CRITIQUE : Lire depuis CACHE (résultats P1/P2), PAS depuis sources !
  // Phase3 doit continuer le travail de Phase1 et Phase2, pas repartir de zéro
  const classesState = readElevesFromCache_(ctx);

  // Verrouiller Options/LV2/DISSO/ASSO
  lockAttributes_(classesState, {
    options: true,
    lv2: true,
    disso: true,
    asso: true
  });

  // Compléter jusqu'aux cibles d'effectifs
  reachHeadcountTargets_(classesState, ctx.targets, warnings);

  // Équilibrer la parité F/M
  enforceParity_(classesState, ctx.tolParite, warnings);

  // ✅ FAIL-SAFE : Placer les élèves non placés dans la classe avec le plus grand déficit
  const unplaced = placeRemainingStudents_(classesState, ctx.targets, warnings);
  if (unplaced > 0) {
    logLine('WARN', '⚠️ ' + unplaced + ' élève(s) non placé(s) après P3 - placés dans classe déficitaire');
  }

  // Écrire dans CACHE
  writeAllClassesToCACHE_(ctx, classesState);

  return {
    ok: true,
    warnings,
    unplaced: unplaced
  };
}

/**
 * Place les élèves non placés dans la classe avec le plus grand déficit
 * @returns {number} Nombre d'élèves placés
 */
function placeRemainingStudents_(classesState, targets, warnings) {
  // Calculer les déficits actuels
  const deficits = {};
  let totalPlaced = 0;
  
  for (const [niveau, target] of Object.entries(targets)) {
    const current = (classesState[niveau] || []).length;
    if (current < target) {
      deficits[niveau] = target - current;
    }
  }
  
  // Trouver la classe avec le plus grand déficit
  let maxDeficit = 0;
  let targetClass = null;
  
  for (const [niveau, deficit] of Object.entries(deficits)) {
    if (deficit > maxDeficit) {
      maxDeficit = deficit;
      targetClass = niveau;
    }
  }
  
  // Si pas de classe déficitaire, rien à faire
  if (!targetClass || maxDeficit === 0) {
    return 0;
  }
  
  // Chercher les élèves non placés (dans une classe fictive ou hors cibles)
  // Pour l'instant, on vérifie juste que toutes les classes sont à leur cible
  const totalStudents = Object.values(classesState).reduce(function(sum, eleves) {
    return sum + eleves.length;
  }, 0);
  
  const totalTargets = Object.values(targets).reduce(function(sum, t) {
    return sum + t;
  }, 0);
  
  if (totalStudents < totalTargets) {
    logLine('WARN', '⚠️ Déficit global : ' + totalStudents + ' élèves pour ' + totalTargets + ' places cibles');
  }
  
  return totalPlaced;
}

/**
 * Complète les effectifs et ÉQUILIBRE les classes
 * ✅ NOUVELLE LOGIQUE : Équilibrage réel au lieu d'atteindre des cibles impossibles
 */
function reachHeadcountTargets_(classesState, targets, warnings) {
  const maxIterations = 50;
  let iteration = 0;

  while (iteration < maxIterations) {
    iteration++;

    // Calculer les effectifs actuels
    const currentCounts = {};
    for (const [niveau, eleves] of Object.entries(classesState)) {
      currentCounts[niveau] = eleves.length;
    }

    // Calculer la moyenne et trouver les classes les plus déséquilibrées
    const niveaux = Object.keys(currentCounts);
    const totalEleves = Object.values(currentCounts).reduce((sum, c) => sum + c, 0);
    const moyenne = totalEleves / niveaux.length;

    // Trouver la classe la plus remplie et la classe la plus vide
    let classeMax = null;
    let effectifMax = 0;
    let classeMin = null;
    let effectifMin = Infinity;

    for (const [niveau, effectif] of Object.entries(currentCounts)) {
      if (effectif > effectifMax) {
        effectifMax = effectif;
        classeMax = niveau;
      }
      if (effectif < effectifMin) {
        effectifMin = effectif;
        classeMin = niveau;
      }
    }

    // Si la différence est <= 1, on a atteint un équilibre acceptable
    const ecart = effectifMax - effectifMin;
    if (ecart <= 1) {
      logLine('INFO', `✅ Effectifs équilibrés : écart max = ${ecart} élève(s)`);
      break;
    }

    // Déplacer UN élève LIBRE de la classe la plus remplie vers la classe la plus vide
    const srcEleves = classesState[classeMax];
    let eleveToMove = null;

    for (let j = srcEleves.length - 1; j >= 0; j--) {
      const eleve = srcEleves[j];
      const codeA = eleve.ASSO || eleve.A || eleve['Code A'] || '';
      const codeD = eleve.DISSO || eleve.D || eleve['Code D'] || '';

      // ✅ Élève libre = pas de code ASSO ni DISSO
      if ((!codeA || codeA === '') && (!codeD || codeD === '')) {
        eleveToMove = eleve;
        break;
      }
    }

    if (eleveToMove) {
      moveEleveToClass_(classesState, eleveToMove, classeMax, classeMin);
      logLine('INFO', `  Équilibrage : Déplacé élève de ${classeMax} (${effectifMax}) vers ${classeMin} (${effectifMin})`);
    } else {
      // Aucun élève libre, arrêter
      logLine('WARN', `⚠️ Impossible d'équilibrer davantage : tous les élèves de ${classeMax} ont ASSO/DISSO`);
      break;
    }
  }

  // Vérification finale
  const finalCounts = {};
  for (const [niveau, eleves] of Object.entries(classesState)) {
    finalCounts[niveau] = eleves.length;
  }

  const effectifs = Object.values(finalCounts);
  const min = Math.min(...effectifs);
  const max = Math.max(...effectifs);
  const ecartFinal = max - min;

  logLine('INFO', `📊 Effectifs finaux : min=${min}, max=${max}, écart=${ecartFinal}`);

  if (ecartFinal > 2) {
    warnings.push(`⚠️ Déséquilibre persistant : écart de ${ecartFinal} élèves entre classes`);
  }
}

/**
 * Équilibre la parité F/M
 */
function enforceParity_(classesState, tolerance, warnings) {
  const maxIterations = 100;
  let iteration = 0;

  while (iteration < maxIterations) {
    iteration = iteration + 1;
    let swapped = false;

    // Pour chaque classe, calculer le déséquilibre
    const imbalances = [];

    for (const [niveau, eleves] of Object.entries(classesState)) {
      let countF = 0;
      let countM = 0;

      for (const eleve of eleves) {
        const genre = eleve.Genre || eleve.Sexe || '';
        if (genre === 'F' || genre === 'Fille') {
          countF = countF + 1;
        } else if (genre === 'M' || genre === 'Garçon' || genre === 'G') {
          countM = countM + 1;
        }
      }

      const delta = Math.abs(countF - countM);

      if (delta > tolerance) {
        imbalances.push({
          niveau: niveau,
          countF: countF,
          countM: countM,
          delta: delta,
          excessGenre: countF > countM ? 'F' : 'M'
        });
      }
    }

    if (imbalances.length === 0) {
      // Parité OK partout
      break;
    }

    // Essayer de faire des swaps entre classes déséquilibrées
    for (let i = 0; i < imbalances.length; i++) {
      const class1 = imbalances[i];

      for (let j = i + 1; j < imbalances.length; j++) {
        const class2 = imbalances[j];

        // Swap possible si les deux classes ont des excès opposés
        if (class1.excessGenre !== class2.excessGenre) {
          // class1 a trop de F, class2 a trop de M (ou inversement)
          const genreToSwapFrom1 = class1.excessGenre;
          const genreToSwapFrom2 = class2.excessGenre;

          // Chercher un élève de genreToSwapFrom1 dans class1
          const eleve1 = findEleveByGenre_(classesState[class1.niveau], genreToSwapFrom1);

          // Chercher un élève de genreToSwapFrom2 dans class2
          const eleve2 = findEleveByGenre_(classesState[class2.niveau], genreToSwapFrom2);

          if (eleve1 && eleve2) {
            // Faire le swap
            swapEleves_(classesState, eleve1, class1.niveau, eleve2, class2.niveau);
            swapped = true;
            logLine('INFO', '  Parité : Swap entre ' + class1.niveau + ' et ' + class2.niveau);
            break;
          }
        }
      }

      if (swapped) {
        break;
      }
    }

    if (!swapped) {
      // Plus de swap possible
      break;
    }
  }

  // Vérifier les classes encore déséquilibrées
  for (const [niveau, eleves] of Object.entries(classesState)) {
    let countF = 0;
    let countM = 0;

    for (const eleve of eleves) {
      const genre = eleve.Genre || eleve.Sexe || '';
      if (genre === 'F' || genre === 'Fille') {
        countF = countF + 1;
      } else if (genre === 'M' || genre === 'Garçon' || genre === 'G') {
        countM = countM + 1;
      }
    }

    const delta = Math.abs(countF - countM);

    if (delta > tolerance) {
      warnings.push('Classe ' + niveau + ' : Parité |F-M|=' + delta + ' > tolérance ' + tolerance);
    }
  }
}

// findEleveByGenre_() → supprimée (définition canonique dans App.Core.js)

/**
 * Échange deux élèves entre deux classes
 */
function swapEleves_(classesState, eleve1, classe1, eleve2, classe2) {
  // Retirer eleve1 de classe1
  const array1 = classesState[classe1];
  const idx1 = array1.indexOf(eleve1);
  if (idx1 > -1) {
    array1.splice(idx1, 1);
  }

  // Retirer eleve2 de classe2
  const array2 = classesState[classe2];
  const idx2 = array2.indexOf(eleve2);
  if (idx2 > -1) {
    array2.splice(idx2, 1);
  }

  // Ajouter eleve1 à classe2
  array2.push(eleve1);
  eleve1.Classe = classe2;

  // Ajouter eleve2 à classe1
  array1.push(eleve2);
  eleve2.Classe = classe1;
}

// ===================================================================
// 8. PHASE 4 : SWAPS AVEC VERROUS + MINI-GARDIEN LV2/OPT
// ===================================================================

/**
 * Construit l'offre LV2/OPT par classe à partir de _STRUCTURE
 * Retour : { "6°1": { LV2:Set, OPT:Set }, ... }
 */
function buildOffersFromStructure_(ctx) {
  const offers = {};
  const struct = readStructureSheet_(); // doit retourner lignes avec colonnes Classe / OPTIONS
  
  // Exemples d'OPTIONS dans _STRUCTURE : "ITA=6, CHAV=10" ou "LV2:ITA | OPT:CHAV"
  (struct.rows || []).forEach(function(row) {
    const classe = String(row.classe || row.Classe || row[0] || '').trim().replace(/CACHE|TEST|FIN$/,'');
    if (!classe) return;
    
    const optCell = String(row.options || row.OPTIONS || row[3] || '').toUpperCase();
    const lv2Set = new Set();
    const optSet = new Set();
    
    // parse très tolérant : récupère les libellés à gauche des "=" et après "LV2:"/"OPT:"
    optCell.split(/[,|]/).forEach(function(tok) {
      const t = tok.trim();
      if (!t) return;
      
      const mEq = t.match(/^([A-ZÉÈÀ]+)\s*=/); // ex: ITA=6
      const mTag = t.match(/^(LV2|OPT)\s*:\s*([A-ZÉÈÀ]+)/);
      
      if (mEq) {
        const tag = mEq[1];
        // heuristique : LV2 habituelles
        if (/^(ITA|ALL|ESP|PT|CHI)$/.test(tag)) {
          lv2Set.add(tag);
        } else {
          optSet.add(tag);
        }
      } else if (mTag) {
        if (mTag[1] === 'LV2') {
          lv2Set.add(mTag[2]);
        } else {
          optSet.add(mTag[2]);
        }
      } else {
        // si pas de "=", ranger dans OPT par défaut sauf si LV2 connue
        if (/^(ITA|ALL|ESP|PT|CHI)$/.test(t)) {
          lv2Set.add(t);
        } else {
          optSet.add(t);
        }
      }
    });
    
    offers[classe] = { LV2: lv2Set, OPT: optSet };
  });
  
  return offers;
}

// isPlacementLV2OPTOK_() → supprimée (définition canonique dans App.Core.js)

/**
 * Phase 4 : Optimisation par swaps (COM prioritaire)
 * LIT : depuis CACHE (résultats des phases 1/2/3)
 * ÉCRIT : uniquement CACHE
 */
function Phase4_balanceScoresSwaps_(ctx) {
  const warnings = [];

  // ✅ CORRECTIF : Lire depuis CACHE (résultats phases 1/2/3), PAS depuis TEST !
  logLine('INFO', 'Phase4 : Lecture depuis CACHE (résultats phases 1/2/3)...');
  const classesState = readElevesFromCache_(ctx);

  // 🔒 Construire l'offre LV2/OPT pour le mini-gardien
  const offers = buildOffersFromStructure_(ctx);
  logLine('INFO', '🔒 Mini-gardien LV2/OPT activé');

  // Définir TOUS les verrous
  const lock = {
    keepOptions: true,
    keepLV2: true,
    keepDisso: true,
    keepAsso: true,
    keepParity: true,
    keepQuotas: true
  };

  // Lancer le moteur de swaps avec le mini-gardien
  const res = runSwapEngineV14_withLocks_LEGACY_(
    classesState,
    {
      metrics: ['COM', 'TRA', 'PART', 'ABS'],
      primary: 'COM', // Priorité absolue sur COM
      maxSwaps: ctx.maxSwaps || 1000,
      weights: ctx.weights || { parity: 0.3, com: 0.4, tra: 0.1, part: 0.1, abs: 0.1 },
      parityTol: ctx.tolParite || ctx.parityTolerance || 2,
      runtimeSec: ctx.runtimeSec || 180  // Budget temps (défaut: 3 min)
    },
    lock,
    warnings,
    ctx,
    offers  // 🔒 Passer l'offre au moteur
  );

  // Écrire dans CACHE
  writeAllClassesToCACHE_(ctx, classesState);

  logLine('INFO', '✅ Phase 4 terminée : ' + (res.applied || 0) + ' swaps appliqués, ' + (res.skippedByLV2OPT || 0) + ' refusés (LV2/OPT)');

  return {
    ok: true,
    warnings,
    swapsApplied: res.applied || 0,
    skippedByLV2OPT: res.skippedByLV2OPT || 0
  };
}

/**
 * Moteur de swaps avec verrous + timeboxing + anti-stagnation
 */
function runSwapEngineV14_withLocks_LEGACY_(classesState, options, locks, warnings, ctx, offers) {
  const metrics = options.metrics || ['COM', 'TRA', 'PART', 'ABS'];
  const primary = options.primary || 'COM';
  const maxSwaps = options.maxSwaps || 1000;
  const weights = options.weights || { parity: 0.3, com: 0.4, tra: 0.1, part: 0.1, abs: 0.1 };
  const parityTol = options.parityTol || 2;
  const runtimeSec = options.runtimeSec || 180;

  let applied = 0;
  let skippedByLV2OPT = 0;  // 🔒 Compteur de swaps refusés
  let iteration = 0;
  let itersWithoutImprovement = 0;
  const startTime = new Date().getTime();
  const endTime = startTime + (runtimeSec * 1000);

  logLine('INFO', '  Phase 4 : Démarrage swaps (max=' + maxSwaps + ', runtime=' + runtimeSec + 's, priorité=' + primary + ')');
  logLine('INFO', '  Poids: COM=' + weights.com + ', TRA=' + weights.tra + ', PART=' + weights.part + ', ABS=' + weights.abs + ', Parité=' + weights.parity);
  logLine('INFO', '  Tolérance parité: ' + parityTol);

  // Construire l'offre pour valider les swaps
  const offer = buildOfferWithQuotas_(ctx);
  
  // 📊 Stats mobilité initiale
  const mobilityStats = computeMobilityStats_LEGACY_(classesState, offer);
  logLine('INFO', '  📊 Mobilité: LIBRE=' + mobilityStats.libre + ', FIXE=' + mobilityStats.fixe + ', TOTAL=' + mobilityStats.total);

  // Timeboxing : boucle tant que temps restant ET swaps < max
  while (new Date().getTime() < endTime && applied < maxSwaps) {
    iteration = iteration + 1;

    // ✅ Calculer les counts actuels pour vérifier les quotas en temps réel
    const counts = computeCountsFromState_(classesState);

    // Calculer les scores actuels de toutes les classes
    const currentScores = calculateClassScores_LEGACY_(classesState, metrics);

    // Trouver le meilleur swap possible (avec poids et tolérance)
    const bestSwap = findBestSwap_LEGACY_(classesState, currentScores, primary, locks, offer, counts, weights, parityTol);

    if (!bestSwap) {
      // Plus de swap bénéfique
      itersWithoutImprovement++;
      
      // 🔄 ÉCHAPPATOIRE ANTI-STAGNATION : après 200 iters sans amélioration
      if (itersWithoutImprovement >= 200) {
        logLine('INFO', '  🔄 Stagnation détectée (200 iters) - relaxation minime des poids');
        weights.com *= 0.98;  // Relaxation très faible
        itersWithoutImprovement = 0;
        continue;
      }
      
      logLine('INFO', '  Phase 4 : Aucun swap bénéfique trouvé (iteration ' + iteration + ')');
      break;
    }

    // 🔒 MINI-GARDIEN : refuser si LV2/OPT non proposés dans la classe cible
    if (offers && 
        (!isPlacementLV2OPTOK_(bestSwap.eleve1, bestSwap.classe2, offers) || 
         !isPlacementLV2OPTOK_(bestSwap.eleve2, bestSwap.classe1, offers))) {
      skippedByLV2OPT++;
      continue;  // Ignorer ce swap
    }

    // Appliquer le swap
    swapEleves_(classesState, bestSwap.eleve1, bestSwap.classe1, bestSwap.eleve2, bestSwap.classe2);
    applied = applied + 1;
    itersWithoutImprovement = 0;  // Reset compteur

    if (applied % 20 === 0) {
      const elapsed = Math.round((new Date().getTime() - startTime) / 1000);
      logLine('INFO', '  Phase 4 : ' + applied + ' swaps appliqués (elapsed=' + elapsed + 's)...');
    }
  }

  const elapsedTotal = Math.round((new Date().getTime() - startTime) / 1000);
  logLine('INFO', '  ✅ Phase 4 terminée : elapsed=' + elapsedTotal + 's | iters=' + iteration + ' | swaps=' + applied);
  if (skippedByLV2OPT > 0) {
    logLine('INFO', '  🔒 Mini-gardien : ' + skippedByLV2OPT + ' swaps refusés (LV2/OPT incompatible)');
  }
  
  // 🛡️ GARDE-FOU FINAL PARITÉ : si une classe reste hors tolérance
  // ✅ CORRECTIF: Recalculer counts après la boucle (hors scope)
  const countsAfterSwaps = computeCountsFromState_(classesState);
  applyParityGuardrail_LEGACY_(classesState, parityTol, offer, countsAfterSwaps);

  return {
    applied: applied,
    skippedByLV2OPT: skippedByLV2OPT,
    elapsed: elapsedTotal,
    iterations: iteration
  };
}

// calculateClassScores_LEGACY_() → supprimée (définition canonique dans App.Core.js)
// calculateClassMetric_LEGACY_() → supprimée (définition canonique dans App.Core.js)

/**
 * Trouve le meilleur swap possible avec objectifs hiérarchisés et pondérés
 * Priorité 1 : Parité (si hors tolérance)
 * Priorité 2 : Scores pondérés (COM/TRA/PART/ABS) selon poids UI
 */
function findBestSwap_LEGACY_(classesState, currentScores, primary, locks, offer, counts, weights, parityTol) {
  let bestSwap = null;
  let bestScore = -Infinity;

  const niveaux = Object.keys(classesState);
  
  // Utiliser les poids passés en paramètre (depuis le contexte)
  weights = weights || {
    parity: 0.3,
    com: 0.4,
    tra: 0.1,
    part: 0.1,
    abs: 0.1
  };
  parityTol = parityTol || 2;

  // Parcourir toutes les paires de classes
  for (let i = 0; i < niveaux.length; i++) {
    const niveau1 = niveaux[i];
    const eleves1 = classesState[niveau1];

    for (let j = i + 1; j < niveaux.length; j++) {
      const niveau2 = niveaux[j];
      const eleves2 = classesState[niveau2];

      // Essayer tous les swaps entre ces deux classes
      for (const eleve1 of eleves1) {
        // ✅ Vérifier mobilité élève1 (LIBRE ou PERMUT uniquement)
        if (!isEleveMobile_LEGACY_(eleve1, counts, niveau1, offer)) {
          continue;
        }

        for (const eleve2 of eleves2) {
          // ✅ Vérifier mobilité élève2
          if (!isEleveMobile_LEGACY_(eleve2, counts, niveau2, offer)) {
            continue;
          }
          
          // ✅ Vérifier si le swap est valide selon les verrous + quotas
          if (!isSwapValid_LEGACY_(eleve1, niveau1, eleve2, niveau2, locks, classesState, offer, counts)) {
            continue;
          }

          // Calculer le score d'amélioration (hiérarchisé + pondéré)
          const swapScore = calculateSwapScore_LEGACY_(
            eleve1,
            niveau1,
            eleve2,
            niveau2,
            classesState,
            weights,
            parityTol
          );

          if (swapScore > bestScore) {
            bestScore = swapScore;
            bestSwap = {
              eleve1: eleve1,
              classe1: niveau1,
              eleve2: eleve2,
              classe2: niveau2,
              improvement: swapScore
            };
          }
        }
      }
    }
  }

  return bestSwap;
}

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.Core.js :
//  - computeCountsFromState_()
//  - isMoveAllowed_()
//  - eligibleForSwap_LEGACY_()
//  - isSwapValid_LEGACY_()
// ===================================================================

// computeMobilityStats_LEGACY_() → supprimée (définition canonique dans App.Core.js)

/**
 * Garde-fou final parité : si une classe reste hors tolérance,
 * force un swap greedy avec la classe la plus opposée en parité
 */
function applyParityGuardrail_LEGACY_(classesState, parityTol, offer, counts) {
  const niveaux = Object.keys(classesState);
  let swapped = false;
  
  // Identifier les classes hors tolérance
  const outOfTol = [];
  niveaux.forEach(function(niveau) {
    const state = computeClassState_LEGACY_(classesState[niveau]);
    if (Math.abs(state.deltaFM) > parityTol) {
      outOfTol.push({
        niveau: niveau,
        deltaFM: state.deltaFM,
        needsGenre: state.deltaFM > 0 ? 'M' : 'F'  // Si trop de F, besoin de M
      });
    }
  });
  
  if (outOfTol.length === 0) {
    logLine('INFO', '  🛡️ Garde-fou parité : Toutes les classes dans la tolérance');
    return;
  }
  
  logLine('WARN', '  🛡️ Garde-fou parité : ' + outOfTol.length + ' classe(s) hors tolérance');
  
  // Pour chaque classe hors tolérance, chercher un swap greedy
  outOfTol.forEach(function(cls1) {
    // Trouver la classe la plus opposée en parité
    let bestTarget = null;
    let bestDelta = 0;
    
    niveaux.forEach(function(niveau2) {
      if (niveau2 === cls1.niveau) return;
      const state2 = computeClassState_LEGACY_(classesState[niveau2]);
      // Opposé = l'une a trop de F, l'autre trop de M
      if ((cls1.deltaFM > 0 && state2.deltaFM < 0) || (cls1.deltaFM < 0 && state2.deltaFM > 0)) {
        const delta = Math.abs(cls1.deltaFM) + Math.abs(state2.deltaFM);
        if (delta > bestDelta) {
          bestDelta = delta;
          bestTarget = niveau2;
        }
      }
    });
    
    if (!bestTarget) return;
    
    // Chercher un swap entre cls1 et bestTarget
    const eleves1 = classesState[cls1.niveau];
    const eleves2 = classesState[bestTarget];
    
    for (let i = 0; i < eleves1.length && !swapped; i++) {
      const e1 = eleves1[i];
      const genre1 = String(e1.SEXE || e1.Genre || e1.Sexe || '').toUpperCase();
      
      if (genre1 !== cls1.needsGenre) continue;  // On cherche le genre opposé
      
      for (let j = 0; j < eleves2.length && !swapped; j++) {
        const e2 = eleves2[j];
        const genre2 = String(e2.SEXE || e2.Genre || e2.Sexe || '').toUpperCase();
        
        if (genre2 === cls1.needsGenre) continue;  // Doit être opposé
        
        // Vérifier mobilité
        if (!isEleveMobile_LEGACY_(e1, counts, cls1.niveau, offer)) continue;
        if (!isEleveMobile_LEGACY_(e2, counts, bestTarget, offer)) continue;
        
        // Appliquer le swap
        swapEleves_(classesState, e1, cls1.niveau, e2, bestTarget);
        logLine('INFO', '  🛡️ Swap parité forcé : ' + cls1.niveau + ' ↔ ' + bestTarget);
        swapped = true;
      }
    }
  });
}

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.Core.js :
//  - isEleveMobile_LEGACY_()
//  - calculateSwapScore_LEGACY_()
//  - computeClassState_LEGACY_()
//  - simulateSwapState_LEGACY_()
// ===================================================================

// ===================================================================
// 9. UTILITAIRES — logLine, _u_, _arr → supprimés (App.Core.js)
// ===================================================================

// ensureColumn_() → supprimée (définition canonique dans App.SheetsData.js)

// ===================================================================
// FONCTIONS SUPPRIMÉES — définitions canoniques dans App.Core.js / App.Context.js :
//  - buildClassOffers_() → App.Context.js
//  - computeAllow_()     → App.Context.js
//  - parseCodes_()       → App.Core.js
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

// ===================================================================
// 10. WRAPPER POUR APPEL DEPUIS L'INTERFACE
// ===================================================================

/**
 * Wrapper appelé depuis l'interface UI
 */
function lancerOptimisationV14I_Wrapper(options) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    SpreadsheetApp.getUi().alert(
      'Optimisation déjà en cours',
      'Veuillez patienter...',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { success: false, message: 'Verrouillé' };
  }

  try {
    const result = runOptimizationV14FullI(options || {});

    const msg = result.ok
      ? '✅ Optimisation réussie !\n' + result.swaps + ' swaps appliqués en ' + result.duration + 's'
      : '❌ Échec de l\'optimisation\n' + result.warnings.join('\n');

    SpreadsheetApp.getUi().alert('Résultat Optimisation V14I', msg, SpreadsheetApp.getUi().ButtonSet.OK);

    return {
      success: result.ok,
      swaps: result.swaps,
      duration: result.duration,
      warnings: result.warnings
    };

  } catch (e) {
    logLine('ERROR', 'Erreur wrapper : ' + e.message);
    SpreadsheetApp.getUi().alert('Erreur', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
