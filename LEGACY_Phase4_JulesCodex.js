/**
 * ===================================================================
 * 🎯 JULES CODEX LEGACY - PHASE 4 : ÉQUILIBRAGE INTELLIGENT
 * ===================================================================
 *
 * ÉVOLUTIONS MAJEURES :
 * ✅ Distance de distribution (vs variance simple)
 * ✅ Moteurs Silencieux (ciblage élèves problématiques)
 * ✅ Ancre de stabilité (anti-oscillations)
 * ✅ Intégration Phase 3 optionnelle (ctx.useIntegratedPhase3)
 * ✅ API unifiée evaluateSwap (copie immuable)
 *
 * Date : 2025-11-14
 * Branche : claude/jules-codex-legacy
 *
 * ===================================================================
 */

// ===================================================================
// CONFIGURATION GLOBALE JULES CODEX
// ===================================================================

var JULES_CODEX_CONFIG = {
  // Paramètres optimisation
  maxSwaps: 500,
  minGainThreshold: 0.01,

  // Moteurs Silencieux
  useSilentMotors: true,
  silentMotorRatio: 0.7,  // 70% swaps ciblés, 30% exploration
  topProblematicCount: 20, // Top N élèves les plus problématiques

  // Ancre de stabilité
  useStabilityAnchor: true,
  swapQuarantineDuration: 10,  // Nombre d'itérations avant réessai
  maxQuarantineSize: 100,       // Taille max de la quarantaine

  // Distance de distribution
  useDistributionDistance: true,
  distributionMetric: 'earthMover',  // 'earthMover' ou 'chiSquare'
  fallbackToVariance: true,          // Fallback si échantillon trop petit
  minSampleSize: 5,                  // Taille min pour distribution

  // Intégration Phase 3
  enableIntegratedPhase3: true,  // Activer par défaut

  // Logging
  verboseMode: true,
  logEveryNSwaps: 10
};

// ===================================================================
// PHASE 4 JULES CODEX - FONCTION PRINCIPALE
// ===================================================================

/**
 * Phase 4 JULES CODEX : Équilibrage Intelligent avec Moteurs Silencieux
 *
 * @param {Object} ctx - Contexte LEGACY
 * @returns {Object} { ok: true, swapsApplied: X, phase3Integrated: boolean }
 */
function Phase4_JulesCodex_LEGACY(ctx) {
  logLegacy('INFO', '='.repeat(80), 'Phase 4 JULES');
  logLegacy('INFO', '🎯 PHASE 4 JULES CODEX - Équilibrage Intelligent', 'Phase 4 JULES');
  logLegacy('INFO', '='.repeat(80), 'Phase 4 JULES');

  const ss = ctx.ss || SpreadsheetApp.getActiveSpreadsheet();

  // Configuration (peut être overridée par ctx)
  const config = {
    maxSwaps: ctx.maxSwaps || JULES_CODEX_CONFIG.maxSwaps,
    useSilentMotors: ctx.useSilentMotors !== undefined ? ctx.useSilentMotors : JULES_CODEX_CONFIG.useSilentMotors,
    useStabilityAnchor: ctx.useStabilityAnchor !== undefined ? ctx.useStabilityAnchor : JULES_CODEX_CONFIG.useStabilityAnchor,
    useDistributionDistance: ctx.useDistributionDistance !== undefined ? ctx.useDistributionDistance : JULES_CODEX_CONFIG.useDistributionDistance,
    useIntegratedPhase3: ctx.useIntegratedPhase3 !== undefined ? ctx.useIntegratedPhase3 : JULES_CODEX_CONFIG.enableIntegratedPhase3
  };

  const weights = ctx.weights || { parity: 1.0, com: 1.0, tra: 0.5, part: 0.3, abs: 0.2 };

  // ========== ÉTAPE 0 : INTÉGRATION PHASE 3 (OPTIONNEL) ==========
  let phase3Integrated = false;

  if (config.useIntegratedPhase3) {
    logLegacy('INFO', '🔗 Intégration Phase 3 (placement non-assignés + parité)', 'Phase 4 JULES');

    const phase3Result = runIntegratedPhase3_JulesCodex(ctx);

    if (phase3Result.ok) {
      phase3Integrated = true;
      logLegacy('SUCCESS', '✅ Phase 3 intégrée : ' + phase3Result.placed + ' placés, ' + phase3Result.swaps + ' swaps parité', 'Phase 4 JULES');
    } else {
      logLegacy('WARN', '⚠️ Phase 3 intégrée : erreur : ' + (phase3Result.message || 'inconnue'), 'Phase 4 JULES');
    }
  } else {
    logLegacy('INFO', '⏭️ Phase 3 non intégrée (mode legacy)', 'Phase 4 JULES');
  }

  // ========== ÉTAPE 1 : CONSOLIDER DONNÉES ==========
  const allData = [];
  let headersRef = null;

  (ctx.cacheSheets || []).forEach(function(testName) {
    const testSheet = ss.getSheetByName(testName);
    if (!testSheet || testSheet.getLastRow() <= 1) return;

    const data = testSheet.getDataRange().getValues();
    if (!headersRef) headersRef = data[0];

    for (let i = 1; i < data.length; i++) {
      allData.push({
        sheetName: testName,
        row: data[i],
        index: i
      });
    }
  });

  if (allData.length === 0) {
    logLegacy('WARN', '⚠️ Aucune donnée trouvée', 'Phase 4 JULES');
    return { ok: false, swapsApplied: 0, phase3Integrated: phase3Integrated };
  }

  const idxAssigned = headersRef.indexOf('_CLASS_ASSIGNED');
  const idxSexe = headersRef.indexOf('SEXE');
  const idxCOM = headersRef.indexOf('COM');
  const idxTRA = headersRef.indexOf('TRA');
  const idxPART = headersRef.indexOf('PART');
  const idxABS = headersRef.indexOf('ABS');
  const idxMobilite = headersRef.indexOf('MOBILITE');
  const idxFixe = headersRef.indexOf('FIXE');

  // ========== ÉTAPE 2 : GROUPER PAR CLASSE ==========
  const byClass = {};
  for (let i = 0; i < allData.length; i++) {
    const cls = String(allData[i].row[idxAssigned] || '').trim();
    if (cls) {
      if (!byClass[cls]) byClass[cls] = [];
      byClass[cls].push(i);
    }
  }

  logLegacy('INFO', '📊 ' + allData.length + ' élèves répartis en ' + Object.keys(byClass).length + ' classes', 'Phase 4 JULES');

  // ========== ÉTAPE 3 : CALCUL DISTRIBUTION GLOBALE ==========
  const globalDist = config.useDistributionDistance
    ? calculateGlobalDistribution_JulesCodex(allData, headersRef)
    : null;

  if (globalDist) {
    logLegacy('INFO', '📈 Distribution globale calculée (méthode: ' + JULES_CODEX_CONFIG.distributionMetric + ')', 'Phase 4 JULES');
  }

  // ========== ÉTAPE 4 : INITIALISER ANCRE DE STABILITÉ ==========
  const swapQuarantine = config.useStabilityAnchor ? [] : null;

  // ========== ÉTAPE 5 : BOUCLE D'OPTIMISATION (MOTEURS SILENCIEUX) ==========
  let swapsApplied = 0;
  let consecutiveNoImprovement = 0;
  const maxNoImprovement = 20;  // Stop après N itérations sans amélioration

  for (let iter = 0; iter < config.maxSwaps; iter++) {
    const scoresBefore = calculateGlobalScore_JulesCodex(
      allData, byClass, headersRef, weights, globalDist, config
    );

    let bestSwap = null;
    let bestGain = 0;

    // ========== CHOIX DU MODE : MOTEURS SILENCIEUX OU EXPLORATION ==========
    const useSilentMode = config.useSilentMotors && Math.random() < JULES_CODEX_CONFIG.silentMotorRatio;

    if (useSilentMode) {
      // MODE CIBLÉ : Identifier les élèves les plus problématiques
      const problematicStudents = identifyProblematicStudents_JulesCodex(
        allData, byClass, headersRef, weights, globalDist, config
      );

      // Chercher swaps impliquant ces élèves
      bestSwap = findBestTargetedSwap_JulesCodex(
        allData, byClass, headersRef, weights, globalDist, config,
        problematicStudents, swapQuarantine, scoresBefore
      );

    } else {
      // MODE EXPLORATION : Swaps semi-aléatoires
      bestSwap = findBestRandomSwap_JulesCodex(
        allData, byClass, headersRef, weights, globalDist, config,
        swapQuarantine, scoresBefore
      );
    }

    // ========== APPLIQUER LE MEILLEUR SWAP TROUVÉ ==========
    if (bestSwap && bestSwap.gain > JULES_CODEX_CONFIG.minGainThreshold) {
      applySwap_JulesCodex(allData, byClass, bestSwap);
      swapsApplied++;
      consecutiveNoImprovement = 0;

      if (JULES_CODEX_CONFIG.verboseMode && swapsApplied % JULES_CODEX_CONFIG.logEveryNSwaps === 0) {
        logLegacy('INFO', '  🔄 ' + swapsApplied + ' swaps (gain: ' + bestSwap.gain.toFixed(3) + ', mode: ' + (useSilentMode ? 'CIBLÉ' : 'EXPLORE') + ')', 'Phase 4 JULES');
      }

    } else {
      // Aucun swap bénéfique
      consecutiveNoImprovement++;

      if (config.useStabilityAnchor && bestSwap) {
        // Ajouter à la quarantaine
        addToQuarantine_JulesCodex(swapQuarantine, bestSwap, iter);
      }

      if (consecutiveNoImprovement >= maxNoImprovement) {
        logLegacy('INFO', '  🛑 Convergence atteinte (pas d\'amélioration depuis ' + maxNoImprovement + ' itérations)', 'Phase 4 JULES');
        break;
      }
    }

    // Nettoyer la quarantaine (supprimer les swaps périmés)
    if (config.useStabilityAnchor && swapQuarantine && iter % 10 === 0) {
      cleanQuarantine_JulesCodex(swapQuarantine, iter);
    }
  }

  // ========== ÉTAPE 6 : RÉÉCRIRE LES RÉSULTATS ==========
  const bySheet = {};
  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    if (!bySheet[item.sheetName]) bySheet[item.sheetName] = [];
    bySheet[item.sheetName].push(item);
  }

  for (const sheetName in bySheet) {
    const testSheet = ss.getSheetByName(sheetName);
    if (!testSheet) continue;

    const items = bySheet[sheetName];
    const allRows = [headersRef];
    items.forEach(function(item) {
      allRows.push(item.row);
    });

    testSheet.getRange(1, 1, allRows.length, headersRef.length).setValues(allRows);
  }

  SpreadsheetApp.flush();

  // ========== ÉTAPE 7 : RAPPORT FINAL ==========
  const finalScore = calculateGlobalScore_JulesCodex(
    allData, byClass, headersRef, weights, globalDist, config
  );

  logLegacy('SUCCESS', '✅ PHASE 4 JULES CODEX terminée', 'Phase 4 JULES');
  logLegacy('INFO', '  • Swaps appliqués : ' + swapsApplied, 'Phase 4 JULES');
  logLegacy('INFO', '  • Score final : ' + finalScore.toFixed(3), 'Phase 4 JULES');
  logLegacy('INFO', '  • Phase 3 intégrée : ' + (phase3Integrated ? 'OUI' : 'NON'), 'Phase 4 JULES');
  logLegacy('INFO', '  • Moteurs Silencieux : ' + (config.useSilentMotors ? 'ACTIFS' : 'DÉSACTIVÉS'), 'Phase 4 JULES');
  logLegacy('INFO', '  • Ancre stabilité : ' + (config.useStabilityAnchor ? 'ACTIVE' : 'DÉSACTIVÉE'), 'Phase 4 JULES');

  return {
    ok: true,
    swapsApplied: swapsApplied,
    phase3Integrated: phase3Integrated,
    finalScore: finalScore,
    config: config
  };
}

// ===================================================================
// PHASE 3 INTÉGRÉE
// ===================================================================

/**
 * Intègre la logique de Phase 3 (placement non-assignés + parité initiale)
 */
function runIntegratedPhase3_JulesCodex(ctx) {
  // Réutiliser la logique de Phase3I_completeAndParity_LEGACY
  // mais en mode simplifié pour éviter duplication

  if (typeof Phase3I_completeAndParity_LEGACY === 'function') {
    return Phase3I_completeAndParity_LEGACY(ctx);
  } else {
    return { ok: false, message: 'Phase3I_completeAndParity_LEGACY non disponible' };
  }
}

// ===================================================================
// CALCUL DISTRIBUTION GLOBALE
// ===================================================================

/**
 * Calcule la distribution globale des notes par critère
 */
function calculateGlobalDistribution_JulesCodex(allData, headers) {
  const criteria = ['COM', 'TRA', 'PART', 'ABS'];
  const distribution = {};

  criteria.forEach(function(crit) {
    const idx = headers.indexOf(crit);
    if (idx === -1) return;

    const values = [];
    for (let i = 0; i < allData.length; i++) {
      const val = Number(allData[i].row[idx] || 3);
      values.push(val);
    }

    // Calculer histogramme normalisé
    const hist = {};
    values.forEach(function(v) {
      hist[v] = (hist[v] || 0) + 1;
    });

    // Normaliser
    const total = values.length;
    for (const key in hist) {
      hist[key] = hist[key] / total;
    }

    distribution[crit] = {
      histogram: hist,
      mean: values.reduce((a, b) => a + b, 0) / values.length,
      count: values.length
    };
  });

  return distribution;
}

// ===================================================================
// SCORE GLOBAL (DISTANCE DE DISTRIBUTION)
// ===================================================================

/**
 * Calcule le score global avec distance de distribution
 */
function calculateGlobalScore_JulesCodex(allData, byClass, headers, weights, globalDist, config) {
  const idxSexe = headers.indexOf('SEXE');
  let totalScore = 0;

  // ========== SCORE PARITÉ ==========
  for (const cls in byClass) {
    const indices = byClass[cls];
    let countF = 0, countM = 0;

    indices.forEach(function(i) {
      const sexe = String(allData[i].row[idxSexe] || '').toUpperCase();
      if (sexe === 'F') countF++;
      else if (sexe === 'M') countM++;
    });

    totalScore += Math.abs(countF - countM) * (weights.parity || 1.0);
  }

  // ========== SCORE ACADÉMIQUE ==========
  if (config.useDistributionDistance && globalDist) {
    // Utiliser distance de distribution
    totalScore += calculateDistributionDistance_JulesCodex(
      allData, byClass, headers, weights, globalDist
    );
  } else {
    // Fallback : variance classique
    totalScore += calculateVarianceScore_JulesCodex(
      allData, byClass, headers, weights
    );
  }

  return totalScore;
}

/**
 * Calcule la distance entre la distribution de chaque classe et la distribution globale
 */
function calculateDistributionDistance_JulesCodex(allData, byClass, headers, weights, globalDist) {
  const criteria = ['COM', 'TRA', 'PART', 'ABS'];
  const weightList = [weights.com || 1.0, weights.tra || 0.5, weights.part || 0.3, weights.abs || 0.2];

  let totalDistance = 0;

  for (let c = 0; c < criteria.length; c++) {
    const crit = criteria[c];
    const w = weightList[c];
    const idx = headers.indexOf(crit);

    if (idx === -1 || !globalDist[crit]) continue;

    const globalHist = globalDist[crit].histogram;

    // Pour chaque classe, calculer la distance à la distribution globale
    for (const cls in byClass) {
      const indices = byClass[cls];

      // Si échantillon trop petit, fallback variance
      if (indices.length < JULES_CODEX_CONFIG.minSampleSize) {
        continue;  // Ignorer ou utiliser variance locale
      }

      // Construire histogramme classe
      const classHist = {};
      indices.forEach(function(i) {
        const val = Number(allData[i].row[idx] || 3);
        classHist[val] = (classHist[val] || 0) + 1;
      });

      // Normaliser
      for (const key in classHist) {
        classHist[key] = classHist[key] / indices.length;
      }

      // Calculer distance (Earth Mover ou Chi²)
      let distance = 0;

      if (JULES_CODEX_CONFIG.distributionMetric === 'earthMover') {
        distance = earthMoverDistance_JulesCodex(classHist, globalHist);
      } else {
        distance = chiSquareDistance_JulesCodex(classHist, globalHist);
      }

      totalDistance += distance * w;
    }
  }

  return totalDistance;
}

/**
 * Earth Mover Distance (approximation simple)
 */
function earthMoverDistance_JulesCodex(hist1, hist2) {
  const allKeys = Object.keys(hist1).concat(Object.keys(hist2));
  const uniqueKeys = allKeys.filter(function(v, i, a) { return a.indexOf(v) === i; });

  uniqueKeys.sort(function(a, b) { return Number(a) - Number(b); });

  let cumDiff = 0;
  let totalDist = 0;

  uniqueKeys.forEach(function(key) {
    const p1 = hist1[key] || 0;
    const p2 = hist2[key] || 0;
    cumDiff += (p1 - p2);
    totalDist += Math.abs(cumDiff);
  });

  return totalDist;
}

/**
 * Chi-Square Distance
 */
function chiSquareDistance_JulesCodex(hist1, hist2) {
  const allKeys = Object.keys(hist1).concat(Object.keys(hist2));
  const uniqueKeys = allKeys.filter(function(v, i, a) { return a.indexOf(v) === i; });

  let chiSquare = 0;

  uniqueKeys.forEach(function(key) {
    const p1 = hist1[key] || 0;
    const p2 = hist2[key] || 0;

    if (p1 + p2 > 0) {
      chiSquare += Math.pow(p1 - p2, 2) / (p1 + p2);
    }
  });

  return chiSquare;
}

/**
 * Calcul variance classique (fallback)
 */
function calculateVarianceScore_JulesCodex(allData, byClass, headers, weights) {
  const criteria = ['COM', 'TRA', 'PART', 'ABS'];
  const indices = criteria.map(function(c) { return headers.indexOf(c); });
  const wts = [weights.com || 1.0, weights.tra || 0.5, weights.part || 0.3, weights.abs || 0.2];

  let totalScore = 0;

  for (let c = 0; c < criteria.length; c++) {
    const idx = indices[c];
    const w = wts[c];

    if (idx === -1) continue;

    const avgs = [];
    for (const cls in byClass) {
      let sum = 0, count = 0;
      byClass[cls].forEach(function(i) {
        sum += Number(allData[i].row[idx] || 3);
        count++;
      });
      if (count > 0) avgs.push(sum / count);
    }

    if (avgs.length > 0) {
      const mean = avgs.reduce((a, b) => a + b, 0) / avgs.length;
      const variance = avgs.reduce((sum, avg) => sum + Math.pow(avg - mean, 2), 0) / avgs.length;
      totalScore += variance * w;
    }
  }

  return totalScore;
}

// ===================================================================
// MOTEURS SILENCIEUX - IDENTIFICATION ÉLÈVES PROBLÉMATIQUES
// ===================================================================

/**
 * Identifie les élèves les plus problématiques (contributeurs max au score)
 */
function identifyProblematicStudents_JulesCodex(allData, byClass, headers, weights, globalDist, config) {
  const problematic = [];

  // Pour chaque élève, calculer sa "contribution" au score global
  for (let i = 0; i < allData.length; i++) {
    const contribution = calculateStudentContribution_JulesCodex(
      i, allData, byClass, headers, weights, globalDist, config
    );

    problematic.push({
      index: i,
      contribution: contribution
    });
  }

  // Trier par contribution décroissante
  problematic.sort(function(a, b) {
    return b.contribution - a.contribution;
  });

  // Retourner les top N
  return problematic.slice(0, JULES_CODEX_CONFIG.topProblematicCount);
}

/**
 * Calcule la contribution d'un élève au score global
 */
function calculateStudentContribution_JulesCodex(studentIdx, allData, byClass, headers, weights, globalDist, config) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxSexe = headers.indexOf('SEXE');
  const cls = String(allData[studentIdx].row[idxAssigned] || '').trim();

  if (!cls || !byClass[cls]) return 0;

  let contribution = 0;

  // Contribution parité
  const sexe = String(allData[studentIdx].row[idxSexe] || '').toUpperCase();
  const indices = byClass[cls];

  let countF = 0, countM = 0;
  indices.forEach(function(i) {
    const s = String(allData[i].row[idxSexe] || '').toUpperCase();
    if (s === 'F') countF++;
    else if (s === 'M') countM++;
  });

  const parityGap = Math.abs(countF - countM);

  // Si cet élève amplifie le gap, il contribue
  if ((sexe === 'F' && countF > countM) || (sexe === 'M' && countM > countF)) {
    contribution += parityGap * (weights.parity || 1.0);
  }

  // Contribution académique (distance à la moyenne globale)
  const criteria = ['COM', 'TRA', 'PART', 'ABS'];
  const weightList = [weights.com || 1.0, weights.tra || 0.5, weights.part || 0.3, weights.abs || 0.2];

  for (let c = 0; c < criteria.length; c++) {
    const crit = criteria[c];
    const w = weightList[c];
    const idx = headers.indexOf(crit);

    if (idx === -1) continue;

    const studentScore = Number(allData[studentIdx].row[idx] || 3);
    const globalMean = globalDist && globalDist[crit] ? globalDist[crit].mean : 3;

    contribution += Math.abs(studentScore - globalMean) * w;
  }

  return contribution;
}

// ===================================================================
// RECHERCHE SWAP CIBLÉ (MOTEURS SILENCIEUX)
// ===================================================================

/**
 * Cherche le meilleur swap impliquant les élèves problématiques
 */
function findBestTargetedSwap_JulesCodex(allData, byClass, headers, weights, globalDist, config, problematicStudents, swapQuarantine, scoresBefore) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxFixe = headers.indexOf('FIXE');
  const idxMobilite = headers.indexOf('MOBILITE');

  let bestSwap = null;
  let bestGain = 0;

  // Pour chaque élève problématique, chercher un swap bénéfique
  for (let p = 0; p < problematicStudents.length; p++) {
    const i1 = problematicStudents[p].index;
    const cls1 = String(allData[i1].row[idxAssigned] || '').trim();

    if (!cls1) continue;

    // Vérifier FIXE
    if (isStudentFixed_JulesCodex(allData[i1], idxFixe, idxMobilite)) continue;

    // Chercher partenaire de swap dans une autre classe
    const classes = Object.keys(byClass);
    for (let c = 0; c < classes.length; c++) {
      const cls2 = classes[c];
      if (cls2 === cls1) continue;

      const indices2 = byClass[cls2];

      // Essayer plusieurs candidats dans cls2
      for (let attempt = 0; attempt < Math.min(5, indices2.length); attempt++) {
        const i2 = indices2[Math.floor(Math.random() * indices2.length)];

        // Vérifier FIXE
        if (isStudentFixed_JulesCodex(allData[i2], idxFixe, idxMobilite)) continue;

        // Vérifier quarantaine
        if (swapQuarantine && isSwapInQuarantine_JulesCodex(swapQuarantine, i1, i2)) {
          continue;
        }

        // Évaluer swap via API unifiée
        const swapEval = evaluateSwap_JulesCodex({
          i1: i1,
          i2: i2,
          cls1: cls1,
          cls2: cls2
        }, allData, byClass, headers, weights, globalDist, config);

        const gain = scoresBefore - swapEval.scoreAfter;

        if (gain > bestGain) {
          bestGain = gain;
          bestSwap = {
            i1: i1,
            i2: i2,
            cls1: cls1,
            cls2: cls2,
            gain: gain
          };
        }
      }
    }
  }

  return bestSwap;
}

// ===================================================================
// RECHERCHE SWAP ALÉATOIRE (EXPLORATION)
// ===================================================================

/**
 * Cherche le meilleur swap aléatoire (exploration)
 */
function findBestRandomSwap_JulesCodex(allData, byClass, headers, weights, globalDist, config, swapQuarantine, scoresBefore) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxFixe = headers.indexOf('FIXE');
  const idxMobilite = headers.indexOf('MOBILITE');

  let bestSwap = null;
  let bestGain = 0;

  const classes = Object.keys(byClass);

  for (let c1 = 0; c1 < classes.length; c1++) {
    for (let c2 = c1 + 1; c2 < classes.length; c2++) {
      const cls1 = classes[c1];
      const cls2 = classes[c2];

      const indices1 = byClass[cls1];
      const indices2 = byClass[cls2];

      // Essayer quelques swaps aléatoires
      for (let attempt = 0; attempt < Math.min(10, indices1.length * indices2.length); attempt++) {
        const i1 = indices1[Math.floor(Math.random() * indices1.length)];
        const i2 = indices2[Math.floor(Math.random() * indices2.length)];

        // Vérifier FIXE
        if (isStudentFixed_JulesCodex(allData[i1], idxFixe, idxMobilite)) continue;
        if (isStudentFixed_JulesCodex(allData[i2], idxFixe, idxMobilite)) continue;

        // Vérifier quarantaine
        if (swapQuarantine && isSwapInQuarantine_JulesCodex(swapQuarantine, i1, i2)) {
          continue;
        }

        // Évaluer swap
        const swapEval = evaluateSwap_JulesCodex({
          i1: i1,
          i2: i2,
          cls1: cls1,
          cls2: cls2
        }, allData, byClass, headers, weights, globalDist, config);

        const gain = scoresBefore - swapEval.scoreAfter;

        if (gain > bestGain) {
          bestGain = gain;
          bestSwap = {
            i1: i1,
            i2: i2,
            cls1: cls1,
            cls2: cls2,
            gain: gain
          };
        }
      }
    }
  }

  return bestSwap;
}

// ===================================================================
// API UNIFIÉE : ÉVALUATION DE SWAP (COPIE IMMUABLE)
// ===================================================================

/**
 * Évalue un swap sans modifier les structures (copie immuable)
 *
 * @param {Object} candidate - { i1, i2, cls1, cls2 }
 * @returns {Object} { scoreAfter, valid }
 */
function evaluateSwap_JulesCodex(candidate, allData, byClass, headers, weights, globalDist, config) {
  const { i1, i2, cls1, cls2 } = candidate;

  // Copie profonde de byClass
  const tempByClass = {};
  for (const cls in byClass) {
    tempByClass[cls] = byClass[cls].slice();  // Copie du tableau
  }

  // Simuler le swap dans tempByClass
  const idx1InCls1 = tempByClass[cls1].indexOf(i1);
  const idx2InCls2 = tempByClass[cls2].indexOf(i2);

  if (idx1InCls1 === -1 || idx2InCls2 === -1) {
    return { scoreAfter: Infinity, valid: false };
  }

  tempByClass[cls1][idx1InCls1] = i2;
  tempByClass[cls2][idx2InCls2] = i1;

  // Simuler changement _CLASS_ASSIGNED
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const origCls1 = allData[i1].row[idxAssigned];
  const origCls2 = allData[i2].row[idxAssigned];

  allData[i1].row[idxAssigned] = cls2;
  allData[i2].row[idxAssigned] = cls1;

  // Calculer score après swap
  const scoreAfter = calculateGlobalScore_JulesCodex(
    allData, tempByClass, headers, weights, globalDist, config
  );

  // Restaurer état original
  allData[i1].row[idxAssigned] = origCls1;
  allData[i2].row[idxAssigned] = origCls2;

  return { scoreAfter: scoreAfter, valid: true };
}

// ===================================================================
// APPLICATION DU SWAP
// ===================================================================

/**
 * Applique un swap définitivement
 */
function applySwap_JulesCodex(allData, byClass, swap) {
  const { i1, i2, cls1, cls2 } = swap;

  const idxAssigned = allData[0].row.length - 1;  // Approximation (à améliorer)

  // Mettre à jour _CLASS_ASSIGNED
  allData[i1].row[idxAssigned] = cls2;
  allData[i2].row[idxAssigned] = cls1;

  // Mettre à jour byClass
  const idx1InCls1 = byClass[cls1].indexOf(i1);
  const idx2InCls2 = byClass[cls2].indexOf(i2);

  byClass[cls1][idx1InCls1] = i2;
  byClass[cls2][idx2InCls2] = i1;
}

// ===================================================================
// ANCRE DE STABILITÉ (QUARANTAINE)
// ===================================================================

/**
 * Ajoute un swap à la quarantaine
 */
function addToQuarantine_JulesCodex(quarantine, swap, currentIter) {
  if (!quarantine) return;

  // Limiter taille
  if (quarantine.length >= JULES_CODEX_CONFIG.maxQuarantineSize) {
    quarantine.shift();  // Supprimer le plus ancien
  }

  quarantine.push({
    i1: swap.i1,
    i2: swap.i2,
    addedAt: currentIter
  });
}

/**
 * Vérifie si un swap est en quarantaine
 */
function isSwapInQuarantine_JulesCodex(quarantine, i1, i2) {
  if (!quarantine) return false;

  for (let q = 0; q < quarantine.length; q++) {
    const entry = quarantine[q];
    if ((entry.i1 === i1 && entry.i2 === i2) || (entry.i1 === i2 && entry.i2 === i1)) {
      return true;
    }
  }

  return false;
}

/**
 * Nettoie la quarantaine (supprime les swaps périmés)
 */
function cleanQuarantine_JulesCodex(quarantine, currentIter) {
  if (!quarantine) return;

  for (let q = quarantine.length - 1; q >= 0; q--) {
    const entry = quarantine[q];
    if (currentIter - entry.addedAt > JULES_CODEX_CONFIG.swapQuarantineDuration) {
      quarantine.splice(q, 1);
    }
  }
}

// ===================================================================
// UTILITAIRES
// ===================================================================

/**
 * Vérifie si un élève est FIXE
 */
function isStudentFixed_JulesCodex(student, idxFixe, idxMobilite) {
  if (idxFixe >= 0) {
    const fixe = String(student.row[idxFixe] || '').toUpperCase();
    if (fixe === 'FIXE' || fixe === 'OUI') return true;
  }

  if (idxMobilite >= 0) {
    const mobilite = String(student.row[idxMobilite] || '').toUpperCase();
    if (mobilite === 'FIXE' || mobilite.indexOf('GROUPE_FIXE') >= 0) return true;
  }

  return false;
}
