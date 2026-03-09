/**
 * ===================================================================
 * 🚀 OPTIMUM PRIME ULTIMATE - MOTEUR D'OPTIMISATION
 * ===================================================================
 * LE CONCEPT :
 * Remplace l'usine à gaz "V15" par une architecture saine.
 * Combine la puissance statistique de "Jules Codex" avec les
 * critères pédagogiques "Humains" (Têtes de classe, Niv1).
 *
 * PONDÉRATION ASYMÉTRIQUE DES EXTRÊMES.
 *
 * PIPELINE : LEGACY uniquement (SCORE CONSOLE → Admin).
 * Le pipeline OPTI (SCORE INTERFACE → Prof) utilise
 * Phase4_balanceScoresSwaps_BASEOPTI_V3 (Phases_BASEOPTI_V3_COMPLETE.js).
 *
 * AUTEUR : Gemini (Expert Apps Script)
 * DATE : 19/11/2025
 * ===================================================================
 */

// U3 REPLICANT : Configuration par défaut, surchargeable via ctx.ultimateConfig
const ULTIMATE_CONFIG_DEFAULTS = {
  maxSwaps: 2000,
  stagnationLimit: 50,
  weights: {
    distrib: 5.0,
    parity: 4.0,
    profiles: 10.0,
    friends: 1000.0
  },
  targets: {
    headMin: 2,
    headMax: 5,
    niv1Max: 4,
    niv1Min: 0
  },
  // Recuit Simulé (Simulated Annealing) — permet de sortir des optima locaux
  // en acceptant ponctuellement des swaps légèrement dégradants avec une
  // probabilité décroissante (température qui refroidit).
  sa: {
    enabled: true,          // Activer le recuit simulé
    initialTemp: 50.0,      // Température initiale T₀ (échelle du score)
    coolingRate: 0.995,      // Facteur de refroidissement géométrique par itération
    minTemp: 0.1,           // Température plancher (en dessous = glouton pur)
    maxDegradation: 200.0   // Gain négatif max toléré (sécurité anti-dégradation)
  }
};

// Config live fusionnée avec les overrides du contexte
// Permet de modifier les poids depuis l'UI sans toucher au code
function getUltimateConfig_(ctx) {
  var cfg = JSON.parse(JSON.stringify(ULTIMATE_CONFIG_DEFAULTS));
  var overrides = (ctx && ctx.ultimateConfig) || {};
  if (overrides.maxSwaps !== undefined) cfg.maxSwaps = overrides.maxSwaps;
  if (overrides.stagnationLimit !== undefined) cfg.stagnationLimit = overrides.stagnationLimit;
  if (overrides.weights) {
    for (var k in overrides.weights) {
      cfg.weights[k] = overrides.weights[k];
    }
  }
  if (overrides.targets) {
    for (var k in overrides.targets) {
      cfg.targets[k] = overrides.targets[k];
    }
  }
  if (overrides.sa) {
    for (var k in overrides.sa) {
      cfg.sa[k] = overrides.sa[k];
    }
  }
  return cfg;
}

// Compat : ULTIMATE_CONFIG reste accessible globalement comme alias
var ULTIMATE_CONFIG = ULTIMATE_CONFIG_DEFAULTS;

/**
 * Point d'entrée principal appelé par le Pipeline OPTI ou LEGACY
 *
 * MULTI-RESTART VERSION :
 * - U1: Sampling priorisé par disruption (top candidats, pas brute-force)
 * - U2: PART dans la complémentarité du partenaire
 * - U3: Config dynamique fusionnée depuis ctx.ultimateConfig
 * - N3: PRNG seedable pour reproductibilité
 * - MR: Multi-restart (5 seeds, garde le meilleur résultat)
 * - SA: Recuit Simulé — accepte ponctuellement des swaps dégradants
 *       avec probabilité e^(gain/T), T décroissante. Permet de sortir
 *       des optima locaux avant de reconverger en glouton pur.
 *
 * @param {Object} ctx - Contexte de l'optimisation
 * @returns {Object} Résultat d'optimisation
 */
function Phase4_Ultimate_Run(ctx) {
  logLine('INFO', '🔧 [engine=LEGACY] Phase4_Ultimate_Run (SCORE CONSOLE / Admin)');
  const ss = ctx.ss || SpreadsheetApp.getActiveSpreadsheet();
  const config = getUltimateConfig_(ctx);
  const mrConfig = MULTI_RESTART_CONFIG;
  const maxRestarts = mrConfig.maxRestarts;

  logLine('INFO', '🚀 Lancement OPTIMUM PRIME ULTIMATE (MULTI-RESTART)...');
  logLine('INFO', '⚖️ Config: ' + JSON.stringify(config.weights));
  logLine('INFO', `🔁 Multi-restart : ${maxRestarts} seeds`);
  if (config.sa && config.sa.enabled) {
    logLine('INFO', `🌡️ Recuit Simulé ACTIF : T₀=${config.sa.initialTemp}, cooling=${config.sa.coolingRate}, Tmin=${config.sa.minTemp}`);
  }

  // 1. CHARGEMENT ET CLASSIFICATION (une seule fois)
  const dataResult = loadAndClassifyData_Ultimate(ctx);
  if (!dataResult.ok) {
    logLine('ERROR', '❌ Échec du chargement des données ULTIMATE');
    return { ok: false, message: 'Erreur chargement données' };
  }

  const { allData, byClass: originalByClass, headers } = dataResult;
  logLine('INFO', `📊 Chargement OK : ${allData.length} élèves répartis en ${Object.keys(originalByClass).length} classes.`);

  // 2. STATISTIQUES GLOBALES (invariantes entre restarts)
  const globalStats = calculateGlobalStats_Ultimate(allData);
  logLine('INFO', `🎯 Cibles : Ratio F=${(globalStats.ratioF*100).toFixed(1)}%, Moyenne COM=${globalStats.avgCOM.toFixed(2)}, PART=${globalStats.avgPART.toFixed(2)}, ABS=${globalStats.avgABS.toFixed(2)}`);

  // ===== MULTI-RESTART LOOP =====
  let bestByClass = null;
  let bestScore = Infinity;
  let bestSwaps = 0;
  let bestSwaps3Way = 0;
  let bestSeed = 0;

  for (let restart = 0; restart < maxRestarts; restart++) {
    const seed = ctx.seed ? ctx.seed + restart * mrConfig.seedSpacing : restart * mrConfig.seedSpacing;
    const rng = createRNG(seed);

    // Copie indépendante de byClass pour ce restart
    const byClass = snapshotByClass_(originalByClass);

    logLine('INFO', `  🔁 Restart ${restart + 1}/${maxRestarts} (seed=${seed})`);

    // Exécuter le coeur du moteur
    const result = runPhase4CoreLoop_Ultimate_(allData, byClass, headers, globalStats, ctx, config, rng);

    // Calculer le score global pour comparer
    let totalScore = 0;
    for (const cls in byClass) {
      totalScore += calculateScore_Ultimate(byClass[cls], allData, globalStats, cls, ctx, config);
    }

    logLine('INFO', `    📊 Score=${totalScore.toFixed(2)}, swaps=${result.swapsApplied}+${result.swaps3Way}(3-way)`);

    // PROP3 : Valider DISSO PAR RESTART — rejeter tout restart invalide
    const restartValidation = validateDISSOConstraints_Ultimate(allData, byClass, headers);
    if (!restartValidation.ok) {
      logLine('WARN', `    ❌ Restart ${restart + 1} rejeté : DISSO invalide (${restartValidation.duplicates.length} conflit(s))`);
      restartValidation.duplicates.forEach(dup => {
        logLine('WARN', `      • ${dup.classe} : ${dup.code} x${dup.count} (${dup.noms.join(', ')})`);
      });
      continue; // On ne considère JAMAIS ce restart, même si son score est meilleur
    }

    if (totalScore < bestScore) {
      bestScore = totalScore;
      bestByClass = snapshotByClass_(byClass);
      bestSwaps = result.swapsApplied;
      bestSwaps3Way = result.swaps3Way;
      bestSeed = seed;
      logLine('INFO', `    ⭐ Nouveau meilleur ! (score=${bestScore.toFixed(2)}, DISSO ✅)`);
    }
  }

  // Si aucun restart n'a passé la validation DISSO, log explicite
  if (!bestByClass) {
    logLine('ERROR', '❌ AUCUN restart valide (tous rejetés par DISSO). Fallback sur l\'état pré-Phase4.');
    return {
      ok: false,
      swapsApplied: 0,
      swaps3Way: 0,
      seed: 0,
      restarts: maxRestarts,
      saveResult: { ok: false },
      validation: { ok: false, duplicates: [], message: 'Tous les restarts rejetés par DISSO' }
    };
  }

  logLine('INFO', `📊 Meilleur restart : seed=${bestSeed}, score=${bestScore.toFixed(2)}, swaps=${bestSwaps}+${bestSwaps3Way}(3-way)`);

  // 4. SAUVEGARDE DU MEILLEUR RÉSULTAT
  const saveResult = saveResults_Ultimate(ss, allData, bestByClass, headers);

  // 5. VALIDATION FINALE (filet de sécurité — le restart sélectionné a déjà passé DISSO)
  const validationResult = validateDISSOConstraints_Ultimate(allData, bestByClass, headers);
  if (!validationResult.ok) {
    logLine('ERROR', '❌ VALIDATION DISSO ÉCHOUÉE après Phase 4 ULTIMATE (incohérence interne !)');
    validationResult.duplicates.forEach(dup => {
      logLine('ERROR', `    • ${dup.classe} : ${dup.code} présent ${dup.count} fois (${dup.noms.join(', ')})`);
    });
  } else {
    logLine('INFO', '✅ Validation DISSO finale : confirmée');
  }

  logLine('SUCCESS', `✅ ULTIMATE MULTI-RESTART Terminé : meilleur sur ${maxRestarts} seeds. Seed gagnant: ${bestSeed}`);
  return {
    ok: true,
    swapsApplied: bestSwaps + bestSwaps3Way,
    swaps3Way: bestSwaps3Way,
    seed: bestSeed,
    restarts: maxRestarts,
    saveResult: saveResult,
    validation: validationResult
  };
}

/**
 * Coeur du moteur Phase 4 Ultimate.
 * Travaille uniquement en mémoire (byClass) — pas d'I/O sheet.
 * Appelé N fois par le multi-restart avec des seeds différentes.
 *
 * @returns {{ swapsApplied, swaps3Way }}
 */
function runPhase4CoreLoop_Ultimate_(allData, byClass, headers, globalStats, ctx, config, rng) {
  let swapsApplied = 0;
  let stagnationCount = 0;
  let noPartnerCount = 0; // TWO-PIPELINE : compteur séparé pour les échecs de partner

  // TWO-PIPELINE : Swap history anti-boucle (porté depuis V3)
  // Pénalise les élèves déjà swappés pour éviter les cycles A↔B↔A
  const swapHistory = new Map();

  // --- RECUIT SIMULÉ (Simulated Annealing) ---
  const saEnabled = config.sa && config.sa.enabled;
  let temperature = saEnabled ? config.sa.initialTemp : 0;
  const coolingRate = saEnabled ? config.sa.coolingRate : 1;
  const minTemp = saEnabled ? config.sa.minTemp : 0;
  const maxDegradation = saEnabled ? config.sa.maxDegradation : 0;
  let saAccepted = 0; // Compteur de swaps dégradants acceptés par SA

  for (let iter = 0; iter < config.maxSwaps; iter++) {
    const worstClassKey = findWorstClass_Ultimate(byClass, allData, globalStats, ctx, config);
    if (!worstClassKey) break;

    const partnerClassKey = findPartnerClass_Ultimate(worstClassKey, byClass, allData, globalStats, rng, config);
    if (!partnerClassKey) {
      noPartnerCount++;
      // TWO-PIPELINE : ne pas casser la boucle trop tôt sur les échecs partner,
      // le RNG 20% peut échouer temporairement sans que ce soit une stagnation réelle
      if (noPartnerCount > 30) break;
      continue;
    }
    noPartnerCount = 0;

    const bestSwap = findBestSwapPrioritized_Ultimate(worstClassKey, partnerClassKey, allData, byClass, headers, globalStats, ctx, rng, config);

    // TWO-PIPELINE : Appliquer la pénalité swap history sur le gain
    if (bestSwap) {
      const h1 = swapHistory.get(bestSwap.idx1) || 0;
      const h2 = swapHistory.get(bestSwap.idx2) || 0;
      bestSwap.gain = bestSwap.gain / (1 + h1 + h2);
    }

    if (bestSwap && bestSwap.gain > 0.0001) {
      // Swap améliorant → toujours accepté (comportement glouton classique)
      applySwap_Ultimate(allData, byClass, bestSwap, headers);
      swapHistory.set(bestSwap.idx1, (swapHistory.get(bestSwap.idx1) || 0) + 1);
      swapHistory.set(bestSwap.idx2, (swapHistory.get(bestSwap.idx2) || 0) + 1);
      swapsApplied++;
      stagnationCount = 0;
    } else if (saEnabled && temperature > minTemp && bestSwap && bestSwap.gain < 0 && bestSwap.gain > -maxDegradation) {
      // RECUIT SIMULÉ : swap légèrement dégradant, accepter avec probabilité e^(gain/T)
      // gain est négatif ici, donc la probabilité ∈ (0, 1) et décroît quand T diminue
      const acceptProbability = Math.exp(bestSwap.gain / temperature);
      if (rng.next() < acceptProbability) {
        applySwap_Ultimate(allData, byClass, bestSwap, headers);
        swapHistory.set(bestSwap.idx1, (swapHistory.get(bestSwap.idx1) || 0) + 1);
        swapHistory.set(bestSwap.idx2, (swapHistory.get(bestSwap.idx2) || 0) + 1);
        swapsApplied++;
        saAccepted++;
        stagnationCount = 0;
        if (saAccepted <= 5 || saAccepted % 10 === 0) {
          logLine('DEBUG', `  🌡️ SA: swap dégradant accepté (gain=${bestSwap.gain.toFixed(4)}, T=${temperature.toFixed(2)}, p=${acceptProbability.toFixed(4)})`);
        }
      } else {
        stagnationCount++;
      }
    } else {
      stagnationCount++;
    }

    // Refroidissement géométrique de la température
    if (saEnabled) {
      temperature *= coolingRate;
    }

    if (stagnationCount >= config.stagnationLimit) break;
  }

  if (saEnabled) {
    logLine('INFO', `  🌡️ Recuit Simulé: ${saAccepted} swaps dégradants acceptés, T finale=${temperature.toFixed(4)}`);
  }

  // Après la phase SA+greedy, relancer une passe glouton pure pour converger
  // (le SA a pu déplacer la solution dans un nouveau bassin d'attraction)
  if (saEnabled && saAccepted > 0) {
    let postSASwaps = 0;
    let postStagnation = 0;
    for (let iter2 = 0; iter2 < Math.floor(config.maxSwaps * 0.3); iter2++) {
      const worstKey = findWorstClass_Ultimate(byClass, allData, globalStats, ctx, config);
      if (!worstKey) break;
      const partnerKey = findPartnerClass_Ultimate(worstKey, byClass, allData, globalStats, rng, config);
      if (!partnerKey) { postStagnation++; if (postStagnation > 10) break; continue; }
      const swap = findBestSwapPrioritized_Ultimate(worstKey, partnerKey, allData, byClass, headers, globalStats, ctx, rng, config);
      if (swap && swap.gain > 0.0001) {
        applySwap_Ultimate(allData, byClass, swap, headers);
        postSASwaps++;
        swapsApplied++;
        postStagnation = 0;
      } else {
        postStagnation++;
      }
      if (postStagnation >= config.stagnationLimit) break;
    }
    if (postSASwaps > 0) {
      logLine('INFO', `  🎯 Post-SA greedy: ${postSASwaps} swaps supplémentaires`);
    }
  }

  // 3-WAY CYCLE SWAPS
  let swaps3Way = 0;
  const classNames = Object.keys(byClass);

  for (let iter3 = 0; iter3 < 200; iter3++) {
    let bestGain3 = 0.001;
    let best3Way = null;

    for (let t = 0; t < 15; t++) {
      const c1 = rng.pick(classNames);
      const c2 = rng.pick(classNames);
      const c3 = rng.pick(classNames);
      if (c1 === c2 || c2 === c3 || c1 === c3) continue;
      if (!byClass[c1].length || !byClass[c2].length || !byClass[c3].length) continue;

      const scoreBefore3 = calculateScore_Ultimate(byClass[c1], allData, globalStats, c1, ctx, config) +
                           calculateScore_Ultimate(byClass[c2], allData, globalStats, c2, ctx, config) +
                           calculateScore_Ultimate(byClass[c3], allData, globalStats, c3, ctx, config);

      for (let s = 0; s < 10; s++) {
        const a = rng.pick(byClass[c1]);
        const b = rng.pick(byClass[c2]);
        const c = rng.pick(byClass[c3]);
        if (isFixed(allData[a]) || isFixed(allData[b]) || isFixed(allData[c])) continue;

        // 3-WAY CYCLE : a:c1→c2, b:c2→c3, c:c3→c1
        // Chaque check bilatéral vérifie les deux directions du swap.
        // Les 3 checks couvrent les 3 mouvements réels (+ 3 checks conservateurs).
        // Sans le 3e check, c→c1 n'était JAMAIS validé (bug critique DISSO/LV2).
        if (!canSwapStudents_Ultimate(a, b, c1, c2, byClass[c1], byClass[c2], allData, headers, ctx)) continue;
        if (!canSwapStudents_Ultimate(b, c, c2, c3, byClass[c2], byClass[c3], allData, headers, ctx)) continue;
        if (!canSwapStudents_Ultimate(c, a, c3, c1, byClass[c3], byClass[c1], allData, headers, ctx)) continue;

        const tempC1 = byClass[c1].filter(x => x !== a).concat([c]);
        const tempC2 = byClass[c2].filter(x => x !== b).concat([a]);
        const tempC3 = byClass[c3].filter(x => x !== c).concat([b]);

        const scoreAfter3 = calculateScore_Ultimate(tempC1, allData, globalStats, c1, ctx, config) +
                            calculateScore_Ultimate(tempC2, allData, globalStats, c2, ctx, config) +
                            calculateScore_Ultimate(tempC3, allData, globalStats, c3, ctx, config);

        const gain3 = scoreBefore3 - scoreAfter3;
        if (gain3 > bestGain3) {
          bestGain3 = gain3;
          best3Way = { a, b, c, c1, c2, c3 };
        }
      }
    }

    if (!best3Way) break;

    const { a, b, c, c1, c2, c3 } = best3Way;
    byClass[c1] = byClass[c1].filter(x => x !== a).concat([c]);
    byClass[c2] = byClass[c2].filter(x => x !== b).concat([a]);
    byClass[c3] = byClass[c3].filter(x => x !== c).concat([b]);
    swaps3Way++;
    swapsApplied++;
  }

  return { swapsApplied: swapsApplied, swaps3Way: swaps3Way };
}

// ===================================================================
// 🧠 LE CERVEAU : CALCUL DES SCORES (La logique pédagogique)
// ===================================================================

/**
 * Calcule le score de "maladie" d'une classe
 * PONDÉRATION ASYMÉTRIQUE DES EXTRÊMES :
 * - Pénalité forte (au carré) si manque de têtes
 * - Pénalité modérée si excès de têtes
 * - Pénalité très forte (au cube) si excès de Niv1
 * ✅ BUG #4 CORRECTION : Ajout critère d'effectif
 */
function calculateScore_Ultimate(indices, allData, globalStats, className, ctx, config) {
  config = config || ULTIMATE_CONFIG; // fallback sécurité
  let score = 0;
  const students = indices.map(i => allData[i]);
  const total = students.length;
  if (total === 0) return 10000;

  // --- 0. CRITÈRE EFFECTIF (BUG #4 CORRECTION - PRIORITÉ HAUTE) ---
  if (className && ctx && ctx.targets && ctx.targets[className]) {
    const targetSize = ctx.targets[className];
    const sizeDiff = total - targetSize;
    // Pénalité quadratique pour les écarts d'effectif
    score += Math.pow(sizeDiff, 2) * 800;
  }

  // --- 1. CRITÈRE PROFILS (Héritage LEGACY - Priorité Absolue) ---
  const nbTetes = students.filter(s => s.isHead).length;
  const nbNiv1 = students.filter(s => s.isNiv1).length;

  // PONDÉRATION ASYMÉTRIQUE DES EXTRÊMES
  if (nbTetes < config.targets.headMin) {
    score += Math.pow(config.targets.headMin - nbTetes, 2) * 500;
  }
  if (nbTetes > config.targets.headMax) {
    score += (nbTetes - config.targets.headMax) * 200;
  }

  if (nbNiv1 > config.targets.niv1Max) {
    score += Math.pow(nbNiv1 - config.targets.niv1Max, 3) * 100;
  }

  // --- 2. CRITÈRE PARITÉ (Adaptatif) ---
  const nbFilles = students.filter(s => s.sexe === 'F').length;
  const ratioF = nbFilles / total;
  score += Math.abs(ratioF - globalStats.ratioF) * 1000 * config.weights.parity;

  // --- 3. CRITÈRE DISTRIBUTION ACADÉMIQUE (Jules Codex) ---
  // ✅ FIX #2 : Inclure ABS dans le scoring (était absent malgré le commentaire)
  const avgCOM = students.reduce((acc, s) => acc + (s.COM || 2), 0) / total;
  const avgTRA = students.reduce((acc, s) => acc + (s.TRA || 2), 0) / total;
  const avgPART = students.reduce((acc, s) => acc + (s.PART || 2), 0) / total;
  const avgABS = students.reduce((acc, s) => acc + (s.ABS || 2), 0) / total;

  score += Math.abs(avgCOM - globalStats.avgCOM) * 100 * config.weights.distrib;
  score += Math.abs(avgTRA - globalStats.avgTRA) * 100 * config.weights.distrib;
  score += Math.abs(avgPART - (globalStats.avgPART || 2)) * 50 * config.weights.distrib;
  score += Math.abs(avgABS - (globalStats.avgABS || 2)) * 50 * config.weights.distrib;

  return score;
}

/**
 * U1 REPLICANT : Sampling priorisé par disruption.
 * Au lieu de brute-force random, on sélectionne les élèves les plus "déplacés"
 * dans chaque classe et on ne teste que ceux-là.
 */
function findBestSwapPrioritized_Ultimate(cls1Name, cls2Name, allData, byClass, headers, globalStats, ctx, rng, config) {
  const idxList1 = byClass[cls1Name];
  const idxList2 = byClass[cls2Name];

  // Calculer les profils cibles
  const avgCls1 = allData.filter((s, i) => idxList1.indexOf(i) >= 0).reduce((acc, s) => acc + s.COM, 0) / idxList1.length;
  const avgCls2 = allData.filter((s, i) => idxList2.indexOf(i) >= 0).reduce((acc, s) => acc + s.TRA, 0) / idxList2.length;

  // Trier par disruption (distance au profil moyen de leur classe)
  function sortByDisruption(indices) {
    const total = indices.length;
    const students = indices.map(i => allData[i]);
    const avgCOM = students.reduce((s, st) => s + st.COM, 0) / total;
    const avgTRA = students.reduce((s, st) => s + st.TRA, 0) / total;
    const avgPART = students.reduce((s, st) => s + (st.PART || 2), 0) / total;

    var avgABS = students.reduce((s, st) => s + (st.ABS || 2), 0) / total;

    return indices.slice().sort(function(a, b) {
      var distA = Math.abs(allData[a].COM - avgCOM) + Math.abs(allData[a].TRA - avgTRA) + Math.abs((allData[a].PART || 2) - avgPART) * 0.5 + Math.abs((allData[a].ABS || 2) - avgABS) * 0.5;
      var distB = Math.abs(allData[b].COM - avgCOM) + Math.abs(allData[b].TRA - avgTRA) + Math.abs((allData[b].PART || 2) - avgPART) * 0.5 + Math.abs((allData[b].ABS || 2) - avgABS) * 0.5;
      return distB - distA; // Plus perturbant en premier
    }).filter(i => !isFixed(allData[i]));
  }

  // Prendre les top 60% les plus perturbants dans chaque classe
  const sorted1 = sortByDisruption(idxList1);
  const sorted2 = sortByDisruption(idxList2);
  const topCount1 = Math.max(5, Math.ceil(sorted1.length * 0.6));
  const topCount2 = Math.max(5, Math.ceil(sorted2.length * 0.6));
  const candidates1 = sorted1.slice(0, topCount1);
  const candidates2 = sorted2.slice(0, topCount2);

  const scoreBefore = calculateScore_Ultimate(idxList1, allData, globalStats, cls1Name, ctx, config) +
                      calculateScore_Ultimate(idxList2, allData, globalStats, cls2Name, ctx, config);

  let bestSwap = null;
  let maxGain = 0;

  // SA : garder aussi le "moins pire" swap dégradant pour le recuit simulé
  let leastBadSwap = null;
  let leastBadGain = -Infinity;

  // Tester les paires priorisées (top candidats seulement)
  for (let i = 0; i < candidates1.length; i++) {
    const i1 = candidates1[i];
    const s1 = allData[i1];

    for (let j = 0; j < candidates2.length; j++) {
      const i2 = candidates2[j];
      const s2 = allData[i2];

      // ✅ BUG #5 CORRECTION : Vérifier compatibilité LV2/OPT AVANT swap
      if (!canSwapStudents_Ultimate(i1, i2, cls1Name, cls2Name, idxList1, idxList2, allData, headers, ctx)) {
        continue; // Swap interdit par contraintes LV2/OPT/DISSO
      }

      // Simulation du swap
      const tempList1 = idxList1.filter(idx => idx !== i1).concat([i2]);
      const tempList2 = idxList2.filter(idx => idx !== i2).concat([i1]);

      const scoreAfter = calculateScore_Ultimate(tempList1, allData, globalStats, cls1Name, ctx, config) +
                         calculateScore_Ultimate(tempList2, allData, globalStats, cls2Name, ctx, config);

      const gain = scoreBefore - scoreAfter;

      if (gain > maxGain) {
        maxGain = gain;
        bestSwap = {
          idx1: i1,
          idx2: i2,
          cls1: cls1Name,
          cls2: cls2Name,
          gain: gain,
          reason: `Swap ${s1.isHead ? 'Tête' : 'Std'}/${s1.isNiv1 ? 'Niv1' : 'Std'}`
        };
      } else if (gain < 0 && gain > leastBadGain) {
        // SA : meilleur candidat dégradant (gain négatif le plus proche de 0)
        leastBadGain = gain;
        leastBadSwap = {
          idx1: i1,
          idx2: i2,
          cls1: cls1Name,
          cls2: cls2Name,
          gain: gain,
          reason: `SA-Swap ${s1.isHead ? 'Tête' : 'Std'}/${s1.isNiv1 ? 'Niv1' : 'Std'}`
        };
      }
    }
  }

  // Si aucun swap améliorant trouvé, renvoyer le moins dégradant pour le recuit simulé
  return bestSwap || leastBadSwap;
}

/**
 * Charge et classifie toutes les données élèves
 */
function loadAndClassifyData_Ultimate(ctx) {
  const ss = ctx.ss || SpreadsheetApp.getActiveSpreadsheet();
  const allData = [];
  const byClass = {};
  let headersRef = null;
  
  // 🌟 APPROCHE UNIVERSELLE : Détecter LV2 universelles
  const nbClasses = (ctx.niveaux || []).length;
  const lv2Counts = {};
  
  for (const classe in (ctx.quotas || {})) {
    const quotas = ctx.quotas[classe];
    for (const optName in quotas) {
      if (isKnownLV2(optName) && quotas[optName] > 0) {
        lv2Counts[optName] = (lv2Counts[optName] || 0) + 1;
      }
    }
  }
  
  const lv2Universelles = [];
  for (const lv2 in lv2Counts) {
    if (lv2Counts[lv2] === nbClasses) {
      lv2Universelles.push(lv2);
    }
  }
  
  // Ajouter au contexte
  ctx.lv2Universelles = lv2Universelles;

  // ✅ CORRECTION CRITIQUE : Lire UNIQUEMENT depuis les onglets TEST
  //    qui contiennent le résultat des Phases 1-2-3, PAS depuis les sources
  const testSheets = (ctx.cacheSheets || []).map(name => ss.getSheetByName(name)).filter(s => s);

  testSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    if (!headersRef) headersRef = data[0];
    const headers = data[0];

    const idx = {
      ID: headers.indexOf('ID_ELEVE'),
      SEXE: headers.indexOf('SEXE'),
      COM: headers.indexOf('COM'),
      TRA: headers.indexOf('TRA'),
      PART: headers.indexOf('PART'),
      ABS: headers.indexOf('ABSENCE') !== -1 ? headers.indexOf('ABSENCE') : headers.indexOf('ABS'),
      MOB: headers.indexOf('MOBILITE'),
      FIXE: headers.indexOf('FIXE')
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idx.ID]) continue;

      const student = {
        row: row,
        originalSheet: sheet.getName(),
        index: i,
        sexe: String(row[idx.SEXE] || 'M').toUpperCase().trim().charAt(0),
        COM: Number(row[idx.COM]) || 2,
        TRA: Number(row[idx.TRA]) || 2,
        PART: Number(row[idx.PART]) || 2,
        ABS: idx.ABS >= 0 ? (Number(row[idx.ABS]) || 2) : 2,
        mobilite: String(row[idx.MOB] || row[idx.FIXE] || '').toUpperCase()
      };

      // --- CLASSIFICATION LOGIQUE ---
      const scoreMoy = (student.COM + student.TRA + student.PART) / 3;
      student.isHead = (student.COM >= 4 || student.TRA >= 4) || scoreMoy >= 3.5;
      student.isNiv1 = (student.COM <= 1 || student.TRA <= 1);

      allData.push(student);

      // ✅ CORRECTION : Extraire le nom de classe depuis le nom de l'onglet TEST
      //    Ex: "5°1TEST" → "5°1"
      const sheetName = sheet.getName();
      const className = sheetName.replace(/TEST$/i, '');
      
      if (!byClass[className]) byClass[className] = [];
      byClass[className].push(allData.length - 1);
    }
  });

  return { ok: true, allData: allData, byClass: byClass, headers: headersRef };
}

/**
 * Calcule les statistiques globales
 */
function calculateGlobalStats_Ultimate(allData) {
  let total = allData.length;
  if (total === 0) return { ratioF: 0.5, avgCOM: 2.5, avgTRA: 2.5, avgPART: 2.5, avgABS: 2.5 };

  const nbFilles = allData.filter(s => s.sexe === 'F').length;
  const sumCOM = allData.reduce((sum, s) => sum + s.COM, 0);
  const sumTRA = allData.reduce((sum, s) => sum + s.TRA, 0);
  const sumPART = allData.reduce((sum, s) => sum + (s.PART || 2), 0);
  // ✅ FIX #2 : Calculer avgABS (manquait)
  const sumABS = allData.reduce((sum, s) => sum + (s.ABS || 2), 0);

  return {
    ratioF: nbFilles / total,
    avgCOM: sumCOM / total,
    avgTRA: sumTRA / total,
    avgPART: sumPART / total,
    avgABS: sumABS / total
  };
}

/**
 * Identifie la classe "malade" (score le plus élevé)
 */
function findWorstClass_Ultimate(byClass, allData, globalStats, ctx, config) {
  config = config || ULTIMATE_CONFIG; // fallback sécurité
  let maxScore = -1;
  let worstClass = null;
  for (const cls in byClass) {
    const score = calculateScore_Ultimate(byClass[cls], allData, globalStats, cls, ctx, config);
    if (score > maxScore) {
      maxScore = score;
      worstClass = cls;
    }
  }
  return worstClass;
}

/**
 * Trouve la meilleure classe partenaire pour un swap.
 * REPLICANT (F2+U2) : Complémentarité enrichie avec PART + RNG seedable.
 */
function findPartnerClass_Ultimate(worstClass, byClass, allData, globalStats, rng, config) {
  config = config || ULTIMATE_CONFIG; // fallback sécurité
  var _rng = rng || { next: Math.random, pick: function(a) { return a[Math.floor(Math.random() * a.length)]; } };
  const classes = Object.keys(byClass).filter(c => c !== worstClass);
  if (classes.length === 0) return null;

  const worstStudents = byClass[worstClass].map(i => allData[i]);
  const worstTotal = worstStudents.length;
  if (worstTotal === 0) return null;

  const worstNbTetes = worstStudents.filter(s => s.isHead).length;
  const worstNbNiv1 = worstStudents.filter(s => s.isNiv1).length;
  const worstRatioF = worstStudents.filter(s => s.sexe === 'F').length / worstTotal;
  const worstAvgCOM = worstStudents.reduce((s, st) => s + st.COM, 0) / worstTotal;
  // U2: Ajouter PART à la complémentarité
  const worstAvgPART = worstStudents.reduce((s, st) => s + (st.PART || 2), 0) / worstTotal;

  let bestPartner = null;
  let bestComplementarity = -Infinity;

  for (let c = 0; c < classes.length; c++) {
    const cls = classes[c];
    const clsStudents = byClass[cls].map(i => allData[i]);
    const clsTotal = clsStudents.length;
    if (clsTotal === 0) continue;

    const clsNbTetes = clsStudents.filter(s => s.isHead).length;
    const clsNbNiv1 = clsStudents.filter(s => s.isNiv1).length;
    const clsRatioF = clsStudents.filter(s => s.sexe === 'F').length / clsTotal;
    const clsAvgCOM = clsStudents.reduce((s, st) => s + st.COM, 0) / clsTotal;
    const clsAvgPART = clsStudents.reduce((s, st) => s + (st.PART || 2), 0) / clsTotal;

    let comp = 0;

    // Têtes de classe croisées
    const teteDiff = (worstNbTetes - config.targets.headMin) - (clsNbTetes - config.targets.headMin);
    comp += Math.abs(teteDiff) * 3;

    // Niv1 croisés
    const niv1Diff = (worstNbNiv1 - config.targets.niv1Max) - (clsNbNiv1 - config.targets.niv1Max);
    comp += Math.abs(niv1Diff) * 3;

    // Parité complémentaire
    if ((worstRatioF > globalStats.ratioF && clsRatioF < globalStats.ratioF) ||
        (worstRatioF < globalStats.ratioF && clsRatioF > globalStats.ratioF)) {
      comp += 2;
    }

    // COM complémentaire
    if ((worstAvgCOM > globalStats.avgCOM && clsAvgCOM < globalStats.avgCOM) ||
        (worstAvgCOM < globalStats.avgCOM && clsAvgCOM > globalStats.avgCOM)) {
      comp += Math.abs(worstAvgCOM - clsAvgCOM) * 2;
    }

    // U2: PART complémentaire
    if ((worstAvgPART > (globalStats.avgPART || 2) && clsAvgPART < (globalStats.avgPART || 2)) ||
        (worstAvgPART < (globalStats.avgPART || 2) && clsAvgPART > (globalStats.avgPART || 2))) {
      comp += Math.abs(worstAvgPART - clsAvgPART) * 1.5;
    }

    if (comp > bestComplementarity) {
      bestComplementarity = comp;
      bestPartner = cls;
    }
  }

  // 20% du temps : exploration aléatoire via RNG seedable
  if (_rng.next() < 0.2) {
    return _rng.pick(classes);
  }

  return bestPartner;
}

/**
 * Vérifie si un élève est "fixe" (non mobile)
 */
function isFixed(student) {
  const mob = student.mobilite;
  return mob.includes('FIXE') || mob.includes('NON');
}

/**
 * ✅ BUG #5 CORRECTION : Vérifie si un swap respecte les contraintes LV2/OPT/DISSO
 */
function canSwapStudents_Ultimate(idx1, idx2, cls1Name, cls2Name, idxList1, idxList2, allData, headers, ctx) {
  const s1 = allData[idx1];
  const s2 = allData[idx2];

  // Extraire LV2/OPT/ASSO des élèves
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');
  const idxDISSO = headers.indexOf('DISSO');
  const idxASSO = headers.indexOf('ASSO');

  // HARMONY FIX : Vérifier ASSO - ne jamais séparer un groupe ASSO
  if (idxASSO >= 0) {
    const asso_s1 = String(s1.row[idxASSO] || '').trim().toUpperCase();
    const asso_s2 = String(s2.row[idxASSO] || '').trim().toUpperCase();

    if (asso_s1) {
      // s1 fait partie d'un groupe ASSO, vérifier que ses pairs sont dans cls2
      const pairsInCls1 = idxList1.filter(function(idx) {
        if (idx === idx1) return false;
        return String(allData[idx].row[idxASSO] || '').trim().toUpperCase() === asso_s1;
      });
      // Si des pairs restent dans cls1, on ne peut pas swapper s1 ailleurs
      if (pairsInCls1.length > 0) return false;
    }

    if (asso_s2) {
      const pairsInCls2 = idxList2.filter(function(idx) {
        if (idx === idx2) return false;
        return String(allData[idx].row[idxASSO] || '').trim().toUpperCase() === asso_s2;
      });
      if (pairsInCls2.length > 0) return false;
    }
  }

  // ✅ SAFETY CHECK: Vérifier que les colonnes critiques existent
  if (idxDISSO === -1) {
    logLine('ERROR', '❌ CRITIQUE: Colonne DISSO non trouvée dans les headers! Headers: ' + headers.join(', '));
    // Ne pas autoriser le swap si on ne peut pas valider DISSO
    return false;
  }

  const lv2_s1 = String(s1.row[idxLV2] || '').trim().toUpperCase();
  var opt_s1 = String(s1.row[idxOPT] || '').trim().toUpperCase();
  const lv2_s2 = String(s2.row[idxLV2] || '').trim().toUpperCase();
  var opt_s2 = String(s2.row[idxOPT] || '').trim().toUpperCase();

  // Neutraliser OPT si c'est en réalité une LV2 (anomalie saisie)
  if (isOPTAnomalyLV2(opt_s1)) opt_s1 = '';
  if (isOPTAnomalyLV2(opt_s2)) opt_s2 = '';

  // Bloquer swap si combinaison LV2+OPT interdite dans la classe cible
  if (!isLV2OPTCompatible(lv2_s2, opt_s2) || !isLV2OPTCompatible(lv2_s1, opt_s1)) {
    return false;
  }
  const disso_s1 = String(s1.row[idxDISSO] || '').trim().toUpperCase();
  const disso_s2 = String(s2.row[idxDISSO] || '').trim().toUpperCase();

  // Vérifier si s2 peut aller dans cls1
  const quotas1 = (ctx && ctx.quotas && ctx.quotas[cls1Name]) || {};
  const lv2Universelles = (ctx && ctx.lv2Universelles) || [];
  
  // ✅ BUG CRITIQUE CORRIGÉ : Vérifier LV2 ET OPT séparément (pas else if)
  // Un élève peut avoir LV2=ESP + OPT=CHAV en même temps !
  
  // Vérifier LV2 (LV2 universelles toujours compatibles)
  if (lv2_s2 && lv2Universelles.indexOf(lv2_s2) === -1 && isKnownLV2(lv2_s2)) {
    if (!quotas1[lv2_s2] || quotas1[lv2_s2] <= 0) {
      return false; // Classe cible ne propose pas cette LV2
    }
  }

  // Vérifier OPT (indépendamment de LV2)
  if (opt_s2 && isKnownOPT(opt_s2)) {
    if (!quotas1[opt_s2] || quotas1[opt_s2] <= 0) {
      return false; // Classe cible ne propose pas cette option
    }
  }
  
  // ✅ FIX #4 : Utiliser HARMONY_OPT_LIST au lieu de hardcode LATIN/CHAV
  // Vérifie si la classe propose des options spécialisées et si l'élève n'en a pas
  var classHasOptions = false;
  for (var oi = 0; oi < HARMONY_OPT_LIST.length; oi++) {
    if (quotas1[HARMONY_OPT_LIST[oi]] > 0) { classHasOptions = true; break; }
  }
  var studentHasNoOption = !opt_s2 || !isKnownOPT(opt_s2);

  if (classHasOptions && studentHasNoOption && lv2_s2 && lv2_s2 !== 'ESP') {
    return false;
  }
  
  // Vérifier si s1 peut aller dans cls2
  const quotas2 = (ctx && ctx.quotas && ctx.quotas[cls2Name]) || {};
  
  // Vérifier LV2 (LV2 universelles toujours compatibles)
  if (lv2_s1 && lv2Universelles.indexOf(lv2_s1) === -1 && isKnownLV2(lv2_s1)) {
    if (!quotas2[lv2_s1] || quotas2[lv2_s1] <= 0) {
      return false; // Classe cible ne propose pas cette LV2
    }
  }

  // Vérifier OPT (indépendamment de LV2)
  if (opt_s1 && isKnownOPT(opt_s1)) {
    if (!quotas2[opt_s1] || quotas2[opt_s1] <= 0) {
      return false; // Classe cible ne propose pas cette option
    }
  }
  
  // ✅ FIX #4 : Symétrique — utiliser HARMONY_OPT_LIST
  var class2HasOptions = false;
  for (var oi2 = 0; oi2 < HARMONY_OPT_LIST.length; oi2++) {
    if (quotas2[HARMONY_OPT_LIST[oi2]] > 0) { class2HasOptions = true; break; }
  }
  var student1HasNoOption = !opt_s1 || !isKnownOPT(opt_s1);

  if (class2HasOptions && student1HasNoOption && lv2_s1 && lv2_s1 !== 'ESP') {
    return false;
  }
  
  // Vérifier DISSO : s1 ne doit pas avoir le même code DISSO qu'un élève de cls2 (après swap)
  if (disso_s1) {
    for (let i = 0; i < idxList2.length; i++) {
      const idx = idxList2[i];
      if (idx === idx2) continue; // s2 sera swappé donc ne compte pas
      const otherStudent = allData[idx];
      const otherDisso = String(otherStudent.row[idxDISSO] || '').trim().toUpperCase();
      if (otherDisso && otherDisso === disso_s1) {
        return false; // Conflit DISSO
      }
    }
  }
  
  // Vérifier DISSO : s2 ne doit pas avoir le même code DISSO qu'un élève de cls1 (après swap)
  if (disso_s2) {
    for (let i = 0; i < idxList1.length; i++) {
      const idx = idxList1[i];
      if (idx === idx1) continue; // s1 sera swappé donc ne compte pas
      const otherStudent = allData[idx];
      const otherDisso = String(otherStudent.row[idxDISSO] || '').trim().toUpperCase();
      if (otherDisso && otherDisso === disso_s2) {
        return false; // Conflit DISSO
      }
    }
  }
  
  return true; // Swap autorisé
}

/**
 * Applique un swap d'indices entre deux classes avec logs détaillés
 */
function applySwap_Ultimate(allData, byClass, swap, headers) {
  const idx1 = swap.idx1;
  const idx2 = swap.idx2;

  // 📋 LOG détaillé des élèves swappés
  const s1 = allData[idx1];
  const s2 = allData[idx2];
  const idxNom = headers.indexOf('NOM');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');
  const idxDISSO = headers.indexOf('DISSO');

  const nom1 = idxNom >= 0 ? String(s1.row[idxNom] || '') : 'Élève 1';
  const nom2 = idxNom >= 0 ? String(s2.row[idxNom] || '') : 'Élève 2';

  const details1 = [];
  if (idxLV2 >= 0 && s1.row[idxLV2]) details1.push('LV2=' + s1.row[idxLV2]);
  if (idxOPT >= 0 && s1.row[idxOPT]) details1.push('OPT=' + s1.row[idxOPT]);
  if (idxDISSO >= 0 && s1.row[idxDISSO]) details1.push('DISSO=' + s1.row[idxDISSO]);

  const details2 = [];
  if (idxLV2 >= 0 && s2.row[idxLV2]) details2.push('LV2=' + s2.row[idxLV2]);
  if (idxOPT >= 0 && s2.row[idxOPT]) details2.push('OPT=' + s2.row[idxOPT]);
  if (idxDISSO >= 0 && s2.row[idxDISSO]) details2.push('DISSO=' + s2.row[idxDISSO]);

  logLine('DEBUG', `  🔄 ULTIMATE Swap: ${swap.cls1} ↔ ${swap.cls2}`);
  logLine('DEBUG', `    • ${nom1}: ${swap.cls1} → ${swap.cls2} (${details1.join(', ') || 'aucune contrainte'})`);
  logLine('DEBUG', `    • ${nom2}: ${swap.cls2} → ${swap.cls1} (${details2.join(', ') || 'aucune contrainte'})`);

  // Appliquer le swap
  byClass[swap.cls1] = byClass[swap.cls1].filter(i => i !== idx1).concat([idx2]);
  byClass[swap.cls2] = byClass[swap.cls2].filter(i => i !== idx2).concat([idx1]);
}

/**
 * Sauvegarde physiquement les résultats dans les onglets
 */
function saveResults_Ultimate(ss, allData, byClass, headersRef) {
  logLine('INFO', '💾 Début de l\'écriture physique des résultats...');

  if (!headersRef || headersRef.length === 0) {
    logLine('ERROR', '❌ Impossible de sauvegarder : En-têtes manquants');
    return { ok: false };
  }

  let successCount = 0;
  let errorCount = 0;

  for (const className in byClass) {
    const indices = byClass[className];
    
    // ✅ className est déjà le nom de destination (ex: "5°1")
    //    car Phase4 charge depuis les onglets TEST (5°1TEST)
    const testSheetName = className + 'TEST';
    const sheet = ss.getSheetByName(testSheetName);

    if (!sheet) {
      logLine('WARN', `⚠️ Onglet ${testSheetName} introuvable pour l'écriture.`);
      errorCount++;
      continue;
    }

    try {
      const rowsToWrite = [headersRef];
      indices.forEach(idx => {
        const student = allData[idx];
        rowsToWrite.push(student.row);
      });

      if (rowsToWrite.length > 0) {
        sheet.getRange(1, 1, rowsToWrite.length, headersRef.length).setValues(rowsToWrite);
        const lastRow = sheet.getLastRow();
        if (lastRow > rowsToWrite.length) {
          sheet.getRange(rowsToWrite.length + 1, 1, lastRow - rowsToWrite.length, sheet.getLastColumn()).clearContent();
        }
        logLine('INFO', `  ✅ ${testSheetName} : ${indices.length} élèves écrits.`);
        successCount++;
      }
    } catch (e) {
      logLine('ERROR', `  ❌ Erreur écriture ${testSheetName} : ${e.toString()}`);
      errorCount++;
    }
  }

  SpreadsheetApp.flush();
  logLine('SUCCESS', `💾 Sauvegarde complète : ${successCount} réussi(s), ${errorCount} erreur(s)`);

  return {
    ok: errorCount === 0,
    successCount: successCount,
    errorCount: errorCount
  };
}

// logLine() → supprimée (définition canonique dans App.Core.js)

/**
 * 🔍 VALIDATION FINALE : Vérifie qu'il n'y a pas de codes DISSO dupliqués dans les classes
 */
function validateDISSOConstraints_Ultimate(allData, byClass, headers) {
  const idxDISSO = headers.indexOf('DISSO');
  const idxNom = headers.indexOf('NOM');

  if (idxDISSO === -1) {
    logLine('WARN', '⚠️ Colonne DISSO non trouvée, validation DISSO ignorée');
    return { ok: true, message: 'Colonne DISSO non trouvée' };
  }

  // Vérifier chaque classe
  const duplicates = [];
  for (const cls in byClass) {
    const indices = byClass[cls];
    const dissoCounts = {};

    for (let i = 0; i < indices.length; i++) {
      const idx = indices[i];
      const student = allData[idx];
      const disso = String(student.row[idxDISSO] || '').trim().toUpperCase();
      if (!disso) continue;

      if (!dissoCounts[disso]) {
        dissoCounts[disso] = {
          code: disso,
          count: 0,
          noms: []
        };
      }

      dissoCounts[disso].count++;
      const nom = idxNom >= 0 ? String(student.row[idxNom] || '') : `Élève ${idx}`;
      dissoCounts[disso].noms.push(nom);
    }

    // Détecter duplications
    for (const code in dissoCounts) {
      if (dissoCounts[code].count > 1) {
        duplicates.push({
          classe: cls,
          code: code,
          count: dissoCounts[code].count,
          noms: dissoCounts[code].noms
        });
      }
    }
  }

  return {
    ok: duplicates.length === 0,
    duplicates: duplicates
  };
}

// ===================================================================
// TEST FUNCTIONS
// ===================================================================

/**
 * Lance le moteur Ultimate en mode test
 */
function testPhase4Ultimate() {
  const ctx = {
    ss: SpreadsheetApp.getActiveSpreadsheet()
  };
  const result = Phase4_Ultimate_Run(ctx);
  Logger.log('=== RÉSULTAT TEST ULTIMATE ===');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
