/**
 * ===================================================================
 * 🔄 PRIME LEGACY - PHASE 4 : ÉQUILIBRAGE SCORES (OPTIMUM PRIME)
 * ===================================================================
 *
 * Basé sur : OPTIMUM PRIME (Phase4_balanceScoresSwaps_BASEOPTI_V3)
 * Phase 4 : Optimise via swaps (Harmonie académique & Parité)
 * LIT : Onglets TEST
 * ÉCRIT : Onglets TEST (swaps _CLASS_ASSIGNED)
 *
 * Date : 2025-11-13
 * Branche : claude/prime-legacy-cleanup-015Zz6D3gh1QcbpR19TUYMLw
 *
 * ===================================================================
 */

/**
 * Phase 4 LEGACY : Équilibrage Scores via Swaps (OPTIMUM PRIME)
 * ✅ IMPLÉMENTATION FONCTIONNELLE basée sur OPTIMUM PRIME
 */
function Phase4_balanceScoresSwaps_LEGACY(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 4 LEGACY - Équilibrage Scores (OPTIMUM PRIME)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const maxSwaps = ctx.maxSwaps || 500;
  const weights = ctx.weights || { parity: 1.0, com: 1.0, tra: 0.5, part: 0.3, abs: 0.2 };

  // ========== CONSOLIDER DONNÉES ==========
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
    return { ok: false, swapsApplied: 0 };
  }

  const idxAssigned = headersRef.indexOf('_CLASS_ASSIGNED');
  const idxSexe = headersRef.indexOf('SEXE');
  const idxCOM = headersRef.indexOf('COM');
  const idxTRA = headersRef.indexOf('TRA');
  const idxPART = headersRef.indexOf('PART');
  const idxABS = headersRef.indexOf('ABS');
  const idxMobilite = headersRef.indexOf('MOBILITE');
  const idxFixe = headersRef.indexOf('FIXE');

  // ========== GROUPER PAR CLASSE ==========
  const byClass = {};
  for (let i = 0; i < allData.length; i++) {
    const cls = String(allData[i].row[idxAssigned] || '').trim();
    if (cls) {
      if (!byClass[cls]) byClass[cls] = [];
      byClass[cls].push(i);
    }
  }

  // ========== SWAPS OPTIMISATION ==========
  let swapsApplied = 0;

  for (let iter = 0; iter < maxSwaps; iter++) {
    const scoresBefore = calculateGlobalScore(allData, byClass, headersRef, weights);

    let bestSwap = null;
    let bestGain = 0;

    // Chercher le meilleur swap possible
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
          if (String(allData[i1].row[idxFixe] || '').toUpperCase() === 'FIXE') continue;
          if (String(allData[i2].row[idxFixe] || '').toUpperCase() === 'FIXE') continue;
          if (String(allData[i1].row[idxMobilite] || '').toUpperCase() === 'FIXE') continue;
          if (String(allData[i2].row[idxMobilite] || '').toUpperCase() === 'FIXE') continue;

          // Simuler swap
          allData[i1].row[idxAssigned] = cls2;
          allData[i2].row[idxAssigned] = cls1;

          // Recalculer byClass temporairement
          const tempByClass = JSON.parse(JSON.stringify(byClass));
          tempByClass[cls1][tempByClass[cls1].indexOf(i1)] = i2;
          tempByClass[cls2][tempByClass[cls2].indexOf(i2)] = i1;

          const scoresAfter = calculateGlobalScore(allData, tempByClass, headersRef, weights);
          const gain = scoresBefore - scoresAfter; // Réduction du score = amélioration

          // Annuler swap
          allData[i1].row[idxAssigned] = cls1;
          allData[i2].row[idxAssigned] = cls2;

          if (gain > bestGain) {
            bestGain = gain;
            bestSwap = { i1: i1, i2: i2, cls1: cls1, cls2: cls2 };
          }
        }
      }
    }

    // Appliquer le meilleur swap trouvé
    if (bestSwap && bestGain > 0.01) {
      const { i1, i2, cls1, cls2 } = bestSwap;

      allData[i1].row[idxAssigned] = cls2;
      allData[i2].row[idxAssigned] = cls1;

      // Mettre à jour byClass
      byClass[cls1][byClass[cls1].indexOf(i1)] = i2;
      byClass[cls2][byClass[cls2].indexOf(i2)] = i1;

      swapsApplied++;

      if (swapsApplied % 10 === 0) {
        logLine('INFO', '  🔄 ' + swapsApplied + ' swaps appliqués (gain cumulé: ' + bestGain.toFixed(2) + ')');
      }
    } else {
      logLine('INFO', '  🛑 Convergence atteinte (aucun swap bénéfique)');
      break;
    }
  }

  // ========== RÉÉCRIRE ==========
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

  logLine('INFO', '✅ PHASE 4 LEGACY terminée : ' + swapsApplied + ' swaps appliqués');

  return { ok: true, swapsApplied: swapsApplied };
}

/**
 * Calcule le score global (erreur à minimiser)
 */
function calculateGlobalScore(allData, byClass, headers, weights) {
  const idxSexe = headers.indexOf('SEXE');
  const idxCOM = headers.indexOf('COM');
  const idxTRA = headers.indexOf('TRA');
  const idxPART = headers.indexOf('PART');
  const idxABS = headers.indexOf('ABS');

  let totalScore = 0;

  // Score parité
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

  // Score académique (variance des moyennes)
  const criteria = ['COM', 'TRA', 'PART', 'ABS'];
  const indices = [idxCOM, idxTRA, idxPART, idxABS];
  const wts = [weights.com || 1.0, weights.tra || 0.5, weights.part || 0.3, weights.abs || 0.2];

  for (let c = 0; c < criteria.length; c++) {
    const idx = indices[c];
    const w = wts[c];

    const avgs = [];
    for (const cls in byClass) {
      let sum = 0, count = 0;
      byClass[cls].forEach(function(i) {
        sum += Number(allData[i].row[idx] || 3);
        count++;
      });
      if (count > 0) avgs.push(sum / count);
    }

    // Variance
    if (avgs.length > 0) {
      const mean = avgs.reduce((a, b) => a + b, 0) / avgs.length;
      const variance = avgs.reduce((sum, avg) => sum + Math.pow(avg - mean, 2), 0) / avgs.length;
      totalScore += variance * w;
    }
  }

  return totalScore;
}
