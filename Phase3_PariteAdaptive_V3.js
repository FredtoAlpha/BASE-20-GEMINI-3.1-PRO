/**
 * ===================================================================
 * PHASE 3 V3 - PARITÉ ADAPTATIVE
 * ===================================================================
 *
 * Nouveau principe : Remplir les classes en respectant au mieux la parité F/M
 * réelle du vivier restant, la tolérance configurée dans l'UI, et les contraintes
 * pédagogiques (LV2/OPT/ASSO/DISSO), sans biais systématique.
 *
 * Changements majeurs :
 * - Ratio basé sur poolF/poolM uniquement (pas _BASEOPTI global)
 * - Méthode des plus forts restes pour la distribution des quotas
 * - Tolérance de parité réellement appliquée
 * - Tie-break neutre avec alternance F/M
 * - Fallback contrôlé avec scoring et logging
 * - Contraintes LV2/OPT toujours actives via canPlaceInClass_V3
 */

// ===================================================================
// FONCTIONS UTILITAIRES
// ===================================================================

/**
 * Vérifie si une classe est dans la tolérance de parité
 */
function isWithinTolerance_V3(C, parityTolerance) {
  const finalF = C.currentF + (C.pendingF || 0);
  const finalM = C.currentM + (C.pendingM || 0);
  const diffF = finalF - C.targetF_total;
  const diffM = finalM - (C.targetTotal - C.targetF_total);
  return (Math.abs(diffF) <= parityTolerance && Math.abs(diffM) <= parityTolerance);
}

/**
 * Calcule le besoin global pour un sexe donné
 */
function globalNeed_V3(sex, targetFGlobal, targetMGlobal, placedF, placedM) {
  if (sex === 'F') return targetFGlobal - placedF;
  return targetMGlobal - placedM;
}

/**
 * Retourne le sexe opposé
 */
function oppositeSex_V3(sex) {
  return sex === 'F' ? 'M' : 'F';
}

/**
 * Calcule la pénalité de parité après placement d'un élève
 */
function parityPenaltyAfterPlacement_V3(C, sex, parityTolerance) {
  const finalF = C.currentF + (C.pendingF || 0) + (sex === 'F' ? 1 : 0);
  const finalM = C.currentM + (C.pendingM || 0) + (sex === 'M' ? 1 : 0);

  const targetF = C.targetF_total;
  const targetM = C.targetTotal - C.targetF_total;

  const diffF = Math.abs(finalF - targetF);
  const diffM = Math.abs(finalM - targetM);

  const overTolF = Math.max(0, diffF - parityTolerance);
  const overTolM = Math.max(0, diffM - parityTolerance);

  return overTolF + overTolM;
}

// ===================================================================
// CALCUL DES QUOTAS F/M
// ===================================================================

/**
 * Calcule les quotas F/M par classe selon la méthode des plus forts restes
 *
 * @param {Object} ctx - Contexte contenant classes, poolF, poolM
 * @param {Array} classes - Tableau des classes avec currentF, currentM, targetTotal
 * @param {Array} poolF - Pool des filles non assignées
 * @param {Array} poolM - Pool des garçons non assignés
 * @returns {Object} { targetFGlobal, targetMGlobal, ratioF, ratioM }
 */
function calculateParityTargets_V3(ctx, classes, poolF, poolM) {
  const totalF = poolF.length;
  const totalM = poolM.length;
  const totalPool = totalF + totalM;

  // Calcul du ratio réel du vivier
  const ratioF = totalPool > 0 ? totalF / totalPool : 0.5;
  const ratioM = 1 - ratioF;

  logLine('INFO', '📊 Ratio vivier restant : ' + (ratioF * 100).toFixed(1) + '% F / ' + (ratioM * 100).toFixed(1) + '% M');
  logLine('INFO', '   Pool : ' + totalF + ' F, ' + totalM + ' M (' + totalPool + ' élèves)');

  // Calculer le nombre total de sièges à pourvoir
  const totalSlots = classes.reduce(function(sum, C) {
    return sum + C.slotsLeft;
  }, 0);

  // Quota global sur les sièges restants
  const targetFGlobalExact = totalSlots * ratioF;
  const targetFGlobalFloor = Math.floor(targetFGlobalExact);
  const targetFGlobal = Math.min(totalF, targetFGlobalFloor + Math.round(targetFGlobalExact - targetFGlobalFloor));
  const targetMGlobal = totalSlots - targetFGlobal;

  logLine('INFO', '🎯 Quotas globaux pour ' + totalSlots + ' sièges : ' + targetFGlobal + ' F / ' + targetMGlobal + ' M');

  // ===== MÉTHODE DES PLUS FORTS RESTES =====

  // Étape 1 : Cible théorique F totale par classe
  const targets = [];
  let sumBaseFTotal = 0;

  classes.forEach(function(C) {
    const exactFTotal = C.targetTotal * ratioF;
    const baseF = Math.floor(exactFTotal);
    const remainder = exactFTotal - baseF;
    targets.push({ C: C, baseF: baseF, remainder: remainder });
    sumBaseFTotal += baseF;
  });

  // Étape 2 : Redistribuer les filles restantes au prorata des restes
  let remainingFGlobal = Math.max(0, targetFGlobal - sumBaseFTotal);

  targets.sort(function(a, b) {
    return b.remainder - a.remainder;
  });

  for (let i = 0; i < targets.length && remainingFGlobal > 0; i++) {
    targets[i].baseF++;
    remainingFGlobal--;
  }

  // Étape 3 : Tenir compte des filles déjà présentes
  targets.forEach(function(t) {
    const C = t.C;
    C.targetF_total = t.baseF;

    // Si on a déjà plus de filles que la cible, on ne visera plus que des garçons
    const remainingFForClass = Math.max(0, C.targetF_total - C.currentF);
    const localSlots = C.slotsLeft;
    C.targetF_newSlots = Math.min(remainingFForClass, localSlots);
    C.targetM_newSlots = localSlots - C.targetF_newSlots;

    logLine('INFO', '  📌 ' + C.name + ' : cible=' + C.targetF_total + 'F/' + (C.targetTotal - C.targetF_total) + 'M ' +
            '(actuel=' + C.currentF + 'F/' + C.currentM + 'M, à placer=' + C.targetF_newSlots + 'F/' + C.targetM_newSlots + 'M)');
  });

  return {
    targetFGlobal: targetFGlobal,
    targetMGlobal: targetMGlobal,
    ratioF: ratioF,
    ratioM: ratioM
  };
}

// ===================================================================
// DÉCISION DU SEXE À PLACER
// ===================================================================

/**
 * Décide quel sexe placer pour un siège donné en respectant la tolérance de parité
 *
 * @param {Object} C - Classe courante
 * @param {Object} ctx - Contexte avec données _BASEOPTI
 * @param {Object} meta - Métadonnées (parityTolerance, lastSexUsed, targetFGlobal, etc.)
 * @returns {string|null} 'F', 'M', ou null si pas de décision possible
 */
function decideSexForSeat_V3(C, ctx, meta) {
  const { parityTolerance, lastSexUsed, targetFGlobal, targetMGlobal, placedF, placedM, poolF, poolM } = meta;

  // Slots déjà ajoutés par sexe dans cette phase
  const finalF = C.currentF + (C.pendingF || 0);
  const finalM = C.currentM + (C.pendingM || 0);

  const targetF = C.targetF_total;
  const targetM = C.targetTotal - C.targetF_total;

  const diffF = targetF - finalF; // >0 => il manque des filles
  const diffM = targetM - finalM;

  const withinTol = (Math.abs(targetF - finalF) <= parityTolerance &&
                     Math.abs(targetM - finalM) <= parityTolerance);

  const poolFSize = poolF.length;
  const poolMSize = poolM.length;

  // Cas 1 : La classe est déjà dans la tolérance
  if (withinTol) {
    // On ne force plus localement, on regarde plutôt le global/pool
    const globalNeedF = globalNeed_V3('F', targetFGlobal, targetMGlobal, placedF, placedM);
    const globalNeedM = globalNeed_V3('M', targetFGlobal, targetMGlobal, placedF, placedM);

    let sex = null;

    if (poolFSize === 0 && poolMSize === 0) return null;
    if (poolFSize === 0) return 'M';
    if (poolMSize === 0) return 'F';

    if (globalNeedF > globalNeedM) sex = 'F';
    else if (globalNeedM > globalNeedF) sex = 'M';
    else {
      // Quasi équilibre global -> on alterne pour éviter les biais
      sex = (lastSexUsed === 'F') ? 'M' : 'F';
    }
    return sex;
  }

  // Cas 2 : La classe est encore clairement déséquilibrée
  // On regarde d'abord le besoin local
  let sex = null;

  if (diffF > diffM && diffF > 0) sex = 'F';
  else if (diffM > diffF && diffM > 0) sex = 'M';
  else if (diffF > 0 && diffM <= 0) sex = 'F';
  else if (diffM > 0 && diffF <= 0) sex = 'M';
  else {
    // Égalité parfaite ou cas bizarre : on se rabat sur le global
    const globalNeedF = globalNeed_V3('F', targetFGlobal, targetMGlobal, placedF, placedM);
    const globalNeedM = globalNeed_V3('M', targetFGlobal, targetMGlobal, placedF, placedM);

    if (globalNeedF > globalNeedM) sex = 'F';
    else if (globalNeedM > globalNeedF) sex = 'M';
    else sex = (lastSexUsed === 'F') ? 'M' : 'F';
  }

  // Respect des pools : si le sexe voulu est épuisé, on laisse la décision au fallback
  if (sex === 'F' && poolFSize === 0) return null;
  if (sex === 'M' && poolMSize === 0) return null;

  return sex;
}

// ===================================================================
// SÉLECTION DE L'ÉLÈVE
// ===================================================================

/**
 * Sélectionne l'élève du pool qui RAPPROCHE le plus la classe de la moyenne globale.
 * Version HARMONY : Placement intelligent par profil académique.
 *
 * @param {string} sex - 'F' ou 'M'
 * @param {Object} C - Classe courante
 * @param {Object} ctx - Contexte avec data, headers, poolF, poolM
 * @returns {Object|null} { eleve: rowData, index: poolIndex } ou null
 */
function pickStudentFromPool_V3(sex, C, ctx) {
  const pool = (sex === 'F') ? ctx.poolF : ctx.poolM;
  const data = ctx.data;
  const headers = ctx.headers;
  const idxCOM = headers.indexOf('COM');
  const idxTRA = headers.indexOf('TRA');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');

  // Calculer la moyenne COM/TRA actuelle de la classe
  let classSumCOM = 0, classSumTRA = 0, classCount = 0;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idxAssigned] || '').trim() === C.name) {
      classSumCOM += Number(data[r][idxCOM] || 2.5);
      classSumTRA += Number(data[r][idxTRA] || 2.5);
      classCount++;
    }
  }
  const classAvgCOM = classCount > 0 ? classSumCOM / classCount : 2.5;
  const classAvgTRA = classCount > 0 ? classSumTRA / classCount : 2.5;

  // Calculer la moyenne globale (cible)
  let globalSumCOM = 0, globalSumTRA = 0, globalCount = 0;
  for (let r = 1; r < data.length; r++) {
    globalSumCOM += Number(data[r][idxCOM] || 2.5);
    globalSumTRA += Number(data[r][idxTRA] || 2.5);
    globalCount++;
  }
  const globalAvgCOM = globalCount > 0 ? globalSumCOM / globalCount : 2.5;
  const globalAvgTRA = globalCount > 0 ? globalSumTRA / globalCount : 2.5;

  // Trouver le candidat compatible qui rapproche le plus la classe de la cible
  let bestCandidate = null;
  let bestGap = Infinity;

  for (let i = 0; i < pool.length; i++) {
    const eleveIdx = pool[i];
    const check = canPlaceInClass_V3(eleveIdx, C.name, data, headers, undefined, ctx);
    if (!check.ok) continue;

    const eleveCOM = Number(data[eleveIdx][idxCOM] || 2.5);
    const eleveTRA = Number(data[eleveIdx][idxTRA] || 2.5);

    // Calculer l'écart après ajout de cet élève
    const newAvgCOM = (classSumCOM + eleveCOM) / (classCount + 1);
    const newAvgTRA = (classSumTRA + eleveTRA) / (classCount + 1);
    const gap = Math.abs(newAvgCOM - globalAvgCOM) + Math.abs(newAvgTRA - globalAvgTRA);

    if (gap < bestGap) {
      bestGap = gap;
      bestCandidate = { eleve: data[eleveIdx], eleveIdx: eleveIdx, poolIndex: i };
    }
  }

  return bestCandidate;
}

// ===================================================================
// LOGGING DES DÉCISIONS
// ===================================================================

/**
 * Logue une décision de parité dans le système de logs
 *
 * @param {Object} C - Classe courante
 * @param {Object} details - Détails de la décision
 */
function logParityDecision_V3(C, details) {
  try {
    const timestamp = new Date().toISOString();
    const className = C.name || C.id || 'UNKNOWN';

    let message = '[PHASE3_PARITY] ' + className + ' | ' + details.type;

    if (details.type === 'PLACE') {
      message += ' | ' + details.sex + ' | ' + (details.eleveNom || 'N/A') + ' | ' + details.reason;
    } else if (details.type === 'FALLBACK_SEX') {
      message += ' | ' + details.fromSex + '→' + details.toSex + ' | ' + details.reason;
      message += ' | penalty: orig=' + (details.penaltyOriginal || 0).toFixed(2) + ', fallback=' + (details.penaltyFallback || 0).toFixed(2);
    } else if (details.type === 'SKIP_SLOT' || details.type === 'BLOCKED_SLOT') {
      message += ' | ' + details.reason;
      if (details.fromSex) message += ' | ' + details.fromSex + '→' + details.toSex;
    }

    logLine('INFO', '  ' + message);

    // Optionnel : Écrire dans un onglet de logs dédié
    // appendToLogSheet_(timestamp, 'PHASE3_PARITY', className, details);

  } catch (e) {
    // Ne jamais casser la phase 3 pour un simple log
    Logger.warn('Erreur logParityDecision_V3', { phase: 'Phase3' }, e);
  }
}

// ===================================================================
// BOUCLE PRINCIPALE - PHASE 3 PARITÉ ADAPTATIVE
// ===================================================================

/**
 * Phase 3 V3 : Complète les effectifs avec parité adaptative
 *
 * @param {Object} ctx - Contexte avec ss, levels, targets, quotas, parityTolerance
 * @returns {Object} { ok: boolean, placed: number, details: {...} }
 */
function Phase3I_completeAndParity_PariteAdaptive_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 3 V3 - PARITÉ ADAPTATIVE (Nouveau système)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');

  if (!baseSheet) {
    throw new Error('_BASEOPTI introuvable');
  }

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxSexe = headers.indexOf('SEXE');
  const idxNom = headers.indexOf('NOM');
  const idxPrenom = headers.indexOf('PRENOM');

  // Récupérer la tolérance de parité (depuis UI ou défaut)
  const parityTolerance = ctx.parityTolerance !== undefined ? ctx.parityTolerance : 1;
  logLine('INFO', '⚙️ Tolérance de parité : ±' + parityTolerance + ' élève(s)');

  // ===== INITIALISATION DES CLASSES =====

  const classes = [];
  (ctx.levels || []).forEach(function(clsName) {
    const target = (ctx.targets && ctx.targets[clsName]) || 0;
    let currentF = 0;
    let currentM = 0;

    // Compter les élèves déjà assignés
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxAssigned] || '').trim() === clsName) {
        const sexe = String(data[i][idxSexe] || '').toUpperCase();
        if (sexe === 'F') currentF++;
        else if (sexe === 'M') currentM++;
      }
    }

    const current = currentF + currentM;
    const slotsLeft = target - current;

    classes.push({
      name: clsName,
      index: classes.length,
      targetTotal: target,
      currentF: currentF,
      currentM: currentM,
      slotsLeft: slotsLeft,
      pendingF: 0,
      pendingM: 0
    });

    logLine('INFO', '  📋 ' + clsName + ' : ' + current + '/' + target + ' (besoin=' + slotsLeft + ', ' + currentF + 'F/' + currentM + 'M)');
  });

  // ===== CRÉATION DES POOLS F/M =====

  const poolF = [];
  const poolM = [];

  for (let i = 1; i < data.length; i++) {
    const assigned = String(data[i][idxAssigned] || '').trim();
    if (assigned) continue; // Déjà placé

    const sexe = String(data[i][idxSexe] || '').toUpperCase();
    if (sexe === 'F') {
      poolF.push(i);
    } else if (sexe === 'M') {
      poolM.push(i);
    }
  }

  logLine('INFO', '👥 Pool disponible : ' + poolF.length + ' F, ' + poolM.length + ' M');

  // ===== CALCUL DES QUOTAS F/M =====

  const quotaResult = calculateParityTargets_V3(ctx, classes, poolF, poolM);
  const targetFGlobal = quotaResult.targetFGlobal;
  const targetMGlobal = quotaResult.targetMGlobal;

  // ===== BOUCLE DE COMPLÉTION =====

  let lastSexUsed = 'M'; // Pour alterner au premier tie
  let placedF = 0;
  let placedM = 0;
  let totalPlaced = 0;

  // Créer le contexte étendu pour les fonctions
  ctx.data = data;
  ctx.headers = headers;
  ctx.poolF = poolF;
  ctx.poolM = poolM;

  let progress = true;
  let iterations = 0;
  const maxIterations = 1000; // Sécurité pour éviter boucle infinie

  while (progress && iterations < maxIterations) {
    progress = false;
    iterations++;

    for (let c = 0; c < classes.length; c++) {
      const C = classes[c];

      if (C.slotsLeft <= 0) continue;
      if (poolF.length + poolM.length === 0) break;

      const meta = {
        parityTolerance: parityTolerance,
        lastSexUsed: lastSexUsed,
        targetFGlobal: targetFGlobal,
        targetMGlobal: targetMGlobal,
        placedF: placedF,
        placedM: placedM,
        poolF: poolF,
        poolM: poolM
      };

      let sex = decideSexForSeat_V3(C, ctx, meta);
      if (!sex) continue; // Pas de décision possible, on laisse la classe pour un prochain tour

      let chosen = pickStudentFromPool_V3(sex, C, ctx);

      // ===== FALLBACK SI RIEN TROUVÉ =====

      if (!chosen) {
        const altSex = oppositeSex_V3(sex);
        const altCandidate = pickStudentFromPool_V3(altSex, C, ctx);

        if (altCandidate) {
          const penaltyAlt = parityPenaltyAfterPlacement_V3(C, altSex, parityTolerance);
          const penaltyOrig = parityPenaltyAfterPlacement_V3(C, sex, parityTolerance);

          if (penaltyAlt <= penaltyOrig) {
            // On accepte le fallback
            logParityDecision_V3(C, {
              type: 'FALLBACK_SEX',
              fromSex: sex,
              toSex: altSex,
              reason: 'no_compatible_candidate_primary_sex',
              penaltyOriginal: penaltyOrig,
              penaltyFallback: penaltyAlt
            });
            sex = altSex;
            chosen = altCandidate;
          } else {
            // On préfère bloquer et laisser une autre classe consommer ce sexe
            logParityDecision_V3(C, {
              type: 'SKIP_SLOT',
              fromSex: sex,
              toSex: altSex,
              reason: 'fallback_would_worsen_parity'
            });
            continue;
          }
        } else {
          // Aucun sexe compatible pour cette classe à ce tour
          logParityDecision_V3(C, {
            type: 'BLOCKED_SLOT',
            reason: 'no_candidate_any_sex'
          });
          continue;
        }
      }

      // ===== PLACEMENT VALIDÉ =====

      const eleveIdx = chosen.eleveIdx;
      const poolIndex = chosen.poolIndex;
      const eleve = chosen.eleve;

      // Marquer l'élève comme assigné
      data[eleveIdx][idxAssigned] = C.name;

      // Retirer du pool
      if (sex === 'F') {
        poolF.splice(poolIndex, 1);
        C.pendingF++;
        placedF++;
      } else {
        poolM.splice(poolIndex, 1);
        C.pendingM++;
        placedM++;
      }

      C.slotsLeft--;
      lastSexUsed = sex;
      totalPlaced++;
      progress = true;

      const eleveNom = String(eleve[idxNom] || '') + ' ' + String(eleve[idxPrenom] || '');
      logParityDecision_V3(C, {
        type: 'PLACE',
        sex: sex,
        eleveNom: eleveNom.trim(),
        reason: 'primary_parity_choice'
      });
    }
  }

  if (iterations >= maxIterations) {
    logLine('WARN', '⚠️ Limite d\'itérations atteinte (' + maxIterations + ')');
  }

  // ===== FINALISATION =====

  // Mettre à jour les compteurs définitifs
  classes.forEach(function(C) {
    C.currentF += C.pendingF;
    C.currentM += C.pendingM;
  });

  // Écrire dans _BASEOPTI
  baseSheet.getRange(1, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();

  // Sync vers colonnes legacy pour compatibilité audit
  if (typeof syncClassAssignedToLegacy_ === 'function') {
    syncClassAssignedToLegacy_('P3');
  }

  // ✅ CALCUL MOBILITÉ : Recalculer après Phase 3
  if (typeof computeMobilityFlags_ === 'function') {
    computeMobilityFlags_(ctx);
  }

  // ===== RAPPORT FINAL =====

  logLine('INFO', '');
  logLine('INFO', '✅ PHASE 3 V3 - PARITÉ ADAPTATIVE terminée');
  logLine('INFO', '  Placés : ' + totalPlaced + ' élèves (' + placedF + ' F / ' + placedM + ' M)');
  logLine('INFO', '  Pool restant : ' + poolF.length + ' F, ' + poolM.length + ' M');
  logLine('INFO', '');

  // Afficher état final par classe
  logLine('INFO', '📊 État final par classe :');
  classes.forEach(function(C) {
    const finalF = C.currentF;
    const finalM = C.currentM;
    const total = finalF + finalM;
    const ratioF = total > 0 ? (finalF / total * 100).toFixed(1) : 0;
    const diffF = Math.abs(finalF - C.targetF_total);
    const diffM = Math.abs(finalM - (C.targetTotal - C.targetF_total));
    const withinTol = (diffF <= parityTolerance && diffM <= parityTolerance);
    const status = withinTol ? '✅' : '⚠️';

    logLine('INFO', '  ' + status + ' ' + C.name + ' : ' + total + '/' + C.targetTotal + ' (' + finalF + 'F/' + finalM + 'M = ' + ratioF + '% F)');
  });

  // Vérifier élèves non placés
  let remaining = 0;
  for (let i = 1; i < data.length; i++) {
    if (!String(data[i][idxAssigned] || '').trim()) {
      remaining++;
    }
  }

  if (remaining > 0) {
    logLine('WARN', '⚠️ ' + remaining + ' élèves non placés après Phase 3 (contraintes bloquantes)');
  }

  return {
    ok: true,
    placed: totalPlaced,
    placedF: placedF,
    placedM: placedM,
    remaining: remaining,
    iterations: iterations
  };
}
