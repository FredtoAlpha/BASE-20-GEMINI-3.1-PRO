/**
 * ===================================================================
 * ARCHITECTURE V3 - _BASEOPTI COMME VIVIER UNIQUE
 * ===================================================================
 *
 * Principe :
 * - _BASEOPTI = SEULE source de vérité
 * - Colonne _CLASS_ASSIGNED pour marquer l'affectation
 * - CACHE vidé au début, rempli à la fin
 * - Toutes les phases lisent/écrivent dans _BASEOPTI
 */

// ===================================================================
// INIT - VIDER CACHE ET PRÉPARER _BASEOPTI
// ===================================================================

/**
 * Initialise l'optimisation :
 * 1. Vide les onglets CACHE
 * 2. Crée _BASEOPTI depuis les sources
 * 3. Ajoute colonne _CLASS_ASSIGNED (vide)
 */
function initOptimization_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '🔧 INIT V3 - Préparation _BASEOPTI');
  logLine('INFO', '='.repeat(80));

  // 1. VIDER les onglets CACHE
  logLine('INFO', '🧹 Vidage des onglets CACHE...');
  
  // 🔍 AUDIT CRITIQUE : Vérifier ctx.cacheSheets
  if (!ctx || !ctx.cacheSheets || ctx.cacheSheets.length === 0) {
    logLine('ERROR', '❌ PROBLÈME CRITIQUE: ctx.cacheSheets est vide ou undefined !');
    logLine('ERROR', '   ctx existe: ' + (ctx ? 'OUI' : 'NON'));
    if (ctx) {
      logLine('ERROR', '   ctx.cacheSheets: ' + (ctx.cacheSheets ? '[' + ctx.cacheSheets.join(', ') + ']' : 'UNDEFINED'));
      logLine('ERROR', '   Clés de ctx: ' + Object.keys(ctx).join(', '));
    }
  } else {
    logLine('INFO', '  📌 Onglets CACHE à vider: [' + ctx.cacheSheets.join(', ') + ']');
  }
  
  let cacheCleared = 0;
  let cacheCreated = 0;
  
  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const ss = ctx.ss || SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(cacheName);
    
    // Si l'onglet n'existe pas, le créer
    if (!sh) {
      logLine('INFO', '  📂 Création onglet: ' + cacheName);
      sh = ss.insertSheet(cacheName);
      cacheCreated++;
    }
    
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
      logLine('INFO', '  ✅ ' + cacheName + ' vidé (' + (sh.getLastRow() - 1) + ' lignes)');
      cacheCleared++;
    } else if (sh) {
      logLine('INFO', '  ℹ️ ' + cacheName + ' déjà vide');
    }
  });
  
  logLine('INFO', '  📊 Bilan: ' + cacheCleared + ' onglets vidés, ' + cacheCreated + ' onglets créés');

  // 2. Créer _BASEOPTI depuis sources
  logLine('INFO', '🎯 Création _BASEOPTI...');
  const result = createBaseOpti_(ctx);

  if (!result || !result.ok) {
    throw new Error('Échec création _BASEOPTI');
  }

  // 3. Ajouter colonne _CLASS_ASSIGNED
  logLine('INFO', '📋 Ajout colonne _CLASS_ASSIGNED...');
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');

  if (!baseSheet) {
    throw new Error('_BASEOPTI introuvable');
  }

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  // Vérifier si colonne existe déjà
  let colIdx = headers.indexOf('_CLASS_ASSIGNED');

  if (colIdx === -1) {
    // Ajouter la colonne
    headers.push('_CLASS_ASSIGNED');
    baseSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    colIdx = headers.length - 1;

    // Vider la colonne pour tous les élèves
    if (data.length > 1) {
      const emptyCol = Array(data.length - 1).fill(['']);
      baseSheet.getRange(2, colIdx + 1, data.length - 1, 1).setValues(emptyCol);
    }
  } else {
    // Vider la colonne existante
    if (data.length > 1) {
      baseSheet.getRange(2, colIdx + 1, data.length - 1, 1).clearContent();
    }
  }

  SpreadsheetApp.flush();

  logLine('INFO', '✅ INIT V3 terminé : ' + result.total + ' élèves dans _BASEOPTI');

  return {
    ok: true,
    total: result.total
  };
}

// ===================================================================
// PHASE 1 - PLACEMENT OPT/LV2
// ===================================================================

/**
 * Phase 1 V3 : Place les élèves avec OPT/LV2 spécifiques
 * LIT : _BASEOPTI (élèves sans _CLASS_ASSIGNED)
 * ÉCRIT : _BASEOPTI (colonne _CLASS_ASSIGNED)
 */
function Phase1_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 1 V3 - Options & LV2 (depuis _BASEOPTI)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');
  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  const idxID = headers.indexOf('_ID');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');

  if (idxAssigned === -1) {
    throw new Error('Colonne _CLASS_ASSIGNED manquante');
  }

  const stats = { ITA: 0, CHAV: 0, ESP: 0, ALL: 0, PT: 0 };

  // Parcourir les quotas
  for (const classe in (ctx.quotas || {})) {
    const quotas = ctx.quotas[classe];

    for (const optName in quotas) {
      const quota = quotas[optName];
      if (quota <= 0) continue;

      let placed = 0;

      // Parcourir _BASEOPTI et placer les élèves
      for (let i = 1; i < data.length; i++) {
        if (placed >= quota) break;

        const row = data[i];
        const assigned = String(row[idxAssigned] || '').trim();

        // Déjà placé ? Skip
        if (assigned) continue;

        const lv2 = String(row[idxLV2] || '').trim().toUpperCase();
        const opt = String(row[idxOPT] || '').trim().toUpperCase();

        let match = false;
        if (optName === 'ITA' || optName === 'ESP' || optName === 'ALL' || optName === 'PT') {
          match = (lv2 === optName);
        } else {
          match = (opt === optName);
        }

        if (match) {
          // Placer dans cette classe
          data[i][idxAssigned] = classe;
          placed++;
          stats[optName] = (stats[optName] || 0) + 1;
        }
      }

      logLine('INFO', '  ✅ ' + classe + ' : ' + placed + ' × ' + optName);
    }
  }

  // Écrire dans _BASEOPTI
  baseSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.flush();

  logLine('INFO', '✅ PHASE 1 V3 terminée : ' + Object.keys(stats).map(k => k + '=' + stats[k]).join(', '));

  return { ok: true, counts: stats };
}

// ===================================================================
// PHASE 2 - CODES ASSO/DISSO
// ===================================================================

/**
 * Phase 2 V3 : Applique codes A/D
 * LIT : _BASEOPTI (TOUS les élèves, placés ou non)
 * ÉCRIT : _BASEOPTI (UPDATE _CLASS_ASSIGNED)
 */
function Phase2_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 2 V3 - Codes ASSO/DISSO (depuis _BASEOPTI)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');
  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  const idxA = headers.indexOf('ASSO');
  const idxD = headers.indexOf('DISSO');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxNom = headers.indexOf('NOM');
  const idxPrenom = headers.indexOf('PRENOM');

  let assoMoved = 0;
  let dissoMoved = 0;

  // ============= CODES ASSO (A) =============

  const groupsA = {};
  for (let i = 1; i < data.length; i++) {
    const codeA = String(data[i][idxA] || '').trim().toUpperCase();
    if (codeA) {
      if (!groupsA[codeA]) groupsA[codeA] = [];
      groupsA[codeA].push(i);
    }
  }

  logLine('INFO', '🔗 Groupes ASSO : ' + Object.keys(groupsA).length);

  for (const code in groupsA) {
    const indices = groupsA[code];
    if (indices.length <= 1) continue;

    logLine('INFO', '  🔗 Groupe A=' + code + ' : ' + indices.length + ' élèves');

    // Trouver la classe majoritaire
    const classCounts = {};
    indices.forEach(function(idx) {
      const cls = String(data[idx][idxAssigned] || '').trim();
      if (cls) {
        classCounts[cls] = (classCounts[cls] || 0) + 1;
      }
    });

    let targetClass = null;
    let maxCount = 0;
    for (const cls in classCounts) {
      if (classCounts[cls] > maxCount) {
        maxCount = classCounts[cls];
        targetClass = cls;
      }
    }

    // Si aucun n'est placé, choisir classe équilibrée
    if (!targetClass) {
      targetClass = findLeastPopulatedClass_V3(data, headers, ctx);
    }

    logLine('INFO', '    🎯 Classe cible : ' + targetClass);

    // Déplacer tous vers la classe cible
    indices.forEach(function(idx) {
      const currentClass = String(data[idx][idxAssigned] || '').trim();
      if (currentClass !== targetClass) {
        data[idx][idxAssigned] = targetClass;
        assoMoved++;
        logLine('INFO', '      ✅ ' + data[idx][idxNom] + ' : ' + currentClass + ' → ' + targetClass);
      }
    });
  }

  // ============= CODES DISSO (D) =============

  const groupsD = {};
  for (let i = 1; i < data.length; i++) {
    const codeD = String(data[i][idxD] || '').trim().toUpperCase();
    if (codeD) {
      if (!groupsD[codeD]) groupsD[codeD] = [];
      groupsD[codeD].push(i);
    }
  }

  logLine('INFO', '🚫 Groupes DISSO : ' + Object.keys(groupsD).length);

  for (const code in groupsD) {
    const indices = groupsD[code];
    // ✅ CORRECTION : Ne pas skip les groupes avec 1 élève (il faut vérifier la classe)
    // if (indices.length <= 1) continue;

    logLine('INFO', '  🚫 Groupe D=' + code + ' : ' + indices.length + ' élève(s) à vérifier');

    // Vérifier si plusieurs sont dans la même classe
    const byClass = {};
    indices.forEach(function(idx) {
      const cls = String(data[idx][idxAssigned] || '').trim();
      if (cls) {
        if (!byClass[cls]) byClass[cls] = [];
        byClass[cls].push(idx);
      }
    });

    // Pour chaque classe avec > 1 élève D, déplacer les surplus
    for (const cls in byClass) {
      if (byClass[cls].length > 1) {
        logLine('INFO', '    ⚠️ ' + cls + ' contient ' + byClass[cls].length + ' élèves D=' + code);

        // Garder le premier, déplacer les autres
        for (let j = 1; j < byClass[cls].length; j++) {
          const idx = byClass[cls][j];

          // 🔒 Trouver classe sans ce code D (en vérifiant LV2/OPT)
          const targetClass = findClassWithoutCodeD_V3(data, headers, code, groupsD[code], idx, ctx);

          if (targetClass) {
            data[idx][idxAssigned] = targetClass;
            dissoMoved++;
            logLine('INFO', '      ✅ ' + data[idx][idxNom] + ' : ' + cls + ' → ' + targetClass);
          } else {
            logLine('WARN', '      ⚠️ ' + data[idx][idxNom] + ' reste en ' + cls + ' (contrainte LV2/OPT absolue)');
          }
        }
      }
    }
  }

  // Écrire dans _BASEOPTI
  baseSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.flush();

  logLine('INFO', '✅ PHASE 2 V3 terminée : ' + assoMoved + ' ASSO, ' + dissoMoved + ' DISSO');

  return { ok: true, asso: assoMoved, disso: dissoMoved };
}

// ===================================================================
// HELPERS
// ===================================================================
// NOTE: findLeastPopulatedClass_V3() et findClassWithoutCodeD_V3()
// sont définis dans Phases_BASEOPTI_V3_COMPLETE.js (version canonique).
// Les copies ici ont été supprimées pour éliminer les collisions GAS.

// ===================================================================
// PHASE 4 - SWAPS BASÉS SUR SCORES
// ===================================================================

/**
 * Phase 4 V3 : Optimise les scores par swaps
 * LIT : _BASEOPTI
 * PRIORITÉ : COM=1 > COM=2 > TRA > PART > ABS
 */
function Phase4_V3(ctx, weights) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 4 V3 - Swaps scores (depuis _BASEOPTI)');
  logLine('INFO', '='.repeat(80));

  // Récupérer poids depuis UI
  weights = weights || {
    com: 1.0,  // Priorité MAXIMALE
    tra: 0.7,
    part: 0.4,
    abs: 0.2
  };

  logLine('INFO', '⚖️ Poids : COM=' + weights.com + ', TRA=' + weights.tra + ', PART=' + weights.part + ', ABS=' + weights.abs);

  // ✅ RÉSOLU : Utiliser la Phase 4 complète implémentée dans Phases_BASEOPTI_V3_COMPLETE.gs
  // Cette fonction était un stub, la vraie implémentation existe déjà
  if (typeof Phase4_balanceScoresSwaps_BASEOPTI_V3 === 'function') {
    logLine('INFO', '🔀 Redirection vers Phase4_balanceScoresSwaps_BASEOPTI_V3...');
    return Phase4_balanceScoresSwaps_BASEOPTI_V3(ctx, weights, maxSwaps);
  } else {
    logLine('WARN', '⚠️ Phase4_balanceScoresSwaps_BASEOPTI_V3 non disponible');
    return { ok: true, swaps: 0 };
  }
}
