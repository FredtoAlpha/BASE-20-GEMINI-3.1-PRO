/**
 * ===================================================================
 * PHASES 1-2-3-4 V3 - _BASEOPTI COMME VIVIER UNIQUE
 * ⚓ JULES-VERNE-NAUTILUS - VERSION SUBLIMÉE
 * ===================================================================
 *
 * Cette version sublime l'algorithme "Optimus Prime" en introduisant
 * des stratégies avancées pour améliorer la performance et la stabilité.
 *
 * 🚀 AMÉLIORATIONS "NAUTILUS" :
 * 1. ⚡ Moteurs Silencieux (Swap Intelligent) : Au lieu de tester
 *    tous les swaps possibles (O(n²)), l'algorithme identifie les
 *    élèves "perturbateurs" (ceux qui dégradent le plus la qualité
 *    de leur classe) et concentre la recherche sur ce pool restreint.
 *    - Nouvelle fonction : `calculateStudentDisruption()`
 *    - `findBestSwap_V3` modifiée pour une recherche ciblée.
 *
 * 2. ⚓ Ancre d'Amarrage (Stabilité) : Pour éviter les oscillations
 *    (échanges répétitifs des mêmes élèves), une pénalité de mouvement
 *    est introduite. Chaque swap augmente légèrement le "coût" d'un
 *    futur swap impliquant les mêmes élèves, favorisant une
 *    convergence plus rapide et stable.
 *    - `Phase4_balanceScoresSwaps_BASEOPTI_V3` maintient un `swapHistory`.
 *    - `findBestSwap_V3` applique une pénalité basée sur cet historique.
 *
 * Basé sur : claude/optimum-prime-master-01SJDcJv7zHGGBXWhHpzfnxr
 * Date : 2025-11-13
 */

// ===================================================================
// PHASE 1 V3 - OPTIONS & LV2
// ===================================================================

/**
 * Phase 1 V3 : Place les élèves avec OPT/LV2 selon quotas
 * LIT : _BASEOPTI (colonne _CLASS_ASSIGNED vide)
 * ÉCRIT : _BASEOPTI (remplit _CLASS_ASSIGNED)
 */
function Phase1I_dispatchOptionsLV2_BASEOPTI_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 1 V3 - Options & LV2 (depuis _BASEOPTI)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');

  if (!baseSheet) {
    throw new Error('_BASEOPTI introuvable');
  }

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');

  if (idxAssigned === -1) {
    throw new Error('Colonne _CLASS_ASSIGNED manquante');
  }

  const stats = {};

  // Parcourir les quotas par classe
  for (const classe in (ctx.quotas || {})) {
    const quotas = ctx.quotas[classe];

    for (const optName in quotas) {
      const quota = quotas[optName];
      if (quota <= 0) continue;

      let placed = 0;

      // Parcourir _BASEOPTI
      for (let i = 1; i < data.length; i++) {
        if (placed >= quota) break;

        const row = data[i];
        const assigned = String(row[idxAssigned] || '').trim();

        if (assigned) continue; // Déjà placé

        const lv2 = String(row[idxLV2] || '').trim().toUpperCase();
        const opt = String(row[idxOPT] || '').trim().toUpperCase();

        let match = false;
        if (isKnownLV2(optName)) {
          match = (lv2 === optName);
        } else {
          match = (opt === optName);
        }

        if (match) {
          // ✅ PLACER SANS VÉRIFIER DISSO : LV2/OPT = RÈGLE ABSOLUE
          data[i][idxAssigned] = classe;
          placed++;
          stats[optName] = (stats[optName] || 0) + 1;
        }
      }

      if (placed > 0) {
        logLine('INFO', '  ✅ ' + classe + ' : ' + placed + ' × ' + optName + (placed < quota ? ' (⚠️ quota=' + quota + ')' : ''));
      }
    }
  }

  // Écrire dans _BASEOPTI
  baseSheet.getRange(1, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();

  // Sync vers colonnes legacy pour compatibilité audit
  syncClassAssignedToLegacy_('P1');

  // ⚡ OPTIMISATION QUOTA : Ne pas copier vers CACHE en Phase 1 (économiser les appels API)
  // La copie se fera en Phase 4 finale
  // copyBaseoptiToCache_V3(ctx);

  // ✅ CALCUL MOBILITÉ : Déterminer FIXE/PERMUT/LIBRE après Phase 1
  if (typeof computeMobilityFlags_ === 'function') {
    computeMobilityFlags_(ctx);
  } else {
    logLine('WARN', '⚠️ computeMobilityFlags_ non disponible (vérifier que Mobility_System.gs est chargé)');
  }

  logLine('INFO', '✅ PHASE 1 V3 terminée');

  return { ok: true, counts: stats };
}

// ===================================================================
// PHASE 2 V3 - CODES ASSO/DISSO
// ===================================================================

/**
 * Phase 2 V3 : Applique codes A (regrouper) et D (séparer)
 * LIT : _BASEOPTI (TOUS les élèves, placés ou non)
 * ÉCRIT : _BASEOPTI (update _CLASS_ASSIGNED)
 */
function Phase2I_applyDissoAsso_BASEOPTI_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 2 V3 - Codes ASSO/DISSO (depuis _BASEOPTI)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');

  if (!baseSheet) {
    throw new Error('_BASEOPTI introuvable');
  }

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  const idxA = headers.indexOf('ASSO');
  const idxD = headers.indexOf('DISSO');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxNom = headers.indexOf('NOM');

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
    if (indices.length <= 1) {
      logLine('INFO', '  ⏭️ A=' + code + ' : 1 seul élève');
      continue;
    }

    logLine('INFO', '  🔗 A=' + code + ' : ' + indices.length + ' élèves');

    // Trouver classe majoritaire
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

    // Si aucun placé, choisir classe la moins remplie
    if (!targetClass) {
      targetClass = findLeastPopulatedClass_V3(data, headers, ctx);
    }

    logLine('INFO', '    🎯 Cible : ' + targetClass);

    // Déplacer tous vers la cible
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
  // MULTI-RESTART: Tri par densité de contraintes (groupes les plus gros d'abord)
  // Les groupes DISSO avec le plus d'élèves sont les plus contraints,
  // ils doivent être placés en premier pour maximiser les options.

  const groupsD = {};
  for (let i = 1; i < data.length; i++) {
    const codeD = String(data[i][idxD] || '').trim().toUpperCase();
    if (codeD) {
      if (!groupsD[codeD]) groupsD[codeD] = [];
      groupsD[codeD].push(i);
      dissoMoved++;
    }
  }

  // Trier les codes DISSO par taille décroissante (plus contraints en premier)
  const sortedDissoCodes = Object.keys(groupsD).sort(function(a, b) {
    return groupsD[b].length - groupsD[a].length;
  });

  logLine('INFO', '🚫 Groupes DISSO : ' + sortedDissoCodes.length + ' (' + dissoMoved + ' élèves)');
  logLine('INFO', '  📐 Ordre de traitement (plus contraints d\'abord) : ' + sortedDissoCodes.map(function(c) { return c + '(' + groupsD[c].length + ')'; }).join(', '));

  for (let dIdx = 0; dIdx < sortedDissoCodes.length; dIdx++) {
    const code = sortedDissoCodes[dIdx];
    const indices = groupsD[code];

    logLine('INFO', '  🚫 D=' + code + ' : ' + indices.length + ' élève(s) à vérifier');

    // Vérifier si plusieurs sont dans la même classe
    const byClass = {};
    indices.forEach(function(idx) {
      const cls = String(data[idx][idxAssigned] || '').trim();
      if (cls) {
        if (!byClass[cls]) byClass[cls] = [];
        byClass[cls].push(idx);
      }
    });

    // Pour chaque classe avec >1 élève D, déplacer
    for (const cls in byClass) {
      if (byClass[cls].length > 1) {
        logLine('INFO', '    ⚠️ ' + cls + ' contient ' + byClass[cls].length + ' D=' + code);

        // Garder le premier, déplacer les autres
        for (let j = 1; j < byClass[cls].length; j++) {
          const idx = byClass[cls][j];

          // 🔒 Trouver classe sans ce code D (en vérifiant LV2/OPT)
          const targetClass = findClassWithoutCodeD_V3(data, headers, code, groupsD[code], idx, ctx);

          if (targetClass) {
            data[idx][idxAssigned] = targetClass;
            logLine('INFO', '      ✅ ' + data[idx][idxNom] + ' : ' + cls + ' → ' + targetClass + ' (séparation D=' + code + ')');
          } else {
            // ❌ Déplacement simple impossible → essayer un SWAP
            logLine('INFO', '      🔄 Tentative de SWAP pour ' + data[idx][idxNom] + ' (D=' + code + ')');

            const swapResult = findAndSwapForDisso_V3(data, headers, idx, code, cls, groupsD[code], ctx);

            if (swapResult.success) {
              logLine('INFO', '      ✅ SWAP réussi : ' + data[idx][idxNom] + ' ↔ ' + data[swapResult.swappedIdx][idxNom]);
              dissoMoved += 2; // Compter les 2 élèves swappés
            } else {
              logLine('WARN', '      ⚠️ ' + data[idx][idxNom] + ' reste en ' + cls + ' (' + swapResult.reason + ')');
            }
          }
        }
      }
    }
  }

  // Écrire dans _BASEOPTI
  baseSheet.getRange(1, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();

  // Sync vers colonnes legacy pour compatibilité audit
  syncClassAssignedToLegacy_('P2');

  // ⚡ OPTIMISATION QUOTA : Ne pas copier vers CACHE en Phase 2 (économiser les appels API)
  // La copie se fera en Phase 4 finale
  // copyBaseoptiToCache_V3(ctx);

  // ✅ CALCUL MOBILITÉ : Recalculer après Phase 2 (codes A/D peuvent changer les contraintes)
  if (typeof computeMobilityFlags_ === 'function') {
    computeMobilityFlags_(ctx);
  } else {
    logLine('WARN', '⚠️ computeMobilityFlags_ non disponible (vérifier que Mobility_System.gs est chargé)');
  }

  logLine('INFO', '✅ PHASE 2 V3 terminée : ' + assoMoved + ' ASSO, ' + dissoMoved + ' DISSO');

  return { ok: true, asso: assoMoved, disso: dissoMoved };
}

// ===================================================================
// HELPERS
// ===================================================================

function findLeastPopulatedClass_V3(data, headers, ctx) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const counts = {};

  (ctx.levels || []).forEach(function(cls) {
    counts[cls] = 0;
  });

  for (let i = 1; i < data.length; i++) {
    const cls = String(data[i][idxAssigned] || '').trim();
    if (cls && counts[cls] !== undefined) {
      counts[cls]++;
    }
  }

  let minClass = null;
  let minCount = Infinity;
  for (const cls in counts) {
    if (counts[cls] < minCount) {
      minCount = counts[cls];
      minClass = cls;
    }
  }

  return minClass || (ctx.levels && ctx.levels[0]) || '6°1';
}

/**
 * 🔄 SWAP DISSO : Trouve un élève compatible pour échanger et résoudre un conflit DISSO
 * Cherche un élève dans une autre classe qui :
 * - N'a pas le code DISSO problématique
 * - A la même LV2/OPT (pour respecter les quotas)
 * - Peut être placé dans la classe source sans créer de nouveau conflit
 *
 * @param {Array} data - Données _BASEOPTI
 * @param {Array} headers - En-têtes
 * @param {number} eleveIdx - Index de l'élève à déplacer
 * @param {string} codeD - Code DISSO à séparer
 * @param {string} sourceClass - Classe source (où il y a le doublon)
 * @param {Array} indicesWithD - Tous les indices des élèves avec ce code DISSO
 * @param {Object} ctx - Contexte avec quotas
 * @returns {Object} { success: boolean, swappedIdx: number|null, reason: string }
 */
function findAndSwapForDisso_V3(data, headers, eleveIdx, codeD, sourceClass, indicesWithD, ctx) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');
  const idxD = headers.indexOf('DISSO');
  const idxNom = headers.indexOf('NOM');

  // Récupérer LV2/OPT de l'élève à déplacer
  const eleveLV2 = String(data[eleveIdx][idxLV2] || '').trim().toUpperCase();
  const eleveOPT = String(data[eleveIdx][idxOPT] || '').trim().toUpperCase();
  const eleveCodeD = String(data[eleveIdx][idxD] || '').trim().toUpperCase();

  // Classes déjà occupées par ce code DISSO
  const classesWithD = new Set();
  indicesWithD.forEach(function(idx) {
    const cls = String(data[idx][idxAssigned] || '').trim();
    if (cls) classesWithD.add(cls);
  });

  // Parcourir toutes les classes (sauf la classe source)
  const allClasses = new Set();
  for (let i = 1; i < data.length; i++) {
    const cls = String(data[i][idxAssigned] || '').trim();
    if (cls && cls !== sourceClass) {
      allClasses.add(cls);
    }
  }

  // Chercher un élève compatible dans une autre classe
  for (const targetClass of Array.from(allClasses)) {
    // Skip les classes qui ont déjà ce code D
    if (classesWithD.has(targetClass)) continue;

    // Chercher un élève dans cette classe qui peut être swappé
    for (let i = 1; i < data.length; i++) {
      const candidateClass = String(data[i][idxAssigned] || '').trim();
      if (candidateClass !== targetClass) continue;

      const candidateLV2 = String(data[i][idxLV2] || '').trim().toUpperCase();
      const candidateOPT = String(data[i][idxOPT] || '').trim().toUpperCase();
      const candidateCodeD = String(data[i][idxD] || '').trim().toUpperCase();

      // Le candidat doit avoir la même LV2/OPT
      if (candidateLV2 !== eleveLV2 || candidateOPT !== eleveOPT) continue;

      // Le candidat ne doit pas avoir le code D problématique
      if (candidateCodeD === codeD) continue;

      // Vérifier si le candidat peut être placé dans la classe source
      // (pas de conflit DISSO avec son propre code D dans la classe source)
      if (candidateCodeD) {
        let conflict = false;
        for (let j = 1; j < data.length; j++) {
          if (j === i || j === eleveIdx) continue; // Skip le candidat lui-même et l'élève à déplacer
          const otherClass = String(data[j][idxAssigned] || '').trim();
          const otherCodeD = String(data[j][idxD] || '').trim().toUpperCase();
          if (otherClass === sourceClass && otherCodeD === candidateCodeD) {
            conflict = true;
            break;
          }
        }
        if (conflict) continue;
      }

      // ✅ Candidat trouvé ! Faire le swap
      data[eleveIdx][idxAssigned] = targetClass;
      data[i][idxAssigned] = sourceClass;

      logLine('INFO', '        🔄 ' + data[eleveIdx][idxNom] + ' (D=' + eleveCodeD + ') ' + sourceClass + ' → ' + targetClass);
      logLine('INFO', '        🔄 ' + data[i][idxNom] + ' (D=' + (candidateCodeD || 'aucun') + ') ' + targetClass + ' → ' + sourceClass);

      return { success: true, swappedIdx: i, reason: 'Swap réussi' };
    }
  }

  return { success: false, swappedIdx: null, reason: 'Aucun élève compatible pour swap' };
}

/**
 * 🔒 SÉCURITÉ DISSO : Trouve une classe sans le code DISSO spécifié
 * Vérifie aussi les contraintes LV2/OPT (règle absolue)
 *
 * @param {Array} data - Données _BASEOPTI
 * @param {Array} headers - En-têtes
 * @param {string} codeD - Code DISSO à éviter
 * @param {Array} indicesWithD - Indices des élèves avec ce code DISSO
 * @param {number} eleveIdx - Index de l'élève à déplacer
 * @param {Object} ctx - Contexte avec quotas
 * @returns {string|null} Classe cible ou null si impossible
 */
function findClassWithoutCodeD_V3(data, headers, codeD, indicesWithD, eleveIdx, ctx) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');

  // Récupérer LV2/OPT de l'élève
  const eleveLV2 = eleveIdx ? String(data[eleveIdx][idxLV2] || '').trim().toUpperCase() : '';
  const eleveOPT = eleveIdx ? String(data[eleveIdx][idxOPT] || '').trim().toUpperCase() : '';

  // ✅ BUG FIX: Scanner l'état ACTUEL de data pour trouver les classes avec ce code DISSO
  // (au lieu de se baser sur indicesWithD qui reflète l'état INITIAL avant les déplacements)
  const idxDISSO = headers.indexOf('DISSO');
  const classesWithD = new Set();
  for (let i = 1; i < data.length; i++) {
    if (i === eleveIdx) continue; // Ne pas compter l'élève lui-même
    const rowDisso = String(data[i][idxDISSO] || '').trim().toUpperCase();
    if (rowDisso === codeD) {
      const cls = String(data[i][idxAssigned] || '').trim();
      if (cls) classesWithD.add(cls);
    }
  }

  // Collecter toutes les classes
  const allClasses = new Set();
  for (let i = 1; i < data.length; i++) {
    const cls = String(data[i][idxAssigned] || '').trim();
    if (cls) allClasses.add(cls);
  }

  // MULTI-RESTART: Sélection intelligente — score chaque classe compatible,
  // choisir celle qui minimise la disruption de profil (pas first-found).
  const idxCOM = headers.indexOf('COM');
  const idxTRA = headers.indexOf('TRA');
  const eleveCOM = Number(data[eleveIdx][idxCOM] || 2.5);
  const eleveTRA = Number(data[eleveIdx][idxTRA] || 2.5);

  // Calculer la moyenne globale COM/TRA pour référence
  let globalCOM = 0, globalTRA = 0, globalN = 0;
  for (let i = 1; i < data.length; i++) {
    globalCOM += Number(data[i][idxCOM] || 2.5);
    globalTRA += Number(data[i][idxTRA] || 2.5);
    globalN++;
  }
  const avgGlobalCOM = globalN > 0 ? globalCOM / globalN : 2.5;
  const avgGlobalTRA = globalN > 0 ? globalTRA / globalN : 2.5;

  // Collecter les classes candidates avec leur score de disruption
  const candidates = [];

  if (eleveLV2 || eleveOPT) {
    for (const cls of Array.from(allClasses)) {
      if (classesWithD.has(cls)) continue;

      const quotas = (ctx && ctx.quotas && ctx.quotas[cls]) || {};
      let canPlace = false;
      if (eleveLV2 && isKnownLV2(eleveLV2)) {
        canPlace = (quotas[eleveLV2] !== undefined && quotas[eleveLV2] > 0);
      } else if (eleveOPT) {
        canPlace = (quotas[eleveOPT] !== undefined && quotas[eleveOPT] > 0);
      }

      if (canPlace) {
        // Calculer le profil moyen de la classe cible
        let clsCOM = 0, clsTRA = 0, clsN = 0;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idxAssigned] || '').trim() === cls) {
            clsCOM += Number(data[i][idxCOM] || 2.5);
            clsTRA += Number(data[i][idxTRA] || 2.5);
            clsN++;
          }
        }
        var clsAvgCOM = clsN > 0 ? clsCOM / clsN : 2.5;
        var clsAvgTRA = clsN > 0 ? clsTRA / clsN : 2.5;

        // Score = combien l'ajout de cet élève rapproche la classe de la moyenne globale
        var newAvgCOM = (clsCOM + eleveCOM) / (clsN + 1);
        var newAvgTRA = (clsTRA + eleveTRA) / (clsN + 1);
        var gapBefore = Math.abs(clsAvgCOM - avgGlobalCOM) + Math.abs(clsAvgTRA - avgGlobalTRA);
        var gapAfter = Math.abs(newAvgCOM - avgGlobalCOM) + Math.abs(newAvgTRA - avgGlobalTRA);
        var disruption = gapAfter - gapBefore; // Négatif = bon (rapproche de la moyenne)

        candidates.push({ cls: cls, disruption: disruption });
      }
    }

    if (candidates.length > 0) {
      candidates.sort(function(a, b) { return a.disruption - b.disruption; });
      logLine('INFO', '        ✅ Classe ' + candidates[0].cls + ' optimale (propose ' + (eleveLV2 || eleveOPT) + ', disruption=' + candidates[0].disruption.toFixed(3) + ')');
      return candidates[0].cls;
    }

    logLine('WARN', '        ⚠️ Aucune classe sans D=' + codeD + ' ne propose ' + (eleveLV2 || eleveOPT));
    logLine('WARN', '        🔒 CONTRAINTE LV2/OPT ABSOLUE : élève reste dans sa classe (doublon DISSO accepté)');
    return null;
  }

  // Si pas de LV2/OPT spécifique, choisir la classe sans DISSO qui profite le plus de cet élève
  for (const cls of Array.from(allClasses)) {
    if (!classesWithD.has(cls)) {
      candidates.push({ cls: cls, disruption: 0 });
    }
  }

  if (candidates.length > 0) {
    // Scorer chaque candidat pour choisir la meilleure
    for (let ci = 0; ci < candidates.length; ci++) {
      var _cls = candidates[ci].cls;
      var _clsCOM = 0, _clsTRA = 0, _clsN = 0;
      for (var ii = 1; ii < data.length; ii++) {
        if (String(data[ii][idxAssigned] || '').trim() === _cls) {
          _clsCOM += Number(data[ii][idxCOM] || 2.5);
          _clsTRA += Number(data[ii][idxTRA] || 2.5);
          _clsN++;
        }
      }
      var _clsAvg = _clsN > 0 ? _clsCOM / _clsN : 2.5;
      var _newAvg = (_clsCOM + eleveCOM) / (_clsN + 1);
      candidates[ci].disruption = Math.abs(_newAvg - avgGlobalCOM) - Math.abs(_clsAvg - avgGlobalCOM);
    }
    candidates.sort(function(a, b) { return a.disruption - b.disruption; });
    return candidates[0].cls;
  }

  return null;
}

/**
 * 🔒 GARDIEN DISSO/ASSO : Vérifie si un élève peut être placé dans une classe
 * sans violer les contraintes DISSO/ASSO
 *
 * @param {number} eleveIdx - Index de l'élève dans data
 * @param {string} targetClass - Classe de destination
 * @param {Array} data - Données _BASEOPTI
 * @param {Array} headers - En-têtes
 * @param {number} excludeIdx - Index de l'élève à exclure de la vérification (pour les swaps)
 * @param {Object} ctx - Contexte avec quotas (optionnel, pour vérifier LV2/OPT)
 * @returns {Object} { ok: boolean, reason: string }
 */
function canPlaceInClass_V3(eleveIdx, targetClass, data, headers, excludeIdx, ctx) {
  // 🛡️ Validations de sécurité
  if (!data || !headers || !targetClass) {
    logLine('ERROR', 'canPlaceInClass_V3 : paramètres invalides');
    return { ok: false, reason: 'Paramètres invalides' };
  }

  if (!data[eleveIdx]) {
    logLine('ERROR', 'canPlaceInClass_V3 : élève index ' + eleveIdx + ' introuvable');
    return { ok: false, reason: 'Élève introuvable' };
  }

  const idxD = headers.indexOf('DISSO');
  const idxA = headers.indexOf('ASSO');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');

  if (idxAssigned === -1) {
    logLine('ERROR', 'canPlaceInClass_V3 : colonne _CLASS_ASSIGNED introuvable');
    return { ok: false, reason: 'Colonne _CLASS_ASSIGNED manquante' };
  }

  const eleveD = String(data[eleveIdx][idxD] || '').trim().toUpperCase();
  const eleveA = String(data[eleveIdx][idxA] || '').trim().toUpperCase();
  const eleveLV2 = String(data[eleveIdx][idxLV2] || '').trim().toUpperCase();
  const eleveOPT = String(data[eleveIdx][idxOPT] || '').trim().toUpperCase();

  // 🔒 VÉRIFIER LV2/OPT : L'élève a une LV2 ou OPT ?
  if ((eleveLV2 || eleveOPT) && ctx && ctx.quotas) {
    const quotas = ctx.quotas[targetClass] || {};

    // Si l'élève a une LV2 spécifique (ITA, ESP, ALL, PT)
    if (eleveLV2 && isKnownLV2(eleveLV2)) {
      // La classe cible doit proposer cette LV2
      if (quotas[eleveLV2] === undefined || quotas[eleveLV2] === 0) {
        return {
          ok: false,
          reason: 'LV2 violation : élève a ' + eleveLV2 + ' mais classe ' + targetClass + ' ne la propose pas'
        };
      }
    }

    // Si l'élève a une option spécifique (CHAV, etc.)
    if (eleveOPT) {
      // La classe cible doit proposer cette option
      if (quotas[eleveOPT] === undefined || quotas[eleveOPT] === 0) {
        return {
          ok: false,
          reason: 'OPT violation : élève a ' + eleveOPT + ' mais classe ' + targetClass + ' ne la propose pas'
        };
      }
    }
  }

  // Vérifier DISSO : L'élève a un code D ?
  if (eleveD) {
    // Vérifier si la classe cible contient déjà un élève avec ce code D
    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue; // Skip lignes vides
      if (i === eleveIdx) continue; // Skip l'élève lui-même
      if (excludeIdx !== undefined && i === excludeIdx) continue; // Skip l'élève exclu (swap)

      const cls = String(data[i][idxAssigned] || '').trim();
      if (cls !== targetClass) continue; // Pas dans la classe cible

      const otherD = String(data[i][idxD] || '').trim().toUpperCase();
      if (otherD === eleveD) {
        return {
          ok: false,
          reason: 'DISSO violation : classe ' + targetClass + ' contient déjà un élève avec code D=' + eleveD
        };
      }
    }
  }

  // Vérifier ASSO : L'élève a un code A ?
  if (eleveA) {
    // Trouver où sont les autres membres du groupe ASSO
    let groupClass = null;
    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue; // Skip lignes vides
      if (i === eleveIdx) continue;

      const otherA = String(data[i][idxA] || '').trim().toUpperCase();
      if (otherA === eleveA) {
        const cls = String(data[i][idxAssigned] || '').trim();
        if (cls) {
          if (groupClass === null) {
            groupClass = cls;
          } else if (groupClass !== cls) {
            // Groupe ASSO déjà dispersé - ne pas ajouter de contrainte
            groupClass = null;
            break;
          }
        }
      }
    }

    // Si le groupe ASSO est déjà établi dans une classe, l'élève doit y aller
    if (groupClass && groupClass !== targetClass) {
      return {
        ok: false,
        reason: 'ASSO violation : groupe A=' + eleveA + ' est dans ' + groupClass + ', pas dans ' + targetClass
      };
    }
  }

  return { ok: true, reason: '' };
}

/**
 * 🔒 GARDIEN SWAP : Vérifie si un swap entre deux élèves viole DISSO/ASSO/LV2/OPT
 *
 * @param {number} idx1 - Index élève 1
 * @param {string} cls1 - Classe actuelle élève 1
 * @param {number} idx2 - Index élève 2
 * @param {string} cls2 - Classe actuelle élève 2
 * @param {Array} data - Données _BASEOPTI
 * @param {Array} headers - En-têtes
 * @param {Object} ctx - Contexte avec quotas
 * @returns {Object} { ok: boolean, reason: string }
 */
function canSwapStudents_V3(idx1, cls1, idx2, cls2, data, headers, ctx) {
  // Vérifier si élève 1 peut aller dans cls2 (en excluant idx2 qui va partir de cls2)
  const check1 = canPlaceInClass_V3(idx1, cls2, data, headers, idx2, ctx);
  if (!check1.ok) {
    return { ok: false, reason: 'Swap impossible : élève 1 → ' + cls2 + ' : ' + check1.reason };
  }

  // Vérifier si élève 2 peut aller dans cls1 (en excluant idx1 qui va partir de cls1)
  const check2 = canPlaceInClass_V3(idx2, cls1, data, headers, idx1, ctx);
  if (!check2.ok) {
    return { ok: false, reason: 'Swap impossible : élève 2 → ' + cls1 + ' : ' + check2.reason };
  }

  return { ok: true, reason: '' };
}

/**
 * TWO-PIPELINE PROP3 : Validation DISSO post-restart pour le pipeline V3.
 * Vérifie qu'aucune classe ne contient 2+ élèves avec le même code DISSO.
 * Opère sur le tableau data[] en mémoire (pas de lecture sheet).
 *
 * @param {Array} data - Tableau 2D [row][col] incluant headers en data[0]
 * @param {Array} headers - En-têtes (data[0])
 * @returns {{ ok: boolean, duplicates: Array }}
 */
function validateDISSOConstraints_V3_(data, headers) {
  var idxDISSO = headers.indexOf('DISSO');
  var idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  var idxNom = headers.indexOf('NOM');

  if (idxDISSO === -1 || idxAssigned === -1) {
    return { ok: true, duplicates: [], message: 'Colonnes DISSO ou _CLASS_ASSIGNED absentes' };
  }

  var byClassDisso = {};
  for (var i = 1; i < data.length; i++) {
    var cls = String(data[i][idxAssigned] || '').trim();
    if (!cls) continue;
    var disso = String(data[i][idxDISSO] || '').trim().toUpperCase();
    if (!disso) continue;

    if (!byClassDisso[cls]) byClassDisso[cls] = {};
    if (!byClassDisso[cls][disso]) byClassDisso[cls][disso] = { count: 0, noms: [] };
    byClassDisso[cls][disso].count++;
    var nom = idxNom >= 0 ? String(data[i][idxNom] || '') : 'Élève ' + i;
    byClassDisso[cls][disso].noms.push(nom);
  }

  var duplicates = [];
  for (var cls in byClassDisso) {
    for (var code in byClassDisso[cls]) {
      if (byClassDisso[cls][code].count > 1) {
        duplicates.push({
          classe: cls,
          code: code,
          count: byClassDisso[cls][code].count,
          noms: byClassDisso[cls][code].noms
        });
      }
    }
  }

  return { ok: duplicates.length === 0, duplicates: duplicates };
}

/**
 * Copie _BASEOPTI vers les onglets CACHE pour affichage live
 */
function copyBaseoptiToCache_V3(ctx) {
  logLine('INFO', '📋 copyBaseoptiToCache_V3: Début copie vers CACHE...');
  
  // 🔍 AUDIT CRITIQUE : Vérifier ctx.cacheSheets
  if (!ctx || !ctx.cacheSheets) {
    logLine('ERROR', '❌ PROBLÈME CRITIQUE: ctx.cacheSheets est undefined ou null !');
    logLine('ERROR', '   ctx existe: ' + (ctx ? 'OUI' : 'NON'));
    if (ctx) {
      logLine('ERROR', '   ctx.cacheSheets: ' + (ctx.cacheSheets || 'UNDEFINED'));
      logLine('ERROR', '   Clés de ctx: ' + Object.keys(ctx).join(', '));
    }
    return;
  }
  
  logLine('INFO', '  📌 ctx.cacheSheets: [' + ctx.cacheSheets.join(', ') + ']');
  
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');

  if (!baseSheet) {
    logLine('ERROR', '❌ _BASEOPTI introuvable !');
    return;
  }

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');

  logLine('INFO', '  📊 _BASEOPTI: ' + (data.length - 1) + ' élèves, colonne _CLASS_ASSIGNED: index=' + idxAssigned);

  // Grouper par classe
  const byClass = {};
  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const cls = cacheName.replace('CACHE', '').trim();
    byClass[cls] = [];
    logLine('INFO', '  📂 Initialisation classe: ' + cls + ' (onglet: ' + cacheName + ')');
  });

  // 🔍 AUDIT : Compter les élèves assignés
  let totalAssigned = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const cls = String(row[idxAssigned] || '').trim();

    if (cls) {
      totalAssigned++;
      if (byClass[cls]) {
        byClass[cls].push(row);
      } else {
        logLine('WARN', '  ⚠️ Élève assigné à classe inconnue: ' + cls + ' (ligne ' + (i+1) + ')');
      }
    }
  }

  logLine('INFO', '  📊 Élèves assignés: ' + totalAssigned + '/' + (data.length - 1));

  // 🔍 AUDIT : Afficher répartition par classe
  for (const cls in byClass) {
    logLine('INFO', '  📌 ' + cls + ': ' + byClass[cls].length + ' élèves');
  }

  // Écrire dans CACHE
  let sheetsWritten = 0;
  for (const cls in byClass) {
    const cacheName = cls + 'CACHE';
    const sh = ss.getSheetByName(cacheName);

    if (!sh) {
      logLine('WARN', '  ⚠️ Onglet CACHE introuvable: ' + cacheName);
      continue;
    }

    // Vider TOUT le contenu (y compris les en-têtes pour forcer la mise à jour)
    if (sh.getLastRow() > 0) {
      sh.clearContents();
    }

    // ✅ TOUJOURS écrire les en-têtes de _BASEOPTI (pour synchroniser les colonnes)
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Écrire élèves
    if (byClass[cls].length > 0) {
      sh.getRange(2, 1, byClass[cls].length, headers.length).setValues(byClass[cls]);
      logLine('INFO', '  ✅ ' + cacheName + ': ' + byClass[cls].length + ' élèves écrits');
      sheetsWritten++;
    } else {
      logLine('INFO', '  ℹ️ ' + cacheName + ': 0 élèves (vide)');
    }
  }

  SpreadsheetApp.flush();
  
  logLine('INFO', '✅ copyBaseoptiToCache_V3: ' + sheetsWritten + ' onglets CACHE remplis');
}

// ===================================================================
// PHASE 3 V3 - COMPLÉTER EFFECTIFS & PARITÉ (MODULAIRE)
// ===================================================================

/**
 * Phase 3 V3 : Wrapper pour appeler le module de parité adaptative.
 * LIT : _BASEOPTI (élèves non assignés)
 * ÉCRIT : _BASEOPTI (update _CLASS_ASSIGNED)
 */
function Phase3I_completeAndParity_BASEOPTI_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 3 V3 - Appel du module de parité adaptative');
  logLine('INFO', '='.repeat(80));

  // Vérifier si la fonction modulaire existe
  if (typeof Phase3I_completeAndParity_PariteAdaptive_V3 !== 'function') {
    throw new Error("La fonction modulaire 'Phase3I_completeAndParity_PariteAdaptive_V3' est introuvable. Assurez-vous que le fichier Phase3_PariteAdaptive_V3.gs est chargé.");
  }

  // Appeler directement la fonction modulaire avec le contexte
  return Phase3I_completeAndParity_PariteAdaptive_V3(ctx);
}

// NOTE: Première définition de logParityDecision supprimée (doublon dans le même fichier).
// La définition canonique est ci-dessous.

function logParityDecision(cls, details) {
  try {
    function formatPenalty(value) {
      if (value === undefined || value === null) return '';
      if (value === Infinity) return '∞';
      if (value === -Infinity) return '-∞';
      return value;
    }

    const row = [
      new Date(),
      'PHASE3_PARITY',
      cls && cls.name ? cls.name : (cls && cls.id ? cls.id : ''),
      details.type || '',
      details.sex || '',
      details.fromSex || '',
      details.toSex || '',
      details.reason || '',
      formatPenalty(details.penaltyOriginal),
      formatPenalty(details.penaltyFallback),
      details.eleveId || ''
    ];

    if (typeof appendLogRow === 'function') {
      appendLogRow(row);
    } else if (typeof logLine === 'function') {
      logLine('INFO', '📓 P3[' + row[2] + '] ' + JSON.stringify(details));
    }
  } catch (err) {
    if (typeof logLine === 'function') {
      logLine('WARN', '⚠️ Erreur logParityDecision : ' + err);
    }
  }
}



// ===================================================================
// CROSS-PHASE : RESHUFFLE DE LA PIRE CLASSE
// ===================================================================

/**
 * Identifie la classe avec la plus grande erreur et réinjecte une fraction
 * de ses élèves mobiles dans le pool non-assigné (_CLASS_ASSIGNED = '').
 * Phase 3 les replaceera ensuite de manière plus intelligente.
 *
 * Utilisé par la boucle cross-phase de l'orchestrateur V14I.
 */
function reshuffleWorstClass_V3_(ctx) {
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');
  if (!baseSheet) return;

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];
  const hIdx = {
    ASSIGNED: headers.indexOf('_CLASS_ASSIGNED'),
    SEXE: headers.indexOf('SEXE'),
    COM: headers.indexOf('COM'),
    TRA: headers.indexOf('TRA'),
    PART: headers.indexOf('PART'),
    ABS: headers.indexOf('ABS'),
    MOBILITE: headers.indexOf('MOBILITE'),
    FIXE: headers.indexOf('FIXE')
  };

  // Construire byClass
  var byClass = {};
  for (var i = 1; i < data.length; i++) {
    var cls = String(data[i][hIdx.ASSIGNED] || '').trim();
    if (cls) {
      if (!byClass[cls]) byClass[cls] = [];
      byClass[cls].push(i);
    }
  }

  // Calculer les stats globales
  var totalStudents = 0, globalNbF = 0, globalSumCOM = 0, globalSumTRA = 0;
  for (i = 1; i < data.length; i++) {
    totalStudents++;
    if (String(data[i][hIdx.SEXE] || '').toUpperCase().charAt(0) === 'F') globalNbF++;
    globalSumCOM += Number(data[i][hIdx.COM] || 2.5);
    globalSumTRA += Number(data[i][hIdx.TRA] || 2.5);
  }
  var globalStats = {
    ratioF: totalStudents > 0 ? globalNbF / totalStudents : 0.5,
    avgCOM: totalStudents > 0 ? globalSumCOM / totalStudents : 2.5,
    avgTRA: totalStudents > 0 ? globalSumTRA / totalStudents : 2.5,
    totalStudents: totalStudents
  };

  var weights = ctx.weights || { parity: 1.0, com: 1.0, tra: 0.5, part: 0.3, abs: 0.2, profiles: 2.0, effectif: 2.0 };

  // Trouver la pire classe (erreur la plus élevée)
  var worstClass = null;
  var worstError = -1;
  for (var clsName in byClass) {
    var targetSize = (ctx.targets && ctx.targets[clsName]) || 0;
    var cs = new ClassState(clsName, byClass[clsName], data, hIdx, targetSize);
    var err = cs.computeError(globalStats, null, weights);
    if (err > worstError) {
      worstError = err;
      worstClass = clsName;
    }
  }

  if (!worstClass) return;

  // Réinjecter une fraction des élèves mobiles de la pire classe
  var ratio = MULTI_RESTART_CONFIG.reshuffleWorstRatio;
  var indices = byClass[worstClass].slice();
  var mobileIndices = indices.filter(function(idx) {
    var mob = String(data[idx][hIdx.MOBILITE] || '').toUpperCase();
    var fixe = String(data[idx][hIdx.FIXE] || '').toUpperCase();
    return !mob.includes('FIXE') && !fixe.includes('FIXE') && !fixe.includes('OUI');
  });

  var nbToReshuffle = Math.max(1, Math.floor(mobileIndices.length * ratio));
  // Trier par profil extrême (les plus éloignés de la moyenne globale en premier)
  mobileIndices.sort(function(a, b) {
    var distA = Math.abs(Number(data[a][hIdx.COM] || 2.5) - globalStats.avgCOM) +
                Math.abs(Number(data[a][hIdx.TRA] || 2.5) - globalStats.avgTRA);
    var distB = Math.abs(Number(data[b][hIdx.COM] || 2.5) - globalStats.avgCOM) +
                Math.abs(Number(data[b][hIdx.TRA] || 2.5) - globalStats.avgTRA);
    return distB - distA; // Plus extrêmes en premier
  });

  var ejected = 0;
  for (var j = 0; j < nbToReshuffle && j < mobileIndices.length; j++) {
    data[mobileIndices[j]][hIdx.ASSIGNED] = ''; // Remettre dans le pool
    ejected++;
  }

  logLine('INFO', '  🔄 Reshuffle: ' + ejected + ' élèves éjectés de ' + worstClass + ' (erreur=' + worstError.toFixed(2) + ')');

  // Écrire dans _BASEOPTI
  baseSheet.getRange(1, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();
}

// ===================================================================
// PHASE 4 V3 - SWAPS BASÉS SUR L'HARMONIE ET LA PARITÉ
// ===================================================================

/**
 * Phase 4 V3 : Optimise la répartition par swaps en se basant sur un score
 * composite qui vise à harmoniser la distribution des scores dans les classes
 * tout en équilibrant la parité F/M.
 *
 * MULTI-RESTART VERSION (ex-REPLICANT) :
 * - N1: ClassState incrémental (O(1) par simulation au lieu de O(N))
 * - N2: Mini recuit simulé (accepte swaps négatifs avec budget décroissant)
 * - N3: PRNG seedable pour reproductibilité
 * - F5: 3-way cycle swaps avec ClassState
 * - MR: Multi-restart (5 seeds, garde le meilleur résultat)
 * - EFF: Pénalité d'effectif dans ClassState.computeError
 */
function Phase4_balanceScoresSwaps_BASEOPTI_V3(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 4 V3 - MULTI-RESTART NAUTILUS');
  logLine('INFO', '='.repeat(80));

  const weights = ctx.weights || { parity: 1.0, com: 1.0, tra: 0.5, part: 0.3, abs: 0.2, profiles: 2.0, effectif: 2.0 };
  const maxSwaps = ctx.maxSwaps || 500;
  const mrConfig = MULTI_RESTART_CONFIG;
  const maxRestarts = mrConfig.maxRestarts;

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');
  const snapshot = baseSheet.getDataRange().getValues();
  const headers = snapshot[0];

  // Index des colonnes (invariant entre restarts)
  const hIdx = {
    ASSIGNED: headers.indexOf('_CLASS_ASSIGNED'),
    SEXE: headers.indexOf('SEXE'),
    COM: headers.indexOf('COM'),
    TRA: headers.indexOf('TRA'),
    PART: headers.indexOf('PART'),
    ABS: headers.indexOf('ABS'),
    MOBILITE: headers.indexOf('MOBILITE'),
    FIXE: headers.indexOf('FIXE')
  };

  // Pré-calcul invariant : distribution cible + stats globales
  const snapshotByClass = {};
  for (let i = 1; i < snapshot.length; i++) {
    const cls = String(snapshot[i][hIdx.ASSIGNED] || '').trim();
    if (cls) {
      if (!snapshotByClass[cls]) snapshotByClass[cls] = [];
      snapshotByClass[cls].push(i);
    }
  }
  const targetDistribution = calculateTargetDistribution_V3(snapshot, headers, snapshotByClass);

  let totalStudents = 0, globalSumCOM = 0, globalSumTRA = 0, globalNbF = 0;
  for (let i = 1; i < snapshot.length; i++) {
    totalStudents++;
    globalSumCOM += Number(snapshot[i][hIdx.COM] || 2.5);
    globalSumTRA += Number(snapshot[i][hIdx.TRA] || 2.5);
    if (String(snapshot[i][hIdx.SEXE] || '').toUpperCase().charAt(0) === 'F') globalNbF++;
  }
  const globalStats = {
    ratioF: totalStudents > 0 ? globalNbF / totalStudents : 0.5,
    avgCOM: totalStudents > 0 ? globalSumCOM / totalStudents : 2.5,
    avgTRA: totalStudents > 0 ? globalSumTRA / totalStudents : 2.5,
    totalStudents: totalStudents
  };

  logLine('INFO', `🎯 ${totalStudents} élèves, ${Object.keys(snapshotByClass).length} classes, ratio F=${(globalStats.ratioF*100).toFixed(1)}%`);
  logLine('INFO', `🔁 Multi-restart : ${maxRestarts} seeds`);

  // ===== MULTI-RESTART LOOP =====
  let bestData = null;
  let bestError = Infinity;
  let bestSwaps = 0;
  let bestSwaps3Way = 0;
  let bestAnnealing = 0;
  let bestSeed = 0;

  for (let restart = 0; restart < maxRestarts; restart++) {
    const seed = ctx.seed ? ctx.seed + restart * mrConfig.seedSpacing : restart * mrConfig.seedSpacing;
    const rng = createRNG(seed);

    // Copie indépendante des données pour ce restart
    const data = snapshotData_(snapshot);
    const byClass = snapshotByClass_(snapshotByClass);

    logLine('INFO', `  🔁 Restart ${restart + 1}/${maxRestarts} (seed=${seed})`);

    // Résultat de ce restart
    const result = runPhase4CoreLoop_V3_(data, headers, hIdx, byClass, weights, maxSwaps, rng, globalStats, targetDistribution, ctx);

    logLine('INFO', `    📊 Erreur=${result.finalError.toFixed(2)}, swaps=${result.swapsApplied}+${result.swaps3Way}(3-way)`);

    // TWO-PIPELINE PROP3 : Valider DISSO PAR RESTART — rejeter tout restart invalide
    const restartValidation = validateDISSOConstraints_V3_(data, headers);
    if (!restartValidation.ok) {
      logLine('WARN', `    ❌ Restart ${restart + 1} rejeté : DISSO invalide (${restartValidation.duplicates.length} conflit(s))`);
      restartValidation.duplicates.forEach(function(dup) {
        logLine('WARN', `      • ${dup.classe} : ${dup.code} x${dup.count} (${dup.noms.join(', ')})`);
      });
      continue; // On ne considère JAMAIS ce restart, même si son erreur est meilleure
    }

    if (result.finalError < bestError) {
      bestError = result.finalError;
      bestData = data;
      bestSwaps = result.swapsApplied;
      bestSwaps3Way = result.swaps3Way;
      bestAnnealing = result.annealingUsed;
      bestSeed = seed;
      logLine('INFO', `    ⭐ Nouveau meilleur ! (erreur=${bestError.toFixed(2)}, DISSO ✅)`);
    }
  }

  // TWO-PIPELINE : Si aucun restart n'a passé la validation DISSO
  if (!bestData) {
    logLine('ERROR', '❌ AUCUN restart valide (tous rejetés par DISSO). Fallback sur l\'état pré-Phase4.');
    return {
      ok: false,
      swapsApplied: 0,
      swaps3Way: 0,
      annealingUsed: 0,
      seed: 0,
      restarts: maxRestarts,
      finalError: Infinity
    };
  }

  // Écrire le meilleur résultat
  logLine('INFO', `📊 Meilleur restart : seed=${bestSeed}, erreur=${bestError.toFixed(2)}, swaps=${bestSwaps}+${bestSwaps3Way}(3-way)`);
  baseSheet.getRange(1, 1, bestData.length, headers.length).setValues(bestData);
  SpreadsheetApp.flush();
  copyBaseoptiToCache_V3(ctx);
  if (typeof computeMobilityFlags_ === 'function') computeMobilityFlags_(ctx);

  logLine('INFO', `✅ PHASE 4 MULTI-RESTART terminée : meilleur sur ${maxRestarts} seeds. Seed gagnant: ${bestSeed}`);

  return {
    ok: true,
    swapsApplied: bestSwaps + bestSwaps3Way,
    swaps3Way: bestSwaps3Way,
    annealingUsed: bestAnnealing,
    seed: bestSeed,
    restarts: maxRestarts,
    finalError: bestError
  };
}

/**
 * Coeur du moteur de swaps Phase 4 V3.
 * Travaille uniquement en mémoire (data[]) — pas d'I/O sheet.
 * Appelé N fois par le multi-restart avec des seeds différentes.
 *
 * @returns {{ swapsApplied, swaps3Way, annealingUsed, finalError }}
 */
function runPhase4CoreLoop_V3_(data, headers, hIdx, byClass, weights, maxSwaps, rng, globalStats, targetDistribution, ctx) {
  // N1: Initialiser les ClassState incrémentaux
  const classStates = {};
  for (const cls in byClass) {
    const targetSize = (ctx.targets && ctx.targets[cls]) || 0;
    classStates[cls] = new ClassState(cls, byClass[cls], data, hIdx, targetSize);
  }

  function totalError() {
    return computeTotalError_(classStates, globalStats, targetDistribution, weights);
  }

  let currentError = totalError();

  function isStudentFixed(idx) {
    const mob = String(data[idx][hIdx.MOBILITE] || '').toUpperCase();
    const fixe = String(data[idx][hIdx.FIXE] || '').toUpperCase();
    return mob.includes('FIXE') || fixe.includes('FIXE') || fixe.includes('OUI');
  }

  const swapHistory = new Map();
  let swapsApplied = 0;

  // TWO-PIPELINE : Recuit simulé amélioré (refroidissement géométrique)
  let saTemperature = 50.0;
  const saCoolingRate = 0.995;
  const saMinTemp = 0.1;
  const saMaxDegradation = 0.02 * currentError; // 2% de l'erreur initiale max
  let annealingUsed = 0;

  const classNames = Object.keys(byClass);
  let stagnation = 0;
  const stagnationLimit = 50;

  for (let iter = 0; iter < maxSwaps; iter++) {
    const candidates = [];
    for (const cls in byClass) {
      const classError = classStates[cls].computeError(globalStats, targetDistribution, weights);
      for (let k = 0; k < byClass[cls].length; k++) {
        const si = byClass[cls][k];
        if (isStudentFixed(si)) continue;
        const errorWithout = classStates[cls].simulateSwap(si, si, data, hIdx, globalStats, targetDistribution, weights);
        const disruption = Math.abs(classError - errorWithout);
        candidates.push({ idx: si, cls: cls, disruption: disruption });
      }
    }
    candidates.sort((a, b) => b.disruption - a.disruption);

    const poolSize = Math.max(10, Math.floor(candidates.length * 0.6));
    const pool = candidates.slice(0, poolSize);

    let bestSwap = null;
    let bestGain = -Infinity;

    const maxPairs = Math.min(300, pool.length * (pool.length - 1) / 2);
    for (let p = 0; p < maxPairs; p++) {
      const s1 = rng.pick(pool);
      const s2 = rng.pick(pool);
      if (!s1 || !s2 || s1.cls === s2.cls || s1.idx === s2.idx) continue;

      const idx1 = s1.idx, idx2 = s2.idx;
      const cls1 = s1.cls, cls2 = s2.cls;

      const swapCheck = canSwapStudents_V3(idx1, cls1, idx2, cls2, data, headers, ctx);
      if (!swapCheck.ok) continue;

      const errBefore = classStates[cls1].computeError(globalStats, targetDistribution, weights) +
                        classStates[cls2].computeError(globalStats, targetDistribution, weights);
      const errCls1After = classStates[cls1].simulateSwap(idx1, idx2, data, hIdx, globalStats, targetDistribution, weights);
      const errCls2After = classStates[cls2].simulateSwap(idx2, idx1, data, hIdx, globalStats, targetDistribution, weights);

      let gain = errBefore - (errCls1After + errCls2After);
      const h1 = swapHistory.get(idx1) || 0;
      const h2 = swapHistory.get(idx2) || 0;
      gain = gain / (1 + h1 + h2);

      if (gain > bestGain) {
        bestGain = gain;
        bestSwap = { idx1, idx2, cls1, cls2, gain };
      }
    }

    let accepted = false;
    if (bestSwap && bestSwap.gain > 1e-6) {
      accepted = true;
    } else if (bestSwap && saTemperature > saMinTemp && bestSwap.gain > -saMaxDegradation && bestSwap.gain < 0) {
      // TWO-PIPELINE : SA géométrique (probabilité e^(gain/T), T décroissante)
      const acceptProba = Math.exp(bestSwap.gain / saTemperature);
      if (rng.next() < acceptProba) {
        accepted = true;
        annealingUsed++;
      }
    }
    // Refroidissement géométrique
    saTemperature *= saCoolingRate;

    if (!accepted) {
      stagnation++;
      if (stagnation >= stagnationLimit) break;
      continue;
    }

    stagnation = 0;
    const { idx1, idx2, cls1, cls2 } = bestSwap;

    classStates[cls1].applySwap(idx1, idx2, data, hIdx);
    classStates[cls2].applySwap(idx2, idx1, data, hIdx);

    data[idx1][hIdx.ASSIGNED] = cls2;
    data[idx2][hIdx.ASSIGNED] = cls1;
    const pos1 = byClass[cls1].indexOf(idx1);
    const pos2 = byClass[cls2].indexOf(idx2);
    if (pos1 !== -1) byClass[cls1].splice(pos1, 1, idx2);
    if (pos2 !== -1) byClass[cls2].splice(pos2, 1, idx1);

    swapHistory.set(idx1, (swapHistory.get(idx1) || 0) + 1);
    swapHistory.set(idx2, (swapHistory.get(idx2) || 0) + 1);
    swapsApplied++;
    currentError = totalError();
  }

  // TWO-PIPELINE : Post-SA greedy convergence (reconverger dans le nouveau bassin)
  if (annealingUsed > 0) {
    logLine('INFO', `    🌡️ SA: ${annealingUsed} swaps dégradants acceptés, T finale=${saTemperature.toFixed(4)}`);
    let postSASwaps = 0;
    let postStagnation = 0;
    for (let iter2 = 0; iter2 < Math.floor(maxSwaps * 0.3); iter2++) {
      const candidates2 = [];
      for (const cls2 in byClass) {
        const classError2 = classStates[cls2].computeError(globalStats, targetDistribution, weights);
        for (let k2 = 0; k2 < byClass[cls2].length; k2++) {
          const si2 = byClass[cls2][k2];
          if (isStudentFixed(si2)) continue;
          candidates2.push({ idx: si2, cls: cls2, disruption: classError2 });
        }
      }
      const poolSize2 = Math.max(10, Math.floor(candidates2.length * 0.6));
      candidates2.sort(function(a, b) { return b.disruption - a.disruption; });
      const pool2 = candidates2.slice(0, poolSize2);
      let bestSwap2 = null;
      let bestGain2 = 1e-6;
      for (let p2 = 0; p2 < Math.min(300, pool2.length * (pool2.length - 1) / 2); p2++) {
        const s1 = rng.pick(pool2);
        const s2 = rng.pick(pool2);
        if (!s1 || !s2 || s1.cls === s2.cls || s1.idx === s2.idx) continue;
        const chk = canSwapStudents_V3(s1.idx, s1.cls, s2.idx, s2.cls, data, headers, ctx);
        if (!chk.ok) continue;
        const errBefore = classStates[s1.cls].computeError(globalStats, targetDistribution, weights) +
                          classStates[s2.cls].computeError(globalStats, targetDistribution, weights);
        const errA = classStates[s1.cls].simulateSwap(s1.idx, s2.idx, data, hIdx, globalStats, targetDistribution, weights);
        const errB = classStates[s2.cls].simulateSwap(s2.idx, s1.idx, data, hIdx, globalStats, targetDistribution, weights);
        const g = errBefore - (errA + errB);
        if (g > bestGain2) { bestGain2 = g; bestSwap2 = { idx1: s1.idx, idx2: s2.idx, cls1: s1.cls, cls2: s2.cls }; }
      }
      if (!bestSwap2) { postStagnation++; if (postStagnation >= stagnationLimit) break; continue; }
      postStagnation = 0;
      classStates[bestSwap2.cls1].applySwap(bestSwap2.idx1, bestSwap2.idx2, data, hIdx);
      classStates[bestSwap2.cls2].applySwap(bestSwap2.idx2, bestSwap2.idx1, data, hIdx);
      data[bestSwap2.idx1][hIdx.ASSIGNED] = bestSwap2.cls2;
      data[bestSwap2.idx2][hIdx.ASSIGNED] = bestSwap2.cls1;
      const p1 = byClass[bestSwap2.cls1].indexOf(bestSwap2.idx1);
      const p2b = byClass[bestSwap2.cls2].indexOf(bestSwap2.idx2);
      if (p1 !== -1) byClass[bestSwap2.cls1].splice(p1, 1, bestSwap2.idx2);
      if (p2b !== -1) byClass[bestSwap2.cls2].splice(p2b, 1, bestSwap2.idx1);
      postSASwaps++;
      swapsApplied++;
    }
    if (postSASwaps > 0) {
      logLine('INFO', `    🎯 Post-SA greedy: ${postSASwaps} swaps supplémentaires`);
    }
  }

  // ===== 3-WAY CYCLE SWAPS =====
  const max3Way = Math.min(100, Math.floor(maxSwaps * 0.2));
  let swaps3Way = 0;

  for (let iter3 = 0; iter3 < max3Way; iter3++) {
    let bestGain3 = 1e-6;
    let best3Way = null;

    for (let t = 0; t < Math.min(20, classNames.length * (classNames.length - 1)); t++) {
      const c1 = rng.pick(classNames);
      const c2 = rng.pick(classNames);
      const c3 = rng.pick(classNames);
      if (c1 === c2 || c2 === c3 || c1 === c3) continue;

      const errBefore3 = classStates[c1].computeError(globalStats, targetDistribution, weights) +
                         classStates[c2].computeError(globalStats, targetDistribution, weights) +
                         classStates[c3].computeError(globalStats, targetDistribution, weights);

      for (let s = 0; s < 8; s++) {
        const a = rng.pick(byClass[c1]);
        const b = rng.pick(byClass[c2]);
        const c = rng.pick(byClass[c3]);
        if (!a || !b || !c) continue;
        if (isStudentFixed(a) || isStudentFixed(b) || isStudentFixed(c)) continue;

        const chkA = canPlaceInClass_V3(a, c2, data, headers, b, ctx);
        const chkB = canPlaceInClass_V3(b, c3, data, headers, c, ctx);
        const chkC = canPlaceInClass_V3(c, c1, data, headers, a, ctx);
        if (!chkA.ok || !chkB.ok || !chkC.ok) continue;

        const errC1After = classStates[c1].simulateSwap(a, c, data, hIdx, globalStats, targetDistribution, weights);
        const errC2After = classStates[c2].simulateSwap(b, a, data, hIdx, globalStats, targetDistribution, weights);
        const errC3After = classStates[c3].simulateSwap(c, b, data, hIdx, globalStats, targetDistribution, weights);

        const gain3 = errBefore3 - (errC1After + errC2After + errC3After);
        if (gain3 > bestGain3) {
          bestGain3 = gain3;
          best3Way = { a, b, c, c1, c2, c3 };
        }
      }
    }

    if (!best3Way) break;

    const { a, b, c, c1, c2, c3 } = best3Way;
    classStates[c1].applySwap(a, c, data, hIdx);
    classStates[c2].applySwap(b, a, data, hIdx);
    classStates[c3].applySwap(c, b, data, hIdx);

    data[a][hIdx.ASSIGNED] = c2;
    data[b][hIdx.ASSIGNED] = c3;
    data[c][hIdx.ASSIGNED] = c1;

    byClass[c1] = byClass[c1].filter(x => x !== a).concat([c]);
    byClass[c2] = byClass[c2].filter(x => x !== b).concat([a]);
    byClass[c3] = byClass[c3].filter(x => x !== c).concat([b]);

    swaps3Way++;
    swapsApplied++;
  }

  return {
    swapsApplied: swapsApplied,
    swaps3Way: swaps3Way,
    annealingUsed: annealingUsed,
    finalError: totalError()
  };
}

/**
 * Calcule la distribution cible des scores pour chaque critère.
 * @returns {Object} ex: { COM: { '1': 0.1, '2': 0.2, ... }, ... }
 */
function calculateTargetDistribution_V3(data, headers, byClass) {
  const criteria = ['COM', 'TRA', 'PART', 'ABS'];
  const totalStudents = data.length - 1;
  const globalCounts = {};

  criteria.forEach(crit => {
    const idx = headers.indexOf(crit);
    globalCounts[crit] = { '1': 0, '2': 0, '3': 0, '4': 0 };
    for (let i = 1; i < data.length; i++) {
      const score = String(data[i][idx] || '3');
      if (globalCounts[crit][score]) {
        globalCounts[crit][score]++;
      } else {
        globalCounts[crit][score] = 1; // Handle potential other values
      }
    }
  });

  const targetDistribution = {};
  criteria.forEach(crit => {
    targetDistribution[crit] = {};
    for (let s = 1; s <= 4; s++) {
      targetDistribution[crit][s] = globalCounts[crit][s] / totalStudents;
    }
  });

  return targetDistribution;
}

/**
 * Calcule "l'erreur d'harmonie" pour un critère et une répartition donnés.
 */
function calculateHarmonyError_V3(byClass, data, headers, criterion, targetDistribution) {
  const idx = headers.indexOf(criterion);
  let totalError = 0;

  for (const cls in byClass) {
    const classSize = byClass[cls].length;
    const currentCounts = { '1': 0, '2': 0, '3': 0, '4': 0 };

    byClass[cls].forEach(studentIdx => {
      const score = String(data[studentIdx][idx] || '3');
      if (currentCounts[score]) currentCounts[score]++; else currentCounts[score] = 1;
    });

    for (let s = 1; s <= 4; s++) {
      const targetCount = targetDistribution[criterion][s] * classSize;
      totalError += Math.abs(currentCounts[s] - targetCount);
    }
  }
  return totalError;
}

/**
 * Calcule le score de parité (somme des |F-M|).
 */
function calculateParityError_V3(byClass, data, headers) {
  const idxSexe = headers.indexOf('SEXE');
  let totalParityGap = 0;
  for (const cls in byClass) {
    let countF = 0, countM = 0;
    byClass[cls].forEach(idx => {
      if (String(data[idx][idxSexe] || '').toUpperCase() === 'F') countF++; else countM++;
    });
    totalParityGap += Math.abs(countF - countM);
  }
  return totalParityGap;
}

/**
 * Calcule le score composite d'une répartition (erreur totale à minimiser).
 */
function calculateCompositeSwapScore_V3(data, headers, byClass, targetDistribution, weights, swap) {
  const criteria = HARMONY_CRITERIA;
  let totalError = 0;

  if (swap) { // Simulation rapide
    const { idx1, idx2 } = swap;
    const cls1 = String(data[idx1][headers.indexOf('_CLASS_ASSIGNED')]);
    const cls2 = String(data[idx2][headers.indexOf('_CLASS_ASSIGNED')]);

    const tempByClass = JSON.parse(JSON.stringify(byClass));
    const pos1 = tempByClass[cls1].indexOf(idx1);
    const pos2 = tempByClass[cls2].indexOf(idx2);
    if(pos1 !== -1) tempByClass[cls1].splice(pos1, 1, idx2);
    if(pos2 !== -1) tempByClass[cls2].splice(pos2, 1, idx1);
    byClass = tempByClass;
  }

  // Erreur de parité
  totalError += calculateParityError_V3(byClass, data, headers) * (weights.parity || 1.0);

  // Erreurs d'harmonie (distribution des scores 1-4)
  criteria.forEach(crit => {
    totalError += calculateHarmonyError_V3(byClass, data, headers, crit, targetDistribution) * (weights[crit.toLowerCase()] || 0.1);
  });

  // HARMONY FIX : Ajouter pénalité pour têtes de classe et niv1 mal répartis
  const idxCOM = headers.indexOf('COM');
  const idxTRA = headers.indexOf('TRA');
  const idxPART = headers.indexOf('PART');
  const profileWeight = weights.profiles || 2.0;

  if (idxCOM >= 0 && idxTRA >= 0) {
    for (const cls in byClass) {
      const indices = byClass[cls];
      const clsSize = indices.length;
      if (clsSize === 0) continue;

      let nbTetes = 0, nbNiv1 = 0;
      for (let k = 0; k < clsSize; k++) {
        const com = Number(data[indices[k]][idxCOM] || 2.5);
        const tra = Number(data[indices[k]][idxTRA] || 2.5);
        const part = idxPART >= 0 ? Number(data[indices[k]][idxPART] || 2.5) : 2.5;
        if (isHeadStudent(com, tra, part)) nbTetes++;
        if (isNiv1Student(com, tra)) nbNiv1++;
      }

      // Cibles proportionnelles : chaque classe devrait avoir ~même ratio
      const targetHeadMin = 2;
      const targetHeadMax = Math.max(5, Math.ceil(clsSize * 0.2));
      const targetNiv1Max = Math.max(4, Math.ceil(clsSize * 0.15));

      // Pénalité asymétrique (manque de têtes = critique, excès de niv1 = critique)
      if (nbTetes < targetHeadMin) {
        totalError += Math.pow(targetHeadMin - nbTetes, 2) * profileWeight;
      }
      if (nbTetes > targetHeadMax) {
        totalError += (nbTetes - targetHeadMax) * profileWeight * 0.5;
      }
      if (nbNiv1 > targetNiv1Max) {
        totalError += Math.pow(nbNiv1 - targetNiv1Max, 2) * profileWeight;
      }
    }
  }

  return totalError;
}


/**
 * Trouve le meilleur swap possible en maximisant le gain sur le score composite.
 * Version "Nautilus" : utilise une recherche ciblée et une pénalité de stabilité.
 * @param {Map} swapHistory - Garde une trace du nombre de fois où chaque élève a été échangé.
 */
function findBestSwap_V3(data, headers, byClass, targetDistribution, weights, ctx, swapHistory) {
  const idxMobilite = headers.indexOf('MOBILITE');
  const idxFixe = headers.indexOf('FIXE');
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxSexe = headers.indexOf('SEXE');

  let bestSwap = null;
  let bestCompositeGain = 1e-6;

  const errorBefore = calculateCompositeSwapScore_V3(data, headers, byClass, targetDistribution, weights, null);
  const studentsWithDisruption = [];

  // Moteurs Silencieux - Étape 1: Identifier les élèves "perturbateurs" dans chaque classe.
  for (const cls in byClass) {
    byClass[cls].forEach(studentIdx => {
      const disruptionScore = calculateStudentDisruption(studentIdx, cls, data, headers, byClass, targetDistribution, weights);
      studentsWithDisruption.push({ studentIdx, cls, disruptionScore });
    });
  }

  studentsWithDisruption.sort((a, b) => b.disruptionScore - a.disruptionScore);

  const topPercent = 0.6; // HARMONY FIX : Élargir le pool de 40% à 60%
  const poolSize = Math.max(10, Math.floor(studentsWithDisruption.length * topPercent));
  const candidateStudents = studentsWithDisruption.slice(0, poolSize);

  for (let i = 0; i < candidateStudents.length; i++) {
    const s1 = candidateStudents[i];
    const idx1 = s1.studentIdx;
    const cls1 = s1.cls;

    if (String(data[idx1][idxMobilite] || '').toUpperCase() === 'FIXE' || String(data[idx1][idxFixe] || '').toUpperCase() === 'FIXE') continue;

    for (let j = i + 1; j < candidateStudents.length; j++) {
      const s2 = candidateStudents[j];
      const idx2 = s2.studentIdx;
      const cls2 = s2.cls;

      if (cls1 === cls2) continue;
      if (String(data[idx2][idxMobilite] || '').toUpperCase() === 'FIXE' || String(data[idx2][idxFixe] || '').toUpperCase() === 'FIXE') continue;

      const swapCheck = canSwapStudents_V3(idx1, cls1, idx2, cls2, data, headers, ctx);
      if (!swapCheck.ok) continue;

      const errorAfter = calculateCompositeSwapScore_V3(data, headers, byClass, targetDistribution, weights, { idx1, idx2 });
      let compositeGain = errorBefore - errorAfter;

      // Ancre d'amarrage : appliquer une pénalité de mouvement pour stabiliser l'algorithme.
      // Le gain d'un swap est divisé par un facteur qui augmente à chaque fois
      // qu'un des élèves impliqués est échangé.
      const history1 = swapHistory.get(idx1) || 0;
      const history2 = swapHistory.get(idx2) || 0;
      const penalty = 1 + history1 + history2;
      const penalizedGain = compositeGain / penalty;

      if (penalizedGain > bestCompositeGain) {
        bestCompositeGain = penalizedGain;
        bestSwap = { idx1, idx2, compositeGain: penalizedGain }; // On stocke le gain pénalisé pour la comparaison
      }
    }
  }

  return bestSwap;
}

/**
 * Calcule le "score de perturbation" d'un élève (stratégie "Moteurs Silencieux").
 * Cette fonction mesure à quel point un élève contribue négativement au score composite
 * (l'erreur) de sa classe actuelle. Un score de perturbation élevé signifie que le profil
 * de l'élève s'éloigne significativement de la composition idéale de la classe,
 * faisant de lui un excellent candidat pour un swap.
 *
 * @returns {number} Le score de perturbation (réduction de l'erreur si l'élève est retiré).
 */
function calculateStudentDisruption(studentIdx, cls, data, headers, byClass, targetDistribution, weights) {
    const errorWithStudent = calculateCompositeSwapScore_V3(data, headers, { [cls]: byClass[cls] }, targetDistribution, weights, null);

    // Simuler le retrait de l'élève
    const classWithoutStudent = byClass[cls].filter(idx => idx !== studentIdx);

    // Si la classe devient vide, la perturbation est nulle.
    if (classWithoutStudent.length === 0) return 0;

    const errorWithoutStudent = calculateCompositeSwapScore_V3(data, headers, { [cls]: classWithoutStudent }, targetDistribution, weights, null);

    // La perturbation est la réduction de l'erreur si on enlève l'élève
    return errorWithStudent - errorWithoutStudent;
}


// ===================================================================
// SYNC LEGACY COLUMNS FOR AUDIT COMPATIBILITY
// ===================================================================

/**
 * Synchronise _CLASS_ASSIGNED vers les colonnes legacy (_PLACED, CLASSE_FINAL, _TARGET_CLASS)
 * pour assurer la compatibilité avec les fonctions d'audit existantes.
 *
 * @param {string} phaseLabel - Label de la phase (P1, P2, P3, P4)
 */
function syncClassAssignedToLegacy_(phaseLabel) {
  const ss = SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName('_BASEOPTI');

  if (!baseSheet) {
    logLine('WARN', '⚠️ syncClassAssignedToLegacy_: _BASEOPTI introuvable');
    return;
  }

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];

  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxPlaced = headers.indexOf('_PLACED');
  const idxClasseFinal = headers.indexOf('CLASSE_FINAL');
  const idxTargetClass = headers.indexOf('_TARGET_CLASS');

  if (idxAssigned === -1) {
    logLine('WARN', '⚠️ syncClassAssignedToLegacy_: colonne _CLASS_ASSIGNED introuvable');
    return;
  }

  let synced = 0;

  for (let i = 1; i < data.length; i++) {
    const assigned = String(data[i][idxAssigned] || '').trim();

    if (assigned) {
      // Sync _PLACED
      if (idxPlaced >= 0) {
        data[i][idxPlaced] = phaseLabel;
      }

      // Sync CLASSE_FINAL
      if (idxClasseFinal >= 0) {
        data[i][idxClasseFinal] = assigned;
      }

      // Sync _TARGET_CLASS
      if (idxTargetClass >= 0) {
        data[i][idxTargetClass] = assigned;
      }

      synced++;
    }
  }

  // Écrire les modifications
  baseSheet.getRange(1, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();

  logLine('INFO', '  🔄 Sync legacy: ' + synced + ' élèves (' + phaseLabel + ')');
}

/**
 * ============================================================
 * AUDIT COMPLET DE FIN D'OPTIMISATION
 * ============================================================
 * Génère un rapport détaillé avec :
 * - Répartition par classe (effectifs, parité F/M)
 * - Distribution des scores (COM, TRA, PART, ABS)
 * - Respect des quotas LV2/OPT
 * - Statuts de mobilité (FIXE, PERMUT, LIBRE)
 * - Codes ASSO/DISSO
 * - Métriques de qualité (variance, écart-type)
 *
 * @param {Object} ctx - Contexte d'optimisation
 * @param {Array} data - Données _BASEOPTI
 * @param {Array} headers - En-têtes _BASEOPTI
 * @param {Object} byClass - Index des élèves par classe
 * @param {Object} distributions - Distributions des scores
 * @param {Object} metrics - Métriques d'optimisation
 * @returns {Object} Rapport d'audit complet
 */
function generateOptimizationAudit_V3(ctx, data, headers, byClass, distributions, metrics) {
  logLine('INFO', '');
  logLine('INFO', '═══════════════════════════════════════════════════════');
  logLine('INFO', '📊 AUDIT COMPLET DE FIN D\'OPTIMISATION');
  logLine('INFO', '═══════════════════════════════════════════════════════');
  logLine('INFO', '');

  const report = {
    timestamp: new Date().toISOString(),
    metrics: metrics,
    classes: {},
    global: {
      totalStudents: 0,
      totalFemale: 0,
      totalMale: 0,
      parityRatio: 0
    },
    quotas: {},
    mobility: {
      FIXE: 0,
      PERMUT: 0,
      LIBRE: 0,
      CONFLIT: 0
    },
    codes: {
      ASSO: {},
      DISSO: {}
    },
    quality: {}
  };

  // Indices des colonnes
  const idxNom = headers.indexOf('NOM');
  const idxPrenom = headers.indexOf('PRENOM');
  const idxSexe = headers.indexOf('SEXE');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');
  const idxCOM = headers.indexOf('COM');
  const idxTRA = headers.indexOf('TRA');
  const idxPART = headers.indexOf('PART');
  const idxABS = headers.indexOf('ABS');
  const idxMobilite = headers.indexOf('MOBILITE');
  const idxFixe = headers.indexOf('FIXE');
  const idxAsso = headers.indexOf('ASSO');
  const idxDisso = headers.indexOf('DISSO');
  const idxClassAssigned = headers.indexOf('_CLASS_ASSIGNED');

  // ========== 1. ANALYSE PAR CLASSE ==========
  logLine('INFO', '📋 1. RÉPARTITION PAR CLASSE');
  logLine('INFO', '─────────────────────────────────────────────────────');

  for (const cls in byClass) {
    const indices = byClass[cls];
    const classData = {
      name: cls,
      total: indices.length,
      female: 0,
      male: 0,
      parityRatio: 0,
      scores: {
        COM: { 1: 0, 2: 0, 3: 0, 4: 0, avg: 0 },
        TRA: { 1: 0, 2: 0, 3: 0, 4: 0, avg: 0 },
        PART: { 1: 0, 2: 0, 3: 0, 4: 0, avg: 0 },
        ABS: { 1: 0, 2: 0, 3: 0, 4: 0, avg: 0 }
      },
      lv2: {},
      opt: {},
      mobility: { FIXE: 0, PERMUT: 0, LIBRE: 0, CONFLIT: 0 }
    };

    // Compter par sexe, scores, LV2, OPT, mobilité
    indices.forEach(function(idx) {
      const sexe = String(data[idx][idxSexe] || '').toUpperCase();
      if (sexe === 'F') classData.female++;
      else if (sexe === 'M') classData.male++;

      // Scores
      const com = Number(data[idx][idxCOM] || 3);
      const tra = Number(data[idx][idxTRA] || 3);
      const part = Number(data[idx][idxPART] || 3);
      const abs = Number(data[idx][idxABS] || 3);

      classData.scores.COM[com]++;
      classData.scores.TRA[tra]++;
      classData.scores.PART[part]++;
      classData.scores.ABS[abs]++;

      // LV2 et OPT
      const lv2 = String(data[idx][idxLV2] || '').trim().toUpperCase();
      const opt = String(data[idx][idxOPT] || '').trim().toUpperCase();

      if (lv2 && lv2 !== 'ESP' && lv2 !== 'ANG') {
        classData.lv2[lv2] = (classData.lv2[lv2] || 0) + 1;
      }
      if (opt) {
        classData.opt[opt] = (classData.opt[opt] || 0) + 1;
      }

      // Mobilité
      const mob = String(data[idx][idxMobilite] || 'LIBRE').toUpperCase();
      if (mob.includes('FIXE')) classData.mobility.FIXE++;
      else if (mob.includes('PERMUT')) classData.mobility.PERMUT++;
      else if (mob.includes('CONFLIT')) classData.mobility.CONFLIT++;
      else classData.mobility.LIBRE++;
    });

    // Calculer parité
    classData.parityRatio = classData.total > 0 
      ? (classData.female / classData.total * 100).toFixed(1) 
      : 0;

    // Calculer moyennes des scores
    ['COM', 'TRA', 'PART', 'ABS'].forEach(function(scoreType) {
      let sum = 0;
      let count = 0;
      for (let score = 1; score <= 4; score++) {
        sum += score * classData.scores[scoreType][score];
        count += classData.scores[scoreType][score];
      }
      classData.scores[scoreType].avg = count > 0 ? (sum / count).toFixed(2) : 0;
    });

    report.classes[cls] = classData;
    report.global.totalStudents += classData.total;
    report.global.totalFemale += classData.female;
    report.global.totalMale += classData.male;

    // Log
    logLine('INFO', '  ' + cls + ' : ' + classData.total + ' élèves (' + classData.female + 'F / ' + classData.male + 'M = ' + classData.parityRatio + '% F)');
    logLine('INFO', '    Scores moyens: COM=' + classData.scores.COM.avg + ', TRA=' + classData.scores.TRA.avg + ', PART=' + classData.scores.PART.avg + ', ABS=' + classData.scores.ABS.avg);
    
    if (Object.keys(classData.lv2).length > 0) {
      logLine('INFO', '    LV2: ' + JSON.stringify(classData.lv2));
    }
    if (Object.keys(classData.opt).length > 0) {
      logLine('INFO', '    OPT: ' + JSON.stringify(classData.opt));
    }
    
    logLine('INFO', '    Mobilité: FIXE=' + classData.mobility.FIXE + ', PERMUT=' + classData.mobility.PERMUT + ', LIBRE=' + classData.mobility.LIBRE);
  }

  // Parité globale
  report.global.parityRatio = report.global.totalStudents > 0
    ? (report.global.totalFemale / report.global.totalStudents * 100).toFixed(1)
    : 0;

  logLine('INFO', '');
  logLine('INFO', '  GLOBAL : ' + report.global.totalStudents + ' élèves (' + report.global.totalFemale + 'F / ' + report.global.totalMale + 'M = ' + report.global.parityRatio + '% F)');

  // ========== 2. RESPECT DES QUOTAS ==========
  logLine('INFO', '');
  logLine('INFO', '📊 2. RESPECT DES QUOTAS LV2/OPT');
  logLine('INFO', '─────────────────────────────────────────────────────');

  if (ctx.quotas) {
    for (const cls in ctx.quotas) {
      const quotasDef = ctx.quotas[cls];
      const classData = report.classes[cls];
      
      if (!classData) continue;

      logLine('INFO', '  ' + cls + ' :');

      for (const option in quotasDef) {
        const quota = quotasDef[option];
        const actual = classData.lv2[option.toUpperCase()] || classData.opt[option.toUpperCase()] || 0;
        const status = actual <= quota ? '✅' : '⚠️';
        
        logLine('INFO', '    ' + status + ' ' + option + ' : ' + actual + ' / ' + quota + ' (quota)');
        
        if (!report.quotas[option]) {
          report.quotas[option] = { quota: quota, actual: 0, classes: [] };
        }
        report.quotas[option].actual += actual;
        report.quotas[option].classes.push({ class: cls, count: actual });
      }
    }
  }

  // ========== 3. CODES ASSO/DISSO ==========
  logLine('INFO', '');
  logLine('INFO', '🔗 3. CODES ASSO/DISSO');
  logLine('INFO', '─────────────────────────────────────────────────────');

  for (let i = 1; i < data.length; i++) {
    const asso = String(data[i][idxAsso] || '').trim().toUpperCase();
    const disso = String(data[i][idxDisso] || '').trim().toUpperCase();
    const classe = String(data[i][idxClassAssigned] || '').trim();

    if (asso) {
      if (!report.codes.ASSO[asso]) {
        report.codes.ASSO[asso] = { count: 0, classes: {} };
      }
      report.codes.ASSO[asso].count++;
      report.codes.ASSO[asso].classes[classe] = (report.codes.ASSO[asso].classes[classe] || 0) + 1;
    }

    if (disso) {
      if (!report.codes.DISSO[disso]) {
        report.codes.DISSO[disso] = { count: 0, classes: {} };
      }
      report.codes.DISSO[disso].count++;
      report.codes.DISSO[disso].classes[classe] = (report.codes.DISSO[disso].classes[classe] || 0) + 1;
    }
  }

  // Log ASSO
  const assoKeys = Object.keys(report.codes.ASSO);
  if (assoKeys.length > 0) {
    logLine('INFO', '  ASSO (' + assoKeys.length + ' groupes) :');
    assoKeys.forEach(function(code) {
      const group = report.codes.ASSO[code];
      const classesStr = Object.keys(group.classes).map(function(c) {
        return c + '(' + group.classes[c] + ')';
      }).join(', ');
      const status = Object.keys(group.classes).length === 1 ? '✅' : '⚠️';
      logLine('INFO', '    ' + status + ' ' + code + ' : ' + group.count + ' élèves → ' + classesStr);
    });
  } else {
    logLine('INFO', '  Aucun code ASSO');
  }

  // Log DISSO
  const dissoKeys = Object.keys(report.codes.DISSO);
  if (dissoKeys.length > 0) {
    logLine('INFO', '  DISSO (' + dissoKeys.length + ' codes) :');
    dissoKeys.forEach(function(code) {
      const group = report.codes.DISSO[code];
      const classesStr = Object.keys(group.classes).map(function(c) {
        return c + '(' + group.classes[c] + ')';
      }).join(', ');
      const status = Object.keys(group.classes).length === group.count ? '✅' : '⚠️';
      logLine('INFO', '    ' + status + ' ' + code + ' : ' + group.count + ' élèves → ' + classesStr);
    });
  } else {
    logLine('INFO', '  Aucun code DISSO');
  }

  // ========== 4. MÉTRIQUES DE QUALITÉ ==========
  logLine('INFO', '');
  logLine('INFO', '📈 4. MÉTRIQUES DE QUALITÉ');
  logLine('INFO', '─────────────────────────────────────────────────────');

  // Calculer écart-type des effectifs
  const effectifs = Object.keys(report.classes).map(function(cls) {
    return report.classes[cls].total;
  });
  const avgEffectif = effectifs.reduce(function(a, b) { return a + b; }, 0) / effectifs.length;
  const varianceEffectif = effectifs.reduce(function(sum, val) {
    return sum + Math.pow(val - avgEffectif, 2);
  }, 0) / effectifs.length;
  const ecartTypeEffectif = Math.sqrt(varianceEffectif);

  report.quality.effectifs = {
    avg: avgEffectif.toFixed(2),
    ecartType: ecartTypeEffectif.toFixed(2)
  };

  logLine('INFO', '  Effectifs : moyenne=' + report.quality.effectifs.avg + ', écart-type=' + report.quality.effectifs.ecartType);

  // Calculer écart-type de la parité
  const parities = Object.keys(report.classes).map(function(cls) {
    return parseFloat(report.classes[cls].parityRatio);
  });
  const avgParity = parities.reduce(function(a, b) { return a + b; }, 0) / parities.length;
  const varianceParity = parities.reduce(function(sum, val) {
    return sum + Math.pow(val - avgParity, 2);
  }, 0) / parities.length;
  const ecartTypeParity = Math.sqrt(varianceParity);

  report.quality.parity = {
    avg: avgParity.toFixed(1),
    ecartType: ecartTypeParity.toFixed(2)
  };

  logLine('INFO', '  Parité F/M : moyenne=' + report.quality.parity.avg + '% F, écart-type=' + report.quality.parity.ecartType);

  // Métriques d'optimisation
  logLine('INFO', '  Harmonie académique pondérée : initiale=' + metrics.initialAcademic.toFixed(2) + ', finale=' + metrics.finalAcademic.toFixed(2) + ' (Δ=' + metrics.academicImprovement.toFixed(2) + ')');
  logLine('INFO', '  Parité totale |F-M| : initiale=' + metrics.initialParity.toFixed(2) + ', finale=' + metrics.finalParity.toFixed(2) + ' (Δ=' + metrics.parityImprovement.toFixed(2) + ')');
  logLine('INFO', '  Score composite : initial=' + metrics.initialComposite.toFixed(2) + ', final=' + metrics.finalComposite.toFixed(2) + ' (Δ=' + metrics.compositeImprovement.toFixed(2) + ')');
  logLine('INFO', '  Harmonisation détaillée : COM=' + metrics.distributionImprovements.COM.toFixed(2) + ', TRA=' + metrics.distributionImprovements.TRA.toFixed(2) + ', PART=' + metrics.distributionImprovements.PART.toFixed(2) + ', ABS=' + metrics.distributionImprovements.ABS.toFixed(2));
  logLine('INFO', '  Swaps appliqués : ' + metrics.swapsApplied);

  // ========== 5. SYNTHÈSE ==========
  logLine('INFO', '');
  logLine('INFO', '✅ 5. SYNTHÈSE');
  logLine('INFO', '─────────────────────────────────────────────────────');
  logLine('INFO', '  Classes : ' + Object.keys(report.classes).length);
  logLine('INFO', '  Élèves : ' + report.global.totalStudents + ' (' + report.global.totalFemale + 'F / ' + report.global.totalMale + 'M)');
  logLine('INFO', '  Parité globale : ' + report.global.parityRatio + '% F');
  logLine('INFO', '  Groupes ASSO : ' + assoKeys.length);
  logLine('INFO', '  Codes DISSO : ' + dissoKeys.length);
  const improvementRatio = metrics.initialComposite !== 0
    ? (metrics.compositeImprovement / metrics.initialComposite * 100).toFixed(1)
    : 'N/A';
  logLine('INFO', '  Amélioration composite : ' + improvementRatio + '%');

  logLine('INFO', '');
  logLine('INFO', '═══════════════════════════════════════════════════════');
  logLine('INFO', '');

  return report;
}

