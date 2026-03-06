/**
 * ===================================================================
 * 🔗 PRIME LEGACY - PHASE 2 : ASSO/DISSO
 * ===================================================================
 *
 * Basé sur : OPTIMUM PRIME (Phase2I_applyDissoAsso_BASEOPTI_V3)
 * Source : Phases_BASEOPTI_V3_COMPLETE.gs (lignes 134-295)
 *
 * Phase 2 : Applique codes A (regrouper) et D (séparer)
 * LIT : Onglets TEST (tous les élèves)
 * ÉCRIT : Onglets TEST (update _CLASS_ASSIGNED)
 *
 * Date : 2025-11-13
 * Branche : claude/prime-legacy-cleanup-015Zz6D3gh1QcbpR19TUYMLw
 *
 * ===================================================================
 */

/**
 * Phase 2 LEGACY : Applique codes ASSO/DISSO
 * ✅ IMPLÉMENTATION COMPLÈTE basée sur OPTIMUM PRIME V3
 *
 * @param {Object} ctx - Contexte LEGACY
 * @returns {Object} { ok: true, asso: X, disso: Y }
 */
function Phase2I_applyDissoAsso_LEGACY(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 2 LEGACY - ASSO/DISSO (OPTIMUM PRIME)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();

  // ========== ÉTAPE 1 : CONSOLIDER DONNÉES (SAC DE BILLES) ==========
  // 🎯 Fusionner TEST (déjà placés) + SOURCE (encore dans le sac)
  logLine('INFO', '📋 Consolidation SAC DE BILLES (TEST + SOURCE)');

  const consolidated = getConsolidatedData_LEGACY(ctx);
  const allData = consolidated.allData;
  const headersRef = consolidated.headersRef;
  
  if (allData.length === 0) {
    logLine('WARN', '⚠️ Aucun élève trouvé');
    return { ok: true, asso: 0, disso: 0 };
  }
  
  // Ajouter métadonnées manquantes pour compatibilité
  for (let i = 0; i < allData.length; i++) {
    if (!allData[i].headers) {
      allData[i].headers = headersRef;
    }
    if (allData[i].rowIndex === undefined) {
      allData[i].rowIndex = i + 1;
    }
  }

  logLine('INFO', '  ✅ ' + allData.length + ' élèves consolidés');

  // ========== ÉTAPE 2 : TROUVER INDEX DES COLONNES ==========
  const idxA = headersRef.indexOf('ASSO');
  const idxD = headersRef.indexOf('DISSO');
  const idxAssigned = headersRef.indexOf('_CLASS_ASSIGNED');
  const idxNom = headersRef.indexOf('NOM');
  const idxPrenom = headersRef.indexOf('PRENOM');
  const idxFIXE = headersRef.indexOf('FIXE');
  const idxMOBILITE = headersRef.indexOf('MOBILITE');

  if (idxAssigned === -1) {
    throw new Error('Colonne _CLASS_ASSIGNED manquante');
  }

  let assoMoved = 0;
  let dissoMoved = 0;

  // ========== ÉTAPE 3 : CODES ASSO (A) ==========
  const groupsA = {};
  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const codeA = String(item.row[idxA] || '').trim().toUpperCase();
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
    indices.forEach(function(i) {
      const cls = String(allData[i].row[idxAssigned] || '').trim();
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
      targetClass = findLeastPopulatedClass_LEGACY(allData, headersRef, ctx);
    }

    logLine('INFO', '    🎯 Cible : ' + targetClass);

    // Déplacer tous vers la cible (sauf élèves FIXE)
    indices.forEach(function(i) {
      const item = allData[i];
      const currentClass = String(item.row[idxAssigned] || '').trim();
      
      // ✅ RESPECT COLONNE P : Ne pas déplacer les élèves FIXE ou GROUPE_FIXE
      const fixe = String(item.row[idxFIXE] || '').trim().toUpperCase();
      const mobilite = String(item.row[idxMOBILITE] || '').trim().toUpperCase();
      
      if (fixe === 'OUI' || mobilite === 'FIXE' || mobilite === 'GROUPE_FIXE') {
        const nom = String(item.row[idxNom] || '');
        logLine('WARN', '      ⚠️ ' + nom + ' est FIXE, ne peut être déplacé pour ASSO');
        return; // Skip cet élève
      }
      
      if (currentClass !== targetClass) {
        item.row[idxAssigned] = targetClass;
        assoMoved++;

        const nom = String(item.row[idxNom] || '');
        const prenom = String(item.row[idxPrenom] || '');
        logLine('INFO', '      ✅ ' + nom + ' ' + prenom + ' : ' + currentClass + ' → ' + targetClass);
      }
    });
  }

  // ========== ÉTAPE 4 : CODES DISSO (D) ==========
  // MULTI-RESTART: Tri par densité de contraintes (groupes les plus gros d'abord)
  const groupsD = {};
  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const codeD = String(item.row[idxD] || '').trim().toUpperCase();
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
    indices.forEach(function(i) {
      const cls = String(allData[i].row[idxAssigned] || '').trim();
      if (cls) {
        if (!byClass[cls]) byClass[cls] = [];
        byClass[cls].push(i);
      }
    });

    // Pour chaque classe avec >1 élève D, déplacer
    for (const cls in byClass) {
      if (byClass[cls].length > 1) {
        logLine('INFO', '    ⚠️ ' + cls + ' contient ' + byClass[cls].length + ' D=' + code);

        // Garder le premier, déplacer les autres
        for (let j = 1; j < byClass[cls].length; j++) {
          const i = byClass[cls][j];
          const item = allData[i];
          
          // ✅ RESPECT COLONNE P : Ne pas déplacer les élèves FIXE
          const fixe = String(item.row[idxFIXE] || '').trim().toUpperCase();
          const mobilite = String(item.row[idxMOBILITE] || '').trim().toUpperCase();
          const nom = String(item.row[idxNom] || '');
          
          if (fixe === 'OUI' || mobilite === 'FIXE' || mobilite === 'GROUPE_FIXE') {
            logLine('WARN', '      ⚠️ ' + nom + ' est FIXE, ne peut être déplacé pour DISSO (conflit accepté)');
            continue; // Skip cet élève
          }

          // 🔒 Trouver classe sans ce code D
          const targetClass = findClassWithoutCodeD_LEGACY(allData, headersRef, code, groupsD[code], i, ctx);

          if (targetClass) {
            item.row[idxAssigned] = targetClass;

            const prenom = String(item.row[idxPrenom] || '');
            logLine('INFO', '      ✅ ' + nom + ' ' + prenom + ' : ' + cls + ' → ' + targetClass + ' (séparation D=' + code + ')');
          } else {
            logLine('WARN', '      ⚠️ ' + nom + ' reste en ' + cls + ' (contrainte LV2/OPT absolue)');
          }
        }
      }
    }
  }

  // ========== ÉTAPE 5 : RÉÉCRIRE PAR CLASSE ASSIGNÉE ==========
  // ✅ CORRECTION : Regrouper par _CLASS_ASSIGNED pour que les ASSO/DISSO soient effectifs
  logLine('INFO', '📋 Réécriture dans les onglets TEST...');

  const byClass = {};
  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    const assigned = String(item.row[idxAssigned] || '').trim();
    if (assigned) {
      if (!byClass[assigned]) byClass[assigned] = [];
      byClass[assigned].push(item.row);
    }
  }

  // Écrire dans les onglets TEST correspondants
  for (const className in byClass) {
    const testSheetName = className + 'TEST';
    const testSheet = ss.getSheetByName(testSheetName);
    if (!testSheet) {
      logLine('WARN', '⚠️ Onglet ' + testSheetName + ' introuvable, skip');
      continue;
    }

    const rows = byClass[className];
    const allRows = [headersRef].concat(rows);

    // Effacer contenu existant et écrire nouvelles données
    testSheet.clearContents();
    testSheet.getRange(1, 1, allRows.length, headersRef.length).setValues(allRows);
    
    logLine('INFO', '  ✅ ' + testSheetName + ' : ' + rows.length + ' élèves');
  }

  SpreadsheetApp.flush();

  logLine('INFO', '✅ PHASE 2 LEGACY terminée : ' + assoMoved + ' ASSO, ' + dissoMoved + ' DISSO');

  return { ok: true, asso: assoMoved, disso: dissoMoved };
}

// ===================================================================
// HELPERS LEGACY
// ===================================================================

function findLeastPopulatedClass_LEGACY(allData, headers, ctx) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const counts = {};

  (ctx.niveaux || []).forEach(function(cls) {
    counts[cls] = 0;
  });

  for (let i = 0; i < allData.length; i++) {
    const cls = String(allData[i].row[idxAssigned] || '').trim();
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

  return minClass || (ctx.niveaux && ctx.niveaux[0]) || '6°1';
}

function findClassWithoutCodeD_LEGACY(allData, headers, codeD, indicesWithD, eleveIdx, ctx) {
  const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');
  const idxLV2 = headers.indexOf('LV2');
  const idxOPT = headers.indexOf('OPT');

  const eleveLV2 = eleveIdx >= 0 ? String(allData[eleveIdx].row[idxLV2] || '').trim().toUpperCase() : '';
  var eleveOPT = eleveIdx >= 0 ? String(allData[eleveIdx].row[idxOPT] || '').trim().toUpperCase() : '';

  // Détection anomalie saisie : OPT contient une LV2 (ex: OPT=ITA)
  if (isOPTAnomalyLV2(eleveOPT)) {
    logLine('WARN', '⚠️ DISSO: OPT=' + eleveOPT + ' est une LV2 (anomalie saisie), ignoré');
    eleveOPT = '';
  }

  const classesWithD = new Set();
  indicesWithD.forEach(function(i) {
    const cls = String(allData[i].row[idxAssigned] || '').trim();
    if (cls) classesWithD.add(cls);
  });

  const allClasses = new Set();
  for (let i = 0; i < allData.length; i++) {
    const cls = String(allData[i].row[idxAssigned] || '').trim();
    if (cls) allClasses.add(cls);
  }

  // ✅ CORRECTION BUG #2 : Compter effectifs actuels pour chaque classe
  const classCounts = {};
  for (let i = 0; i < allData.length; i++) {
    const cls = String(allData[i].row[idxAssigned] || '').trim();
    if (cls) {
      classCounts[cls] = (classCounts[cls] || 0) + 1;
    }
  }

  if (eleveLV2 || eleveOPT) {
    for (const cls of Array.from(allClasses)) {
      if (classesWithD.has(cls)) continue;

      // ✅ Vérifier effectif cible
      const targetEffectif = (ctx && ctx.targets && ctx.targets[cls]) || 27;
      const currentCount = classCounts[cls] || 0;
      if (currentCount >= targetEffectif) {
        logLine('INFO', '        ⚠️ Classe ' + cls + ' pleine (' + currentCount + '/' + targetEffectif + '), skip');
        continue;
      }

      const quotas = (ctx && ctx.quotas && ctx.quotas[cls]) || {};

      // ✅ FIX #1 : Vérifier LV2 ET OPT cumulativement (pas else if)
      // Un élève peut avoir LV2=ITA + OPT=LATIN simultanément
      let canPlace = true;
      if (eleveLV2 && isKnownLV2(eleveLV2)) {
        if (!(quotas[eleveLV2] !== undefined && quotas[eleveLV2] > 0)) canPlace = false;
      }
      if (eleveOPT && isKnownOPT(eleveOPT)) {
        if (!(quotas[eleveOPT] !== undefined && quotas[eleveOPT] > 0)) canPlace = false;
      }
      // Si ni LV2 ni OPT, l'élève est "simple" → canPlace reste true
      if (!eleveLV2 && !eleveOPT) canPlace = true;

      if (canPlace) {
        logLine('INFO', '        ✅ Classe ' + cls + ' compatible (propose ' + (eleveLV2 || eleveOPT) + ') [' + currentCount + '/' + targetEffectif + ']');
        return cls;
      }
    }

    logLine('WARN', '        ⚠️ Aucune classe sans D=' + codeD + ' ne propose ' + (eleveLV2 || eleveOPT) + ' avec place disponible');
    return null;
  }

  // ✅ FAILLE #1 CORRIGÉE : Fallback dangereux SUPPRIMÉ
  // Mieux vaut garder un conflit DISSO que perdre l'option CHAV/LATIN
  // Si aucune classe compatible n'est trouvée, l'élève reste dans sa classe actuelle
  return null;
}
