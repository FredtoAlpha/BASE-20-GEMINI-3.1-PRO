/**
 * ===================================================================
 * 🎯 PRIME LEGACY - PHASE 1 : OPTIONS & LV2
 * ===================================================================
 *
 * Basé sur : OPTIMUM PRIME (Phase1I_dispatchOptionsLV2_BASEOPTI_V3)
 * Source : Phases_BASEOPTI_V3_COMPLETE.gs (JULES-VERNE-NAUTILUS)
 *
 * Phase 1 : Place les élèves avec OPT/LV2 selon quotas
 * LIT : CONSOLIDATION (le sac de billes)
 * ÉCRIT : Onglets TEST (élèves avec contraintes uniquement)
 *
 * ISOLATION COMPLÈTE :
 * - OPTI : _BASEOPTI (vivier unique)
 * - LEGACY : CONSOLIDATION → TEST (sélectif) → reste dans sac pour Phase 3
 *
 * Date : 2025-11-13
 * Branche : claude/PRIME-LEGACY-01SJDcJv7zHGGBXWhHpzfnxr
 *
 * ===================================================================
 */

/**
 * Phase 1 LEGACY : Place les élèves avec OPT/LV2 selon quotas
 * LIT : CONSOLIDATION (le sac de billes)
 * ÉCRIT : Onglets TEST (élèves avec contraintes uniquement)
 *
 * @param {Object} ctx - Contexte LEGACY
 * @returns {Object} { ok: true, counts: {...} }
 */
function Phase1I_dispatchOptionsLV2_LEGACY(ctx) {
  logLine('INFO', '='.repeat(80));
  logLine('INFO', '📌 PHASE 1 LEGACY - Options & LV2 (OPTIMUM PRIME)');
  logLine('INFO', '='.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const stats = {};

  // ========== ÉTAPE 1 : LIRE DEPUIS CONSOLIDATION (LE SAC) ==========
  // 🎯 CONSOLIDATION = LE SAC DE BILLES depuis le début !
  logLine('INFO', '🪣 Lecture depuis CONSOLIDATION (le sac de billes)...');

  const consolidationSheet = ss.getSheetByName('CONSOLIDATION');
  
  if (!consolidationSheet || consolidationSheet.getLastRow() <= 1) {
    logLine('ERROR', '❌ CONSOLIDATION vide ou introuvable !');
    return { ok: false, counts: stats };
  }

  const data = consolidationSheet.getDataRange().getValues();
  let headersRef = data[0];
  
  // ✅ Ajouter colonne _ELEVE_PLACE si absente (pour tracking)
  const idxElevePlace = headersRef.indexOf('_ELEVE_PLACE');
  if (idxElevePlace === -1) {
    const lastCol = consolidationSheet.getLastColumn();
    consolidationSheet.getRange(1, lastCol + 1).setValue('_ELEVE_PLACE')
      .setBackground('#FFD966').setFontWeight('bold');
    headersRef.push('_ELEVE_PLACE');
    logLine('INFO', '  ✨ Colonne _ELEVE_PLACE créée pour tracking');
  } else {
    // Colonne existe (relance) → La vider
    const lastRow = consolidationSheet.getLastRow();
    if (lastRow > 1) {
      consolidationSheet.getRange(2, idxElevePlace + 1, lastRow - 1, 1).clearContent();
      logLine('INFO', '  🧹 Colonne _ELEVE_PLACE vidée (relance du pipeline)');
    }
  }
  
  // ✅ Nettoyer les headers pour exclure _ELEVE_PLACE (colonne temporaire)
  const cleanHeaders = [];
  for (let h = 0; h < headersRef.length; h++) {
    if (h !== idxElevePlace && headersRef[h] !== '_ELEVE_PLACE') {
      cleanHeaders.push(headersRef[h]);
    }
  }
  headersRef = cleanHeaders; // Utiliser les headers nettoyés
  
  // ✅ Lire les données en excluant _ELEVE_PLACE
  const allData = [];
  for (let i = 1; i < data.length; i++) {
    const rawRow = data[i];
    
    // Nettoyer la ligne (exclure _ELEVE_PLACE)
    const cleanRow = [];
    for (let c = 0; c < rawRow.length; c++) {
      if (c !== idxElevePlace) {
        cleanRow.push(rawRow[c]);
      }
    }
    
    allData.push({
      sheetName: 'CONSOLIDATION',
      row: cleanRow // 15 colonnes de base uniquement
    });
  }

  logLine('INFO', '  ✅ ' + allData.length + ' élèves lus depuis CONSOLIDATION (le sac)');

  // ========== ÉTAPE 2 : TROUVER LES INDEX DES COLONNES ==========
  const idxLV2 = headersRef.indexOf('LV2');
  const idxOPT = headersRef.indexOf('OPT');
  const idxNom = headersRef.indexOf('NOM');
  const idxPrenom = headersRef.indexOf('PRENOM');

  // ✅ CORRECTION : _CLASS_ASSIGNED n'existe PAS dans les sources
  //    On va l'ajouter dynamiquement pour chaque élève
  logLine('INFO', '  📍 Colonnes sources : LV2=' + idxLV2 + ', OPT=' + idxOPT);
  
  // Ajouter _CLASS_ASSIGNED à chaque élève (nouvelle colonne)
  for (let i = 0; i < allData.length; i++) {
    allData[i].assigned = ''; // Nouvelle propriété pour stocker l'affectation
  }

  // ========== ÉTAPE 3 : DÉTECTION LV2 UNIVERSELLES ==========
  // 🌟 APPROCHE UNIVERSELLE : Détecter les LV2 présentes dans TOUTES les classes
  const allClasses = Object.keys(ctx.quotas || {});
  const nbClasses = allClasses.length;
  const lv2Counts = {}; // Compte combien de classes proposent chaque LV2
  
  for (const classe in (ctx.quotas || {})) {
    const quotas = ctx.quotas[classe];
    for (const optName in quotas) {
      if (isKnownLV2(optName)) {
        if (quotas[optName] > 0) {
          lv2Counts[optName] = (lv2Counts[optName] || 0) + 1;
        }
      }
    }
  }
  
  // LV2 universelles = présentes dans TOUTES les classes
  const lv2Universelles = [];
  for (const lv2 in lv2Counts) {
    if (lv2Counts[lv2] === nbClasses) {
      lv2Universelles.push(lv2);
    }
  }
  
  logLine('INFO', '  🌍 LV2 universelles (dans toutes les classes) : ' + (lv2Universelles.length > 0 ? lv2Universelles.join(', ') : 'aucune'));
  logLine('INFO', '  🎯 LV2 rares (placement Phase 1) : ' + Object.keys(lv2Counts).filter(lv2 => lv2Counts[lv2] < nbClasses).join(', '));
  
  // ✅ Compter les effectifs déjà placés par classe
  const classeCounts = {};
  for (const classe in (ctx.quotas || {})) {
    classeCounts[classe] = 0;
  }

  // Parcourir les quotas par classe
  for (const classe in (ctx.quotas || {})) {
    const quotas = ctx.quotas[classe];
    const targetEffectif = (ctx.targets && ctx.targets[classe]) || 27; // Effectif cible

    for (const optName in quotas) {
      const quota = quotas[optName];
      if (quota <= 0) continue;

      let placed = 0;

      // Parcourir tous les élèves consolidés
      for (let i = 0; i < allData.length; i++) {
        // ✅ CORRECTION : Vérifier effectif cible AVANT de placer
        if (classeCounts[classe] >= targetEffectif) {
          logLine('WARN', '  ⚠️ ' + classe + ' : effectif cible atteint (' + targetEffectif + '), arrêt placement ' + optName);
          break;
        }
        
        if (placed >= quota) break;

        const item = allData[i];
        const row = item.row;

        // ✅ Utiliser la propriété assigned au lieu de row[idxAssigned]
        if (item.assigned) continue; // Déjà placé

        const lv2 = String(row[idxLV2] || '').trim().toUpperCase();
        const opt = String(row[idxOPT] || '').trim().toUpperCase();

        let match = false;
        // 🌟 APPROCHE UNIVERSELLE : Ignorer les LV2 universelles (présentes dans toutes les classes)
        if (isKnownLV2(optName)) {
          // Placer uniquement si LV2 "rare" (pas universelle)
          if (lv2Universelles.indexOf(optName) === -1) {
            match = (lv2 === optName);
          }
        } else if (isKnownOPT(optName)) {
          match = (opt === optName);
        }

        if (match) {
          // ✅ COMPATIBILITÉ TOTALE : Vérifier que la classe supporte TOUTES les options de l'élève
          let compatible = true;
          
          // Vérifier LV2 (si l'élève en a une et qu'elle n'est pas universelle)
          if (lv2 && lv2Universelles.indexOf(lv2) === -1) {
            // L'élève a une LV2 rare → La classe doit la proposer
            if (!quotas[lv2] || quotas[lv2] <= 0) {
              compatible = false;
            }
          }
          
          // Vérifier OPT (si l'élève en a une)
          if (compatible && opt && isKnownOPT(opt)) {
            // L'élève a une option → La classe doit la proposer
            if (!quotas[opt] || quotas[opt] <= 0) {
              compatible = false;
            }
          }
          
          if (compatible) {
            // ✅ PLACER : Toutes les options de l'élève sont supportées par la classe
            item.assigned = classe;
            placed++;
            classeCounts[classe]++; // Incrémenter le compteur de la classe
            stats[optName] = (stats[optName] || 0) + 1;

            const nom = String(row[idxNom] || '');
            const prenom = String(row[idxPrenom] || '');
            logLine('INFO', '    ✅ ' + nom + ' ' + prenom + ' → ' + classe + ' (' + optName + ') [' + classeCounts[classe] + '/' + targetEffectif + ']');
          } else {
            // ❌ INCOMPATIBLE : L'élève a d'autres contraintes non supportées
            // On le laisse pour une autre classe qui supporte tout
            const nom = String(row[idxNom] || '');
            const prenom = String(row[idxPrenom] || '');
            logLine('INFO', '    ⏭️ ' + nom + ' ' + prenom + ' : ' + optName + ' mais incompatible avec ' + classe + ' (a aussi: LV2=' + lv2 + ', OPT=' + opt + ')');
          }
        }
      }

      if (placed > 0) {
        logLine('INFO', '  ✅ ' + classe + ' : ' + placed + ' × ' + optName + (placed < quota ? ' (⚠️ quota=' + quota + ')' : ''));
      }
    }
  }

  // ========== ÉTAPE 4 : ÉCRIRE DANS LES ONGLETS TEST ==========
  logLine('INFO', '📋 Écriture dans les onglets TEST...');

  // 🎯 ARCHITECTURE "SAC DE BILLES" : N'écrire QUE les élèves avec contraintes
  // Les élèves ESP sans option restent dans CONSOLIDATION (le "sac") pour Phase 3
  const byClass = {};
  
  for (let i = 0; i < allData.length; i++) {
    const item = allData[i];
    
    // ✅ N'écrire QUE les élèves assignés (LV2 rares + Options)
    if (item.assigned) {
      if (!byClass[item.assigned]) {
        byClass[item.assigned] = [];
      }
      // Structure P=FIXE, Q=MOBILITE, R=_CLASS_ASSIGNED
      const newRow = item.row.concat(['', '', item.assigned]); // FIXE vide, MOBILITE vide, _CLASS_ASSIGNED
      byClass[item.assigned].push(newRow);
    }
    // Les autres restent dans CONSOLIDATION (le sac) → Phase 3 les récupérera
  }

  // Écrire dans les onglets TEST correspondants
  for (const classe in byClass) {
    const testName = classe + 'TEST';
    const testSheet = ss.getSheetByName(testName);

    if (!testSheet) {
      logLine('WARN', '⚠️ Onglet TEST ' + testName + ' introuvable, skip');
      continue;
    }

    const rows = byClass[classe];

    // Écrire les données (à partir de la ligne 2)
    if (rows.length > 0) {
      // Les onglets TEST ont déjà la colonne _CLASS_ASSIGNED en dernière position
      testSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      logLine('INFO', '  ✅ ' + testName + ' : ' + rows.length + ' élèves écrits');
    }
  }

  SpreadsheetApp.flush();

  // ========== ÉTAPE 5 : CALCUL MOBILITÉ ==========
  // ✅ CALCUL MOBILITÉ : Déterminer FIXE/PERMUT/LIBRE après Phase 1
  if (typeof calculerEtRemplirMobilite_LEGACY === 'function') {
    calculerEtRemplirMobilite_LEGACY(ctx);
  } else {
    logLine('WARN', '⚠️ calculerEtRemplirMobilite_LEGACY() non disponible (vérifier que LEGACY_Mobility_Calculator.js est chargé)');
  }

  // Calculer le total des élèves placés
  let totalPlaced = 0;
  for (const classe in byClass) {
    totalPlaced += byClass[classe].length;
  }

  // ========== ÉTAPE 6 : MARQUER LES ÉLÈVES PLACÉS DANS CONSOLIDATION ==========
  // 🏷️ Marquer en batch pour performance
  logLine('INFO', '🏷️ Marquage des élèves placés dans CONSOLIDATION...');
  
  // ✅ Relire les headers de CONSOLIDATION (avec _ELEVE_PLACE)
  const consolidationHeaders = consolidationSheet.getRange(1, 1, 1, consolidationSheet.getLastColumn()).getValues()[0];
  const idxElevePlaceInSheet = consolidationHeaders.indexOf('_ELEVE_PLACE');
  const idxIDInData = headersRef.indexOf('ID_ELEVE');
  const idxIDInSheet = consolidationHeaders.indexOf('ID_ELEVE');
  
  if (idxElevePlaceInSheet !== -1 && idxIDInData !== -1 && idxIDInSheet !== -1) {
    const updates = [];
    const consolidationData = consolidationSheet.getDataRange().getValues();
    
    // Pour chaque élève placé, trouver sa ligne dans CONSOLIDATION
    for (let i = 0; i < allData.length; i++) {
      if (allData[i].assigned) {
        const idEleve = String(allData[i].row[idxIDInData] || '').trim();
        
        // Trouver la ligne correspondante dans CONSOLIDATION
        for (let r = 1; r < consolidationData.length; r++) {
          const idInSheet = String(consolidationData[r][idxIDInSheet] || '').trim();
          if (idInSheet === idEleve) {
            updates.push({
              row: r + 1, // +1 car r=0 est la ligne 1 (header est 0 dans data mais 1 dans sheet)
              value: 'P1:' + allData[i].assigned // Ex: "P1:5°1"
            });
            break;
          }
        }
      }
    }
    
    // Écrire en batch
    if (updates.length > 0) {
      updates.forEach(function(upd) {
        consolidationSheet.getRange(upd.row, idxElevePlaceInSheet + 1).setValue(upd.value);
      });
      logLine('INFO', '  ✅ ' + updates.length + ' élèves marqués dans CONSOLIDATION');
    }
  }
  
  logLine('INFO', '✅ PHASE 1 LEGACY terminée : ' + totalPlaced + ' élèves placés');

  return { ok: true, counts: stats, placed: totalPlaced };
}
