/**
 * ===================================================================
 * 🚀 PRIME LEGACY - PIPELINE PRINCIPAL (GOOGLE APPS SCRIPT)
 * ===================================================================
 *
 * Backend Apps Script pour le pipeline LEGACY classique.
 * Utilise Phase4_Ultimate avec Asymmetric Weighting.
 *
 * ARCHITECTURE :
 * - LECTURE : Onglets sources (6°1, 5°2, etc.)
 * - TRAITEMENT : Phase4_Ultimate (moteur intelligent)
 * - ÉCRITURE : Onglets TEST et FIN
 *
 * ISOLATION COMPLÈTE :
 * - LEGACY : Sources → TEST → FIN
 * - OPTI : _BASEOPTI → _CACHE → FIN
 * - ZÉRO INTERFÉRENCE : Onglets différents, sécurisé
 *
 * Date: 19/11/2025
 * Moteur: Phase4_Ultimate.gs (Asymmetric Weighting)
 * ===================================================================
 */

// ===================================================================
// CONFIGURATION PIPELINE LEGACY
// ===================================================================

const LEGACY_PIPELINE_CONFIG = {
  maxRuntime: 360,        // 6 minutes max (limite GAS)
  enableLogging: true,
  testSheetSuffix: 'TEST',
  finSheetSuffix: 'FIN',
  logLevel: 'INFO'
};

// ===================================================================
// 🚀 POINT D'ENTRÉE PRINCIPAL - APPEL DEPUIS MENU
// ===================================================================

/**
 * Lance le pipeline LEGACY complet
 *
 * APPELÉ PAR: Code.gs → Menu "🚀 PILOTAGE CLASSE"
 *
 * WORKFLOW:
 * 1. Détecter sources (6°1, 5°2, 4°3, etc.)
 * 2. Charger élèves avec profils (Têtes/Niv1)
 * 3. Lancer Phase 4 ULTIMATE
 * 4. Créer onglets TEST
 * 5. Créer onglets FIN (formatés)
 * 6. Afficher résumé
 *
 * @returns {Object} Résultat du pipeline
 */
function legacy_runFullPipeline_PRIME() {
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date();

  logLine('INFO', '═'.repeat(80));
  logLine('INFO', '🚀 LANCEMENT PIPELINE LEGACY PRIME');
  logLine('INFO', '📦 Moteur: OPTIMUM PRIME ULTIMATE (Asymmetric Weighting)');
  logLine('INFO', '═'.repeat(80));

  try {
    // 1. VÉRIFICATION LOCK
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      logLine('WARN', '🔒 Pipeline verrouillé');
      ui.alert('⚠️ Une optimisation est déjà en cours. Veuillez patienter.');
      return { success: false, locked: true };
    }

    // 2. CONSTRUIRE CONTEXTE COMPLET depuis _STRUCTURE
    // ✅ CORRECTION : Utiliser makeCtxFromSourceSheets_LEGACY qui lit _STRUCTURE,
    //    crée le mapping source→dest, charge quotas/effectifs/parité/autorisations
    logLine('INFO', '🔧 Construction du contexte LEGACY complet depuis _STRUCTURE...');
    const ctx = makeCtxFromSourceSheets_LEGACY();
    
    // ✅ Charger les élèves depuis les onglets sources
    logLine('INFO', '📚 Chargement des élèves depuis les onglets sources...');
    const students = loadAllStudentsData(ctx);
    ctx.allStudents = students;
    
    if (!ctx.allStudents || ctx.allStudents.length === 0) {
      logLine('ERROR', '❌ Aucun élève chargé depuis les onglets sources');
      ui.alert('⚠️ Aucun élève trouvé dans les classes sources.\nVérifiez que les onglets sources contiennent des données.');
      return { success: false, error: 'No students' };
    }
    
    logLine('INFO', `✅ Contexte créé: ${ctx.allStudents.length} élèves`);
    logLine('INFO', `📋 Onglets sources: ${(ctx.srcSheets || []).join(', ')}`);
    logLine('INFO', `📋 Onglets TEST cibles: ${(ctx.cacheSheets || []).join(', ')}`);

    // 3. INITIALISER ONGLETS TEST (avec mapping et en-têtes corrects)
    logLine('INFO', '📋 Initialisation des onglets TEST...');
    initEmptyTestTabs_LEGACY(ctx);
    logLine('INFO', `✅ Onglets TEST initialisés: ${ctx.cacheSheets.length}`);

    // 4. PHASE 1 : Répartition OPTIONS/LV2 selon quotas
    logLine('INFO', '\n📌 PHASE 1: Répartition OPTIONS/LV2...');
    const p1Result = Phase1I_dispatchOptionsLV2_LEGACY(ctx);
    if (!p1Result.ok) {
      logLine('ERROR', `❌ Erreur Phase 1: ${p1Result.error || 'Échec'}`);
      ui.alert(`❌ Erreur Phase 1: ${p1Result.error || 'Échec répartition OPTIONS/LV2'}`);
      return { success: false, error: 'Phase 1 failed' };
    }
    logLine('SUCCESS', `✅ Phase 1 terminée: ${p1Result.placed || 0} élèves placés avec OPTIONS/LV2`);

    // 5. PHASE 2 : Codes ASSO/DISSO (D1, fratries, etc.)
    logLine('INFO', '\n📌 PHASE 2: Application codes ASSO/DISSO...');
    const p2Result = Phase2I_applyDissoAsso_LEGACY(ctx);
    if (!p2Result.ok) {
      logLine('ERROR', `❌ Erreur Phase 2: ${p2Result.error || 'Échec'}`);
      ui.alert(`❌ Erreur Phase 2: ${p2Result.error || 'Échec codes ASSO/DISSO'}`);
      return { success: false, error: 'Phase 2 failed' };
    }
    logLine('SUCCESS', `✅ Phase 2 terminée: ASSO=${p2Result.asso || 0}, DISSO=${p2Result.disso || 0}`);

    // 6. PHASE 3 : Compléter effectifs et équilibrer parité
    logLine('INFO', '\n📌 PHASE 3: Effectifs & Parité...');
    const p3Result = Phase3I_completeAndParity_LEGACY(ctx);
    if (!p3Result.ok) {
      logLine('ERROR', `❌ Erreur Phase 3: ${p3Result.error || 'Échec'}`);
      ui.alert(`❌ Erreur Phase 3: ${p3Result.error || 'Échec parité'}`);
      return { success: false, error: 'Phase 3 failed' };
    }
    logLine('SUCCESS', `✅ Phase 3 terminée: ${p3Result.placed || 0} élèves placés, parité équilibrée`);

    // 7. CROSS-PHASE LOOP : Phase 3 → Phase 4 avec feedback
    const crossPhaseLoops = MULTI_RESTART_CONFIG.crossPhaseLoops;
    let p4Result = null;
    let prevSwaps = 0;

    for (let cpLoop = 0; cpLoop <= crossPhaseLoops; cpLoop++) {
      // ✅ FIX #5 : Garde-fou maxRuntime — arrêter avant le timeout GAS (6 min)
      const elapsedSec = (new Date() - startTime) / 1000;
      const maxRuntimeSec = LEGACY_PIPELINE_CONFIG.maxRuntime || 600;
      if (elapsedSec > maxRuntimeSec * 0.85) { // 85% du budget → arrêt préventif
        logLine('WARN', `⏱️ Garde-fou runtime : ${elapsedSec.toFixed(0)}s écoulées (limite: ${maxRuntimeSec}s). Arrêt cross-phase.`);
        break;
      }

      if (cpLoop > 0) {
        logLine('INFO', '\n🔄 CROSS-PHASE boucle ' + cpLoop + '/' + crossPhaseLoops + ' : relance Phase 3 + Phase 4');

        // Re-run Phase 3 pour redistribuer
        const p3b = Phase3I_completeAndParity_LEGACY(ctx);
        logLine('INFO', '  Phase 3 cross-phase : ' + (p3b.placed || 0) + ' élèves replacés');
      }

      logLine('INFO', '\n⚡ PHASE 4: Optimisation ULTIMATE' + (cpLoop > 0 ? ' (cross-phase #' + cpLoop + ')' : '') + '...');
      logLine('INFO', '🔧 [engine=LEGACY] Sélection moteur: Phase4_Ultimate_Run');
      p4Result = Phase4_Ultimate_Run(ctx);

      if (!p4Result.ok) {
        logLine('ERROR', `❌ Erreur moteur: ${p4Result.message}`);
        ui.alert(`❌ Erreur optimisation: ${p4Result.message}`);
        return { success: false, error: p4Result.message };
      }

      logLine('SUCCESS', `✅ Swaps appliqués: ${p4Result.swapsApplied}`);

      // Vérifier si l'amélioration est suffisante pour continuer
      if (cpLoop > 0 && p4Result.swapsApplied <= prevSwaps * 0.1) {
        logLine('INFO', '  🛑 Peu d\'amélioration supplémentaire, arrêt cross-phase.');
        break;
      }
      prevSwaps = p4Result.swapsApplied;
    }

    // 8. CRÉER ONGLETS FIN avec contexte complet
    logLine('INFO', '\n💾 Finalisation avec contexte...');
    const finResult = finalizeAllSheets(ctx);
    logLine('SUCCESS', `✅ Onglets FIN créés: ${finResult.count}`);

    // 9. RÉSUMÉ
    const runtime = (new Date() - startTime) / 1000;
    logLine('SUCCESS', `\n✅ PIPELINE LEGACY TERMINÉ (${runtime.toFixed(1)}s)`);
    logLine('INFO', '═'.repeat(80));

    /*
    ui.alert(
      `✅ RÉPARTITION TERMINÉE\n\n` +
      `• Élèves: ${ctx.allStudents.length}\n` +
      `• Classes: ${ctx.srcSheets.length}\n` +
      `• Optimisations: ${p4Result.swapsApplied}\n` +
      `• Durée: ${runtime.toFixed(1)}s\n\n` +
      `Onglets FIN prêts à utiliser !`
    );
    */

    return {
      success: true,
      students: ctx.allStudents.length,
      classes: ctx.srcSheets.length,
      swaps: p4Result.swapsApplied,
      runtime: runtime,
      timestamp: new Date().toISOString()
    };

  } catch (e) {
    logLine('ERROR', `❌ Erreur pipeline: ${e.toString()}`);
    ui.alert(`❌ Erreur: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}

// ===================================================================
// UTILITAIRES LEGACY
// ===================================================================

/**
 * Détecte les onglets sources (format: 6°1, 5°2, ECOLE°1, etc.)
 */
function detectSourceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => /.+°\d+$/.test(name)) // ✅ Règle stricte °Chiffre
    .sort();
}

/**
 * Crée le contexte LEGACY
 */
function buildLegacyContext(sourceSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ctx = {
    ss: ss,
    allStudents: [],
    byClass: {},
    cacheSheets: sourceSheets,
    timestamp: new Date().getTime()
  };

  // Charger les élèves
  const students = loadAllStudentsData(ctx);
  ctx.allStudents = students;

  // Grouper par classe source
  sourceSheets.forEach(className => {
    ctx.byClass[className] = [];
  });

  return ctx;
}

/**
 * Crée les onglets TEST (vides initialement)
 */
function createTestSheets(ctx) {
  const ss = ctx.ss;

  ctx.cacheSheets.forEach(sourceSheet => {
    const testName = sourceSheet + 'TEST';
    let testSheet = ss.getSheetByName(testName);

    if (!testSheet) {
      testSheet = ss.insertSheet(testName);
      logLine('INFO', `  ✅ Onglet créé: ${testName}`);
    } else {
      testSheet.clearContents();
      logLine('INFO', `  ♻️ Onglet réutilisé: ${testName}`);
    }
  });

  SpreadsheetApp.flush();
}

/**
 * Crée les onglets FIN définitifs avec formatage
 * ✅ CORRECTION : Utiliser le contexte pour copier TEST→FIN avec formatage
 */
function finalizeAllSheets(ctx) {
  try {
    const ss = ctx.ss;
    const createdSheets = [];
    
    // Pour chaque onglet TEST, créer un onglet FIN
    (ctx.cacheSheets || []).forEach(testName => {
      const finName = testName.replace(/TEST$/i, 'FIN');
      const testSheet = ss.getSheetByName(testName);
      
      if (!testSheet) {
        logLine('WARN', `⚠️ Onglet ${testName} introuvable pour finalisation`);
        return;
      }
      
      // Supprimer l'ancien FIN si existe
      let finSheet = ss.getSheetByName(finName);
      if (finSheet) {
        ss.deleteSheet(finSheet);
      }
      
      // Copier TEST → FIN
      finSheet = testSheet.copyTo(ss);
      finSheet.setName(finName);
      
      // ✅ APPLIQUER LA MISE EN FORME
      formatFinSheet_LEGACY(finSheet);
      
      logLine('INFO', `  ✅ ${finName} créé depuis ${testName}`);
      createdSheets.push(finName);
    });
    
    SpreadsheetApp.flush();
    
    return {
      ok: true,
      count: createdSheets.length,
      created: createdSheets
    };
    
  } catch (e) {
    logLine('ERROR', `❌ Erreur finalisation: ${e.message}`);
    return {
      ok: false,
      count: 0,
      created: [],
      error: e.message
    };
  }
}

/**
 * Applique la mise en forme EXACTE de l'Interface V2 sur un onglet FIN
 * Copie fidèle du style avec couleurs cellule par cellule
 * @param {Sheet} sheet - L'onglet FIN à formater
 */
function formatFinSheet_LEGACY(sheet) {
  try {
    if (!sheet || sheet.getLastRow() <= 1) {
      return; // Pas de données
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rowData = data.slice(1);
    
    // ========== COULEURS NOUVELLES (LV2 = COULEURS PAYS, OPT = COULEURS DISTINCTES) ==========
    const COLORS = {
      // En-tête
      header: '#2c3e50',      // Gris foncé
      headerText: '#ffffff',  // Blanc
      
      // SEXE (couleurs personnalisées)
      sexeF: '#f5b7b1',       // Rose personnalisé
      sexeM: '#85c1e9',       // Bleu personnalisé
      
      // LV2 (Couleurs des pays)
      lv2ESP: '#FFB347',      // Orange (Espagne)
      lv2ITA: '#d5f5e3',      // Vert personnalisé (Italie)
      lv2ALL: '#FFED4E',      // Jaune (Allemagne)
      lv2PT: '#32CD32',       // Vert (Portugal)
      lv2OR: '#FFD700',       // Or
      lv2Default: '#ffffff',  // Blanc
      
      // OPT (Couleurs distinctes avec meilleur contraste)
      optCHAV: '#8B4789',     // Violet plus foncé (CHAV) - meilleur contraste
      optLATIN: '#e8f8f5',    // Vert d'eau (LATIN)
      optCHINOIS: '#C41E3A',  // Rouge cardinal (CHINOIS)
      optGREC: '#f6ca9d',     // Orange clair (GREC)
      
      // COM/TRA/PART/ABS (notes)
      note4: '#38761d',       // Vert TRÈS foncé
      note3: '#8ec875',       // Vert personnalisé
      note2: '#f1c232',       // Jaune-orange vif
      note1: '#cc0000',       // Rouge vif
      noteHighText: '#ffffff' // Texte blanc pour 4 et 1
    };
    
    // ========== 0. CACHER COLONNES A, B ET C ==========
    sheet.hideColumns(1, 3); // Cache les 3 premières colonnes
    
    // ========== 1. EN-TÊTE ==========
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(COLORS.header);
    headerRange.setFontColor(COLORS.headerText);
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(12); // En-tête plus grand
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 30);
    
    // ========== TOUT EN GRAS + TAILLE POLICE 11 ==========
    const allDataRange = sheet.getRange(2, 1, rowData.length, headers.length);
    allDataRange.setFontWeight('bold');
    allDataRange.setFontSize(11);
    
    // ========== 2. LARGEURS COLONNES + FORMATAGE NOM/PRENOM ==========
    for (let col = 1; col <= headers.length; col++) {
      const headerName = String(headers[col - 1]).toUpperCase();
      let width = 100;
      
      if (headerName === 'NOM' || headerName.includes('NOM')) {
        width = 180; // Plus large pour NOM
        // Mettre NOM en gras pour toutes les lignes
        const nomCol = sheet.getRange(2, col, rowData.length, 1);
        nomCol.setFontWeight('bold');
      }
      else if (headerName === 'PRENOM') {
        width = 150; // Plus large pour PRENOM
        // Mettre PRENOM en gras pour toutes les lignes
        const prenomCol = sheet.getRange(2, col, rowData.length, 1);
        prenomCol.setFontWeight('bold');
      }
      else if (headerName === 'SEXE') width = 50;
      else if (headerName === 'LV2') width = 60;
      else if (headerName === 'OPT') width = 70;
      else if (['COM', 'TRA', 'PART', 'ABS'].indexOf(headerName) >= 0) width = 50;
      else if (headerName === 'ID_ELEVE') width = 100;
      
      sheet.setColumnWidth(col, width);
    }
    
    // ========== 3. INDEX COLONNES ==========
    const idx = {
      SEXE: headers.indexOf('SEXE'),
      LV2: headers.indexOf('LV2'),
      OPT: headers.indexOf('OPT'),
      COM: headers.indexOf('COM'),
      TRA: headers.indexOf('TRA'),
      PART: headers.indexOf('PART'),
      ABS: headers.indexOf('ABS'),
      ID_ELEVE: headers.indexOf('ID_ELEVE')
    };
    
    // ========== 4. FORMATAGE CELLULE PAR CELLULE ==========
    for (let i = 0; i < rowData.length; i++) {
      const rowNum = i + 2; // +2 (ligne 1 = header, ligne 2 = premier élève)
      const row = rowData[i];
      
      // Alternance blanc/gris clair (lignes impaires = gris)
      const bgDefault = (i % 2 === 0) ? '#ffffff' : '#f3f3f3';
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bgDefault).setFontColor('#000000');
      
      // SEXE (couleurs personnalisées avec texte noir pour meilleur contraste)
      if (idx.SEXE >= 0) {
        const sexe = String(row[idx.SEXE] || '').trim().toUpperCase();
        const cellSexe = sheet.getRange(rowNum, idx.SEXE + 1);
        if (sexe === 'F') {
          cellSexe.setBackground(COLORS.sexeF).setFontColor('#000000');
        } else if (sexe === 'M') {
          cellSexe.setBackground(COLORS.sexeM).setFontColor('#000000');
        }
        cellSexe.setHorizontalAlignment('center').setFontWeight('bold');
      }
      
      // LV2 (Couleurs des pays)
      if (idx.LV2 >= 0) {
        const lv2 = String(row[idx.LV2] || '').trim().toUpperCase();
        const cellLV2 = sheet.getRange(rowNum, idx.LV2 + 1);
        if (lv2 === 'ESP') cellLV2.setBackground(COLORS.lv2ESP);
        else if (lv2 === 'ITA') cellLV2.setBackground(COLORS.lv2ITA);
        else if (lv2 === 'ALL') cellLV2.setBackground(COLORS.lv2ALL);
        else if (lv2 === 'PT') cellLV2.setBackground(COLORS.lv2PT);
        else if (lv2 === 'OR') cellLV2.setBackground(COLORS.lv2OR);
        cellLV2.setHorizontalAlignment('center').setFontWeight('bold');
      }
      
      // OPT (Couleurs distinctes par option avec contraste optimal)
      if (idx.OPT >= 0) {
        const opt = String(row[idx.OPT] || '').trim().toUpperCase();
        const cellOPT = sheet.getRange(rowNum, idx.OPT + 1);
        if (opt === 'CHAV') {
          // Violet foncé → texte blanc gras
          cellOPT.setBackground(COLORS.optCHAV).setFontColor('#ffffff').setFontWeight('bold');
        } else if (opt === 'LATIN') {
          // Vert d'eau → texte noir gras
          cellOPT.setBackground(COLORS.optLATIN).setFontColor('#000000').setFontWeight('bold');
        } else if (opt === 'CHINOIS') {
          // Rouge foncé → texte blanc gras
          cellOPT.setBackground(COLORS.optCHINOIS).setFontColor('#ffffff').setFontWeight('bold');
        } else if (opt === 'GREC') {
          // Orange clair → texte noir gras
          cellOPT.setBackground(COLORS.optGREC).setFontColor('#000000').setFontWeight('bold');
        }
        cellOPT.setHorizontalAlignment('center');
      }
      
      // COM, TRA, PART, ABS (notes)
      ['COM', 'TRA', 'PART', 'ABS'].forEach(col => {
        if (idx[col] >= 0) {
          const val = Number(row[idx[col]]) || 0;
          const cell = sheet.getRange(rowNum, idx[col] + 1);
          
          if (val >= 4) {
            cell.setBackground(COLORS.note4).setFontColor(COLORS.noteHighText);
          } else if (val >= 3) {
            cell.setBackground(COLORS.note3).setFontColor('#000000');
          } else if (val >= 2) {
            cell.setBackground(COLORS.note2).setFontColor('#000000');
          } else if (val >= 1) {
            cell.setBackground(COLORS.note1).setFontColor(COLORS.noteHighText);
          }
          
          cell.setHorizontalAlignment('center').setFontWeight('bold');
        }
      });
    }
    
    // ========== 5. PAS DE GRILLAGE (bordures enlevées) ==========
    // Alternance blanc/gris suffit, pas de bordures
    
    // ========== 6. CACHER COLONNES (garder seulement D-N et R visibles) ==========
    // Colonnes visibles : D(4)=NOM&PRENOM, E(5)=SEXE, F(6)=LV2, G(7)=OPT, 
    //                     H(8)=COM, I(9)=TRA, J(10)=PART, K(11)=ABS,
    //                     L(12)=DISPO, M(13)=ASSO, N(14)=DISSO, R(18)=CLASSE DEF
    const visibleColumns = ['NOM', 'PRENOM', 'NOM & PRENOM', 'SEXE', 'LV2', 'OPT', 
                            'COM', 'TRA', 'PART', 'ABS', 'DISPO', 'ASSO', 'DISSO', 
                            'CLASSE DEF', '_CLASS_ASSIGNED'];
    
    for (let col = 1; col <= headers.length; col++) {
      const headerName = String(headers[col - 1]).toUpperCase();
      const isVisible = visibleColumns.some(v => headerName.includes(v.toUpperCase()));
      
      if (!isVisible) {
        try {
          sheet.hideColumns(col);
        } catch (e) {
          // Erreur lors du masquage
        }
      }
    }
    
    // ========== 7. STATISTIQUES EN BAS ==========
    addStatistics_LEGACY_V2(sheet, headers, rowData, idx);
    
    SpreadsheetApp.flush();
    logLine('INFO', `    🎨 Mise en forme Interface V2 appliquée à ${sheet.getName()}`);
    
  } catch (e) {
    logLine('WARN', `    ⚠️ Erreur formatage ${sheet.getName()}: ${e.message}`);
  }
}

/**
 * Ajoute des statistiques en bas de l'onglet FIN (Style Interface V2)
 * @param {Sheet} sheet - L'onglet FIN
 * @param {Array} headers - En-têtes
 * @param {Array} rowData - Données élèves
 * @param {Object} idx - Index des colonnes
 */
function addStatistics_LEGACY_V2(sheet, headers, rowData, idx) {
  try {
    const statsRow = rowData.length + 3; // +3 pour séparer des données
    
    // Couleurs Interface V2 (EXACTES - vives)
    const COLORS = {
      sexeF: '#f4cccc',
      sexeM: '#cfe2f3',
      lv2ESP: '#ffd966',
      lv2ITA: '#9fc5e8',
      note4: '#38761d',
      note3: '#6aa84f',
      note2: '#f1c232',
      note1: '#cc0000'
    };
    
    // ========== LIGNE 1 : COMPTAGES COLORÉS ==========
    // Comptages par colonne, alignés avec les en-têtes
    
    // SEXE (Filles / Garçons)
    if (idx.SEXE >= 0) {
      const countF = rowData.filter(r => String(r[idx.SEXE]).toUpperCase() === 'F').length;
      const countM = rowData.filter(r => String(r[idx.SEXE]).toUpperCase() === 'M').length;
      
      sheet.getRange(statsRow, idx.SEXE + 1).setValue(countF)
        .setBackground(COLORS.sexeF).setFontWeight('bold').setHorizontalAlignment('center');
      sheet.getRange(statsRow + 1, idx.SEXE + 1).setValue(countM)
        .setBackground(COLORS.sexeM).setFontWeight('bold').setHorizontalAlignment('center');
    }
    
    // LV2 (ESP / ITA)
    if (idx.LV2 >= 0) {
      const countESP = rowData.filter(r => String(r[idx.LV2]).toUpperCase() === 'ESP').length;
      const countITA = rowData.filter(r => String(r[idx.LV2]).toUpperCase() === 'ITA').length;
      
      if (countESP > 0) {
        sheet.getRange(statsRow, idx.LV2 + 1).setValue(countESP)
          .setBackground(COLORS.lv2ESP).setFontWeight('bold').setHorizontalAlignment('center');
      }
      if (countITA > 0) {
        sheet.getRange(statsRow + 1, idx.LV2 + 1).setValue(countITA)
          .setBackground(COLORS.lv2ITA).setFontWeight('bold').setHorizontalAlignment('center');
      }
    }
    
    // OPT (LATIN / CHAV)
    if (idx.OPT >= 0) {
      const countLATIN = rowData.filter(r => String(r[idx.OPT]).toUpperCase() === 'LATIN').length;
      const countCHAV = rowData.filter(r => String(r[idx.OPT]).toUpperCase() === 'CHAV').length;
      
      if (countLATIN > 0) {
        sheet.getRange(statsRow, idx.OPT + 1).setValue(countLATIN)
          .setBackground('#a64d79').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      }
    }
    
    // COM / TRA / PART / ABS : Comptages par note (4, 3, 2, 1)
    ['COM', 'TRA', 'PART', 'ABS'].forEach(col => {
      if (idx[col] >= 0) {
        const count4 = rowData.filter(r => Number(r[idx[col]]) === 4).length;
        const count3 = rowData.filter(r => Number(r[idx[col]]) === 3).length;
        const count2 = rowData.filter(r => Number(r[idx[col]]) === 2).length;
        const count1 = rowData.filter(r => Number(r[idx[col]]) === 1).length;
        
        if (count4 > 0) {
          sheet.getRange(statsRow, idx[col] + 1).setValue(count4)
            .setBackground(COLORS.note4).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
        }
        if (count3 > 0) {
          sheet.getRange(statsRow + 1, idx[col] + 1).setValue(count3)
            .setBackground(COLORS.note3).setFontWeight('bold').setHorizontalAlignment('center');
        }
        if (count2 > 0) {
          sheet.getRange(statsRow + 2, idx[col] + 1).setValue(count2)
            .setBackground(COLORS.note2).setFontWeight('bold').setHorizontalAlignment('center');
        }
        if (count1 > 0) {
          sheet.getRange(statsRow + 3, idx[col] + 1).setValue(count1)
            .setBackground(COLORS.note1).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
        }
      }
    });
    
    // ========== LIGNE SUIVANTE : MOYENNES ==========
    const avgRow = statsRow + 4;
    
    ['COM', 'TRA', 'PART', 'ABS'].forEach(col => {
      if (idx[col] >= 0) {
        const avg = rowData.reduce((sum, r) => sum + (Number(r[idx[col]]) || 0), 0) / rowData.length;
        sheet.getRange(avgRow, idx[col] + 1).setValue(avg.toFixed(2))
          .setFontWeight('bold').setHorizontalAlignment('center');
      }
    });
    
  } catch (e) {
    logLine('WARN', `    ⚠️ Erreur ajout statistiques: ${e.message}`);
  }
}

// logLine() defined in Phase4_Ultimate.gs (single global definition)

// ===================================================================
// ENTRÉES ALTERNATIVES (Menu + Console)
// ===================================================================

/**
 * Entrée depuis Console V3 (Phase 4 button)
 */
function ouvrirPipeline_FromConsole_V3(options) {
  logLine('INFO', '📋 Appel depuis Console V3');
  return legacy_runFullPipeline_PRIME();
}

// legacy_viewSourceClasses() moved to Code.gs (single entry point)

// ===================================================================
// TEST FUNCTION
// ===================================================================

/**
 * Test du pipeline (debug)
 */
function testLEGACY_Pipeline() {
  logLine('INFO', '🧪 TEST PIPELINE LEGACY...');
  const result = legacy_runFullPipeline_PRIME();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
