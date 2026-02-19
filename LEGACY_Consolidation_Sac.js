/**
 * ===================================================================
 * 🎯 ARCHITECTURE "SAC DE BILLES" - Consolidation SOURCE + TEST
 * ===================================================================
 * 
 * Phase 1 : Sort du "sac" (SOURCE) uniquement les élèves avec contraintes
 * Phase 2 & 3 : Fusionnent TEST (déjà placés) + SOURCE (encore dans le sac)
 * 
 * Date : 2025-11-22
 * ===================================================================
 */

/**
 * Fusionne les données TEST (élèves déjà placés) + CONSOLIDATION (élèves dans le sac)
 * @param {Object} ctx - Contexte LEGACY
 * @returns {Object} { allData: Array, headersRef: Array }
 */
function getConsolidatedData_LEGACY(ctx) {
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const allData = [];
  let headersRef = null;
  
  logLine('INFO', '🔄 Consolidation SAC DE BILLES (TEST + CONSOLIDATION)...');
  
  // ========== ÉTAPE 1 : LIRE LES ÉLÈVES DÉJÀ PLACÉS (TEST) ==========
  let countTest = 0;
  const idsPlaces = new Set(); // IDs des élèves déjà placés
  
  (ctx.cacheSheets || []).forEach(function(testName) {
    const testSheet = ss.getSheetByName(testName);
    if (!testSheet || testSheet.getLastRow() <= 1) return;

    const data = testSheet.getDataRange().getValues();
    if (!headersRef) headersRef = data[0];
    
    const idxID = headersRef.indexOf('ID_ELEVE');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const idEleve = String(row[idxID] || '').trim();
      
      if (idEleve) {
        idsPlaces.add(idEleve); // Marquer comme placé
      }
      
      allData.push({
        sheetName: testName,
        row: row,
        source: 'TEST' // Marqueur pour debug
      });
      countTest++;
    }
  });
  
  logLine('INFO', '  ✅ ' + countTest + ' élèves lus depuis TEST (déjà placés)');
  
  // ========== ÉTAPE 2 : LIRE LES ÉLÈVES ENCORE DANS LE SAC (CONSOLIDATION) ==========
  const consolidationSheet = ss.getSheetByName('CONSOLIDATION');
  let countSource = 0;
  
  if (!consolidationSheet || consolidationSheet.getLastRow() <= 1) {
    logLine('WARN', '⚠️ CONSOLIDATION vide ou introuvable, aucun élève dans le sac');
  } else {
    const data = consolidationSheet.getDataRange().getValues();
    const headers = data[0];
    
    // ✅ Trouver et EXCLURE la colonne _ELEVE_PLACE (colonne de tracking temporaire)
    const idxElevePlace = headers.indexOf('_ELEVE_PLACE');
    const cleanHeaders = [];
    for (let h = 0; h < headers.length; h++) {
      if (h !== idxElevePlace) {
        cleanHeaders.push(headers[h]);
      }
    }
    
    // Vérifier que les headers sont compatibles
    if (!headersRef) {
      headersRef = cleanHeaders;
    }
    
    // Index des colonnes
    const idxID = headers.indexOf('ID_ELEVE');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const idEleve = String(row[idxID] || '').trim();
      
      if (!idEleve) continue; // Ligne vide
      
      // Si pas encore placé, l'ajouter au sac
      if (!idsPlaces.has(idEleve)) {
        // ✅ Nettoyer la ligne (exclure _ELEVE_PLACE)
        const cleanRow = [];
        for (let c = 0; c < row.length; c++) {
          if (c !== idxElevePlace) {
            cleanRow.push(row[c]);
          }
        }
        
        // Ajouter colonnes FIXE, MOBILITE, _CLASS_ASSIGNED vides
        const newRow = cleanRow.concat(['', '', '']); // P, Q, R vides
        allData.push({
          sheetName: 'CONSOLIDATION',
          row: newRow,
          source: 'SAC' // Marqueur pour debug
        });
        countSource++;
      }
    }
  }
  
  logLine('INFO', '  ✅ ' + countSource + ' élèves lus depuis CONSOLIDATION (encore dans le sac)');
  logLine('INFO', '  📊 TOTAL CONSOLIDÉ : ' + allData.length + ' élèves');
  
  // Vérifier qu'on a bien des en-têtes
  if (!headersRef) {
    logLine('ERROR', '❌ Aucun en-tête trouvé !');
    return { allData: [], headersRef: [] };
  }
  
  // Vérifier que les colonnes FIXE, MOBILITE, _CLASS_ASSIGNED existent
  const idxFIXE = headersRef.indexOf('FIXE');
  const idxMOBILITE = headersRef.indexOf('MOBILITE');
  const idxAssigned = headersRef.indexOf('_CLASS_ASSIGNED');
  
  if (idxFIXE === -1 || idxMOBILITE === -1 || idxAssigned === -1) {
    logLine('WARN', '⚠️ Colonnes FIXE/MOBILITE/_CLASS_ASSIGNED manquantes, ajout aux headers');
    headersRef = headersRef.concat(['FIXE', 'MOBILITE', '_CLASS_ASSIGNED']);
  }
  
  return {
    allData: allData,
    headersRef: headersRef
  };
}
