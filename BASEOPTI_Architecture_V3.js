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
