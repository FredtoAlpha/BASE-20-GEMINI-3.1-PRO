/**
 * ===================================================================
 * APP.CACHEMANAGER.JS - GESTION DU CACHE
 * ===================================================================
 *
 * Module contenant les fonctions de gestion du cache (onglets CACHE).
 * Responsabilités: lecture, écriture, activation, synchronisation.
 *
 * ARCHITECTURE PHASE 5 - Refactoring progressif
 * Extraction depuis Orchestration_V14I.js
 *
 * Date: 26 novembre 2025
 * Version: 1.1.0 — SAFE: suppression des stubs incomplets en collision avec Orchestration_V14I.js
 *
 * NOTE: computeMobilityFlags_() et auditCacheAgainstStructure_() sont définis
 * dans Orchestration_V14I.js (versions complètes avec résolution groupe A,
 * filtrage D-index, et détection CONFLIT). Les versions ici étaient des
 * simplifications naïves qui produisaient des résultats incorrects.
 * ===================================================================
 */

// ===================================================================
// LECTURE DEPUIS LE CACHE
// ===================================================================

/**
 * Lit les élèves depuis les onglets CACHE
 *
 * @param {Object} ctx - Contexte (avec cacheSheets, niveaux, ss)
 * @returns {Object} État des classes {niveau: [eleves]}
 */
function readElevesFromCache_(ctx) {
  const classesState = {};

  // ✅ FIX: Construire mapping onglet → niveau depuis ctx
  const sheetToNiveau = {};
  for (let i = 0; i < (ctx.cacheSheets || []).length; i++) {
    sheetToNiveau[ctx.cacheSheets[i]] = ctx.niveaux[i];
  }

  for (const sheetName of ctx.cacheSheets) {
    const sheet = ctx.ss.getSheetByName(sheetName);
    if (!sheet) {
      logLine('WARN', 'Feuille CACHE ' + sheetName + ' introuvable');
      continue;
    }

    const eleves = readElevesFromSheet_(sheet);
    // ✅ FIX: Utiliser le mapping au lieu du regex /CACHE$/
    const niveau = sheetToNiveau[sheetName] || sheetName.replace(/CACHE$/, '');
    classesState[niveau] = eleves;
  }

  return classesState;
}

/**
 * Lit les élèves depuis le mode source sélectionné (TEST/FIN/CACHE/...)
 *
 * @param {Object} ctx - Contexte (avec srcSheets, ss, modeSrc)
 * @returns {Object} État des classes {niveau: [eleves]}
 */
function readElevesFromSelectedMode_(ctx) {
  const classesState = {};

  // 🔒 GARDE-FOU : garantir un tableau exploitable
  let srcList = ctx && Array.isArray(ctx.srcSheets) ? ctx.srcSheets : null;
  if (!srcList) {
    if (ctx && typeof ctx.srcSheets === 'string') {
      srcList = ctx.srcSheets.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    } else {
      const tag = (ctx.mode || ctx.sourceTag || 'TEST').toString().trim();
      const lv  = Array.isArray(ctx.levels) ? ctx.levels : [];
      srcList = lv.map(function(l) { return l + tag; });
    }
    ctx.srcSheets = srcList;
  }

  for (const sheetName of srcList) {
    const sheet = ctx.ss.getSheetByName(sheetName);
    if (!sheet) {
      logLine('WARN', 'Feuille source ' + sheetName + ' introuvable');
      continue;
    }

    const eleves = readElevesFromSheet_(sheet);

    // ✅ FIX: Utiliser le mapping pour obtenir le nom de classe de destination
    let niveau;
    if (ctx.sourceToDestMapping && ctx.sourceToDestMapping[sheetName]) {
      niveau = ctx.sourceToDestMapping[sheetName];
      logLine('INFO', '  📌 Lecture ' + sheetName + ' → assignation à ' + niveau);
    } else {
      niveau = sheetName.replace(ctx.modeSrc || 'TEST', '');
    }

    classesState[niveau] = eleves;
  }

  return classesState;
}

// ===================================================================
// ACTIVATION ET SYNCHRONISATION DU CACHE
// ===================================================================

/**
 * Force le mode CACHE dans l'interface et recharge
 *
 * @param {Object} ctx - Contexte
 */
function forceCacheInUIAndReload_(ctx) {
  try {
    setInterfaceModeCACHE_(ctx);
    activateFirstCacheTabIfAny_(ctx);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Onglets CACHE mis à jour',
      'Optimisation V14I',
      3
    );

    triggerUIReloadFromCACHE_();

  } catch (e) {
    logLine('WARN', 'forceCacheInUIAndReload_ failed: ' + e.message);
  }
}

/**
 * Marque le mode CACHE comme actif dans l'interface
 *
 * @param {Object} ctx - Contexte
 */
function setInterfaceModeCACHE_(ctx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uiSheet = ss.getSheetByName('_INTERFACE_V2') || ss.getSheetByName('UI_Config');
  if (!uiSheet) return;

  try {
    uiSheet.getRange('B2').setValue('CACHE');
  } catch (e) {
    logLine('WARN', 'setInterfaceModeCACHE_ failed: ' + e.message);
  }
}

/**
 * Active visuellement le premier onglet CACHE
 *
 * @param {Object} ctx - Contexte (avec cacheSheets)
 */
function activateFirstCacheTabIfAny_(ctx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const firstName = ctx.cacheSheets && ctx.cacheSheets[0];
  if (firstName) {
    const sheet = ss.getSheetByName(firstName);
    if (sheet) ss.setActiveSheet(sheet);
  }
}

/**
 * Déclenche un reload côté front (HTML/JS)
 * Note: No-op côté Apps Script
 */
function triggerUIReloadFromCACHE_() {
  // Côté Apps Script : no-op
  // Côté front (HTML/JS) : ajouter un handler
}

/**
 * Ouvre visuellement les onglets CACHE et force la synchronisation
 *
 * @param {Object} ctx - Contexte (avec cacheSheets, ss)
 * @returns {Object} Résultat {opened: [], active: string, stats: []}
 */
function openCacheTabs_(ctx) {
  try {
    SpreadsheetApp.flush();
    Utilities.sleep(200);

    const opened = [];
    const stats = [];

    for (let i = 0; i < ctx.cacheSheets.length; i++) {
      const name = ctx.cacheSheets[i];
      const sh = ctx.ss.getSheetByName(name);

      if (sh) {
        ctx.ss.setActiveSheet(sh);
        sh.getRange('A1').activate();
        opened.push(name);

        const rows = sh.getLastRow();
        const cols = sh.getLastColumn();
        stats.push({ sheet: name, rows: rows, cols: cols });

        logLine('INFO', '  ✅ Activé: ' + name + ' (' + rows + ' lignes, ' + cols + ' colonnes)');
        Utilities.sleep(80);
      } else {
        logLine('ERROR', '  ❌ Onglet ' + name + ' introuvable !');
      }
    }

    SpreadsheetApp.flush();

    const active = ctx.ss.getActiveSheet() ? ctx.ss.getActiveSheet().getName() : '(aucun)';
    logLine('INFO', '✅ Onglet actif final: ' + active);
    logLine('INFO', '✅ ' + opened.length + ' onglets CACHE activés: ' + opened.join(', '));

    return {
      opened: opened,
      active: active,
      stats: stats
    };
  } catch (e) {
    logLine('WARN', '⚠️ openCacheTabs_ a échoué: ' + e.message);
    return {
      opened: [],
      active: null,
      stats: [],
      error: e.message
    };
  }
}
