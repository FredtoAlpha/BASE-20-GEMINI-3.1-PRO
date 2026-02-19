/**
 * ===================================================================
 * BACKEND CORE - Fonctions essentielles pour InterfaceV2
 * ===================================================================
 * Version: 1.1.0 — SAFE: suppression des 6 fonctions en collision avec Code.js
 *
 * Fonctions supprimées (définitions canoniques dans Code.js) :
 *  - getClassesData()           → Code.js (version avec SHEET_PATTERNS + per-sheet try/catch)
 *  - getLastCacheInfo()         → Code.js (lecture PropertiesService, pas Sheets)
 *  - getBridgeContextAndClear() → Code.js (safeGetUserProperty)
 *  - saveCacheData()            → Code.js (PropertiesService, pas Sheets)
 *  - loadCacheData()            → Code.js (cohérent avec saveCacheData)
 *  - saveElevesSnapshot()       → Backend_Eleves.js (avec saveStudentsToSheet helper)
 *
 * Fonctions conservées (uniques à ce fichier) :
 *  - chargerContraintes()
 *  - testBackendConnection()
 *
 * getUiSettings() → supprimée (collision avec Code.js, Code.js gagne)
 * ===================================================================
 */

/**
 * Charge les contraintes (ASSO/DISSO) depuis _STRUCTURE
 */
function chargerContraintes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const structureSheet = ss.getSheetByName('_STRUCTURE');

    if (!structureSheet) {
      return {
        success: false,
        error: 'Onglet _STRUCTURE introuvable'
      };
    }

    const data = structureSheet.getDataRange().getValues();

    return {
      success: true,
      constraints: {
        structure: data.slice(1)
      }
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Test de connexion backend
 */
function testBackendConnection() {
  return {
    success: true,
    message: 'Backend connecté !',
    timestamp: new Date().toISOString(),
    version: '1.0.0'
  };
}
