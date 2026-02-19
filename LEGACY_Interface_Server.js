/**
 * ===================================================================
 *  PRIME LEGACY - Fonctions Serveur pour Interface Moderne
 * ===================================================================
 *
 * Ce fichier contient les fonctions Google Apps Script appelées par
 * l'interface utilisateur moderne (LEGACY_Interface.html).
 *
 * ===================================================================
 */

/**
 * Point d'entrée pour afficher l'interface moderne du pipeline LEGACY.
 * Appelé depuis le menu.
 */
function showLegacyInterface() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('LEGACY_Interface')
      .setWidth(800)
      .setHeight(700)
      .setTitle('PRIME LEGACY - Interface Moderne');
    SpreadsheetApp.getUi().showModalDialog(html, 'Lancement du Pipeline LEGACY');
  } catch (e) {
    Logger.log('Erreur showLegacyInterface: ' + e.toString());
    SpreadsheetApp.getUi().alert('Erreur: Impossible de charger l\'interface. ' + e.toString());
  }
}

/**
 * Récupère le contexte initial pour l'affichage dans l'UI.
 * Lit les classes sources et compte le nombre total d'élèves.
 * @returns {Object} Un objet avec les informations sur les sources.
 */
function legacy_getUiContext() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();

    // Logique de détection des sources (toujours avec °)
    const sourceSheets = allSheets.filter(function(s) {
      return /^[A-Za-z0-9_-]+°\d+$/.test(s.getName());
    });

    sourceSheets.sort(function(a, b) {
      return a.getName().localeCompare(b.getName());
    });

    let totalEleves = 0;
    const sheetDetails = sourceSheets.map(function(s) {
      const numEleves = Math.max(0, s.getLastRow() - 1);
      totalEleves += numEleves;
      return { name: s.getName(), count: numEleves };
    });

    return {
      success: true,
      sourceCount: sourceSheets.length,
      studentCount: totalEleves,
      sources: sheetDetails
    };

  } catch (e) {
    Logger.log('Erreur legacy_getUiContext: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Exécute une phase spécifique du pipeline LEGACY et retourne le résultat.
 * C'est une fonction wrapper pour rendre les phases compatibles avec l'UI.
 * @param {string} phase - Le nom de la phase à exécuter (ex: 'phase1').
 * @returns {Object} Un objet indiquant le succès et un message.
 */
function legacy_runPhase_fromUI(phase) {
  try {
    // Le contexte est reconstruit à chaque appel pour garantir l'état le plus récent
    const ctx = makeCtxFromSourceSheets_();

    let resultMessage = "Phase non reconnue.";

    switch(phase) {
      case 'phase1':
        Phase1I_dispatchOptionsLV2_(ctx);
        resultMessage = "Options & LV2 répartis.";
        break;
      case 'phase2':
        Phase2I_applyDissoAsso_(ctx);
        resultMessage = "ASSO/DISSO appliqués.";
        break;
      case 'phase3':
        Phase3I_completeAndParity_(ctx);
        resultMessage = "Effectifs & Parité équilibrés.";
        break;
      case 'phase4':
        Phase4_Ultimate_Run(ctx);
        resultMessage = "Équilibrage des scores terminé (ULTIMATE).";
        break;
      default:
        throw new Error("Phase inconnue: " + phase);
    }

    SpreadsheetApp.flush(); // S'assurer que les modifications sont écrites

    return { success: true, message: resultMessage };

  } catch (e) {
    Logger.log('Erreur legacy_runPhase_fromUI (' + phase + '): ' + e.toString() + e.stack);
    return { success: false, error: e.toString() };
  }
}
