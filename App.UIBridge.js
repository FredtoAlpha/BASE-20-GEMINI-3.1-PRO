/**
 * ===================================================================
 * APP.UIBRIDGE.JS - BRIDGE INTERFACE UTILISATEUR
 * ===================================================================
 *
 * Module contenant les fonctions de communication avec l'interface utilisateur.
 * Responsabilités: annonces, toasts, notifications.
 *
 * ARCHITECTURE PHASE 5 - Refactoring progressif
 * Extraction depuis Orchestration_V14I.js
 *
 * Date: 26 novembre 2025
 * Version: 1.0.0
 * ===================================================================
 */

// ===================================================================
// ANNONCES ET NOTIFICATIONS
// ===================================================================

/**
 * Affiche une annonce de fin de phase
 *
 * @param {string} message - Message à afficher
 *
 * @example
 * announcePhaseDone_('Phase 1 terminée : 150 élèves répartis');
 */
function announcePhaseDone_(message) {
  logLine('INFO', '✅ ' + message);
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Phase terminée', 2);
}

// ===================================================================
// EXPORTS (Google Apps Script charge automatiquement)
// ===================================================================

/**
 * Note : Dans Google Apps Script, tous les fichiers .js sont chargés
 * automatiquement dans le scope global. Pas besoin d'export/import.
 *
 * Les fonctions définies ici sont automatiquement disponibles dans
 * tous les autres fichiers du projet.
 */
