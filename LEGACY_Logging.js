/**
 * ===================================================================
 * 📝 PRIME LEGACY - SYSTÈME DE LOGS CENTRALISÉS
 * ===================================================================
 *
 * Gestion centralisée des logs pour le pipeline LEGACY :
 * - Écriture dans la feuille _LOGS_LEGACY
 * - Affichage dans les logs Apps Script (console)
 * - Niveaux : INFO, WARN, ERROR, SUCCESS
 * - Timestamps automatiques
 * - Archivage automatique (limite de lignes)
 *
 * Date : 2025-11-14
 * Branche : claude/legacy-pipeline-renewal-01FK5TFnxx6JjwZ9bMkF5hqw
 *
 * ===================================================================
 */

// Configuration globale
var LEGACY_LOGS_CONFIG = {
  sheetName: '_LOGS_LEGACY',
  maxLines: 5000,           // Nombre max de lignes avant archivage
  archiveThreshold: 4500,   // Seuil pour déclencher l'archivage
  enableConsoleLog: true,   // Logger aussi dans la console Apps Script
  enableSheetLog: true      // Logger dans la feuille _LOGS_LEGACY
};

/**
 * Initialise la feuille de logs _LOGS_LEGACY si elle n'existe pas
 * @returns {Sheet} La feuille de logs
 */
function initLegacyLogsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName(LEGACY_LOGS_CONFIG.sheetName);

  if (!logsSheet) {
    // Créer la feuille
    logsSheet = ss.insertSheet(LEGACY_LOGS_CONFIG.sheetName);

    // Créer les en-têtes
    const headers = ['TIMESTAMP', 'NIVEAU', 'PHASE', 'MESSAGE'];
    logsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Formater les en-têtes
    logsSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4A86E8')  // Bleu
      .setFontColor('#FFFFFF');  // Texte blanc

    // Figer la première ligne
    logsSheet.setFrozenRows(1);

    // Ajuster les largeurs de colonnes
    logsSheet.setColumnWidth(1, 180);  // TIMESTAMP
    logsSheet.setColumnWidth(2, 80);   // NIVEAU
    logsSheet.setColumnWidth(3, 150);  // PHASE
    logsSheet.setColumnWidth(4, 600);  // MESSAGE

    Logger.log('✅ Feuille ' + LEGACY_LOGS_CONFIG.sheetName + ' créée');
  }

  return logsSheet;
}

/**
 * Écrit une ligne de log dans la feuille _LOGS_LEGACY
 *
 * @param {string} level - Niveau de log (INFO, WARN, ERROR, SUCCESS)
 * @param {string} message - Message à logger
 * @param {string} phase - Phase actuelle (optionnel)
 */
function logLegacy(level, message, phase) {
  // ========== CONSOLE LOG ==========
  if (LEGACY_LOGS_CONFIG.enableConsoleLog) {
    const prefix = '[LEGACY-' + level + ']';
    Logger.log(prefix + ' ' + message);
  }

  // ========== SHEET LOG ==========
  if (!LEGACY_LOGS_CONFIG.enableSheetLog) {
    return;
  }

  try {
    const logsSheet = initLegacyLogsSheet_();

    // Vérifier si archivage nécessaire
    if (logsSheet.getLastRow() > LEGACY_LOGS_CONFIG.archiveThreshold) {
      archiveLegacyLogs_(logsSheet);
    }

    // Préparer les données
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const phaseStr = phase || detectCurrentPhase_();

    // Ajouter la ligne
    const newRow = logsSheet.getLastRow() + 1;
    logsSheet.getRange(newRow, 1, 1, 4).setValues([[
      timestamp,
      level,
      phaseStr,
      message
    ]]);

    // Formater selon le niveau
    formatLogRow_(logsSheet, newRow, level);

  } catch (e) {
    // Fallback : log dans la console seulement
    Logger.log('[LEGACY-LOG-ERROR] Impossible d\'écrire dans ' + LEGACY_LOGS_CONFIG.sheetName + ' : ' + e.toString());
    Logger.log('[LEGACY-' + level + '] ' + message);
  }
}

/**
 * Formate une ligne de log selon son niveau
 * @param {Sheet} sheet - Feuille de logs
 * @param {number} row - Numéro de ligne
 * @param {string} level - Niveau de log
 */
function formatLogRow_(sheet, row, level) {
  const range = sheet.getRange(row, 1, 1, 4);

  switch (level) {
    case 'ERROR':
      range.setBackground('#F4C7C3');  // Rouge clair
      sheet.getRange(row, 2).setFontColor('#CC0000');  // Rouge foncé
      break;

    case 'WARN':
      range.setBackground('#FCE8B2');  // Jaune clair
      sheet.getRange(row, 2).setFontColor('#F57C00');  // Orange
      break;

    case 'SUCCESS':
      range.setBackground('#D9EAD3');  // Vert clair
      sheet.getRange(row, 2).setFontColor('#0F9D58');  // Vert
      break;

    case 'INFO':
    default:
      // Pas de formatage spécial
      break;
  }
}

/**
 * Détecte la phase actuelle basée sur la pile d'appels
 * @returns {string}
 */
function detectCurrentPhase_() {
  try {
    const stack = new Error().stack;
    if (stack.indexOf('Phase1') >= 0) return 'Phase 1';
    if (stack.indexOf('Phase2') >= 0) return 'Phase 2';
    if (stack.indexOf('Phase3') >= 0) return 'Phase 3';
    if (stack.indexOf('Phase4') >= 0) return 'Phase 4';
    if (stack.indexOf('Pipeline') >= 0) return 'Pipeline';
    if (stack.indexOf('Init') >= 0) return 'Init';
    if (stack.indexOf('Context') >= 0) return 'Context';
  } catch (e) {
    // Ignore
  }
  return 'General';
}

/**
 * Archive les anciens logs (garde seulement les N dernières lignes)
 * @param {Sheet} logsSheet - Feuille de logs
 */
function archiveLegacyLogs_(logsSheet) {
  const totalRows = logsSheet.getLastRow();

  if (totalRows <= LEGACY_LOGS_CONFIG.maxLines) {
    return;  // Pas besoin d'archiver
  }

  logLegacy('INFO', '📦 Archivage des logs anciens...');

  // Calculer combien de lignes à supprimer
  const rowsToDelete = totalRows - LEGACY_LOGS_CONFIG.maxLines;

  // Supprimer les lignes anciennes (après l'en-tête)
  logsSheet.deleteRows(2, rowsToDelete);

  logLegacy('SUCCESS', '✅ ' + rowsToDelete + ' lignes archivées');
}

/**
 * Efface tous les logs LEGACY
 */
function clearLegacyLogs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🗑️ Effacer les logs LEGACY',
    'Voulez-vous vraiment effacer tous les logs ?\n\n' +
    'Cette action est irréversible.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LEGACY_LOGS_CONFIG.sheetName);

    if (!logsSheet) {
      ui.alert('⚠️ Aucune feuille de logs trouvée');
      return;
    }

    // Garder seulement la ligne d'en-tête
    if (logsSheet.getLastRow() > 1) {
      logsSheet.deleteRows(2, logsSheet.getLastRow() - 1);
    }

    ui.alert('✅ Logs effacés', 'Tous les logs ont été supprimés.', ui.ButtonSet.OK);
    logLegacy('INFO', '🗑️ Logs effacés par l\'utilisateur');

  } catch (e) {
    ui.alert('❌ Erreur', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Affiche les derniers logs dans une alerte
 * @param {number} count - Nombre de logs à afficher (défaut: 20)
 */
function showRecentLegacyLogs(count) {
  const ui = SpreadsheetApp.getUi();
  const numLogs = count || 20;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LEGACY_LOGS_CONFIG.sheetName);

    if (!logsSheet || logsSheet.getLastRow() <= 1) {
      ui.alert('⚠️ Aucun log disponible', 'La feuille ' + LEGACY_LOGS_CONFIG.sheetName + ' est vide.', ui.ButtonSet.OK);
      return;
    }

    const totalRows = logsSheet.getLastRow();
    const startRow = Math.max(2, totalRows - numLogs + 1);
    const numRows = totalRows - startRow + 1;

    const data = logsSheet.getRange(startRow, 1, numRows, 4).getValues();

    // Construire le message
    let message = '';
    data.forEach(function(row) {
      const timestamp = row[0];
      const level = row[1];
      const phase = row[2];
      const msg = row[3];

      message += '[' + timestamp + '] ' + level + ' | ' + phase + '\n';
      message += '  ' + msg + '\n\n';
    });

    ui.alert(
      '📝 Derniers Logs LEGACY (' + numLogs + ')',
      message,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Erreur', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Ouvre la feuille de logs _LOGS_LEGACY
 */
function openLegacyLogsSheet() {
  try {
    const logsSheet = initLegacyLogsSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(logsSheet);

    // Scroller vers le bas (derniers logs)
    const lastRow = logsSheet.getLastRow();
    if (lastRow > 1) {
      logsSheet.setActiveRange(logsSheet.getRange(lastRow, 1));
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Erreur', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Exporte les logs vers un fichier texte (Drive)
 * @returns {string} URL du fichier créé
 */
function exportLegacyLogsToFile() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LEGACY_LOGS_CONFIG.sheetName);

    if (!logsSheet || logsSheet.getLastRow() <= 1) {
      ui.alert('⚠️ Aucun log à exporter');
      return null;
    }

    const data = logsSheet.getDataRange().getValues();

    // Construire le contenu du fichier
    let content = '='.repeat(80) + '\n';
    content += 'PRIME LEGACY - LOGS EXPORT\n';
    content += 'Date : ' + new Date().toISOString() + '\n';
    content += '='.repeat(80) + '\n\n';

    for (let i = 1; i < data.length; i++) {  // Skip header
      const row = data[i];
      content += '[' + row[0] + '] ' + row[1] + ' | ' + row[2] + '\n';
      content += row[3] + '\n';
      content += '-'.repeat(80) + '\n';
    }

    // Créer le fichier dans Drive
    const fileName = 'LEGACY_Logs_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.txt';
    const file = DriveApp.createFile(fileName, content, MimeType.PLAIN_TEXT);

    ui.alert(
      '✅ Export réussi',
      'Logs exportés vers :\n' + file.getName() + '\n\nURL : ' + file.getUrl(),
      ui.ButtonSet.OK
    );

    logLegacy('INFO', '📤 Logs exportés : ' + fileName);

    return file.getUrl();

  } catch (e) {
    ui.alert('❌ Erreur', e.toString(), ui.ButtonSet.OK);
    return null;
  }
}

/**
 * Statistiques des logs
 */
function getLegacyLogsStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LEGACY_LOGS_CONFIG.sheetName);

    if (!logsSheet || logsSheet.getLastRow() <= 1) {
      return {
        total: 0,
        INFO: 0,
        WARN: 0,
        ERROR: 0,
        SUCCESS: 0
      };
    }

    const data = logsSheet.getDataRange().getValues();
    const stats = {
      total: data.length - 1,
      INFO: 0,
      WARN: 0,
      ERROR: 0,
      SUCCESS: 0
    };

    for (let i = 1; i < data.length; i++) {
      const level = String(data[i][1] || '').trim().toUpperCase();
      if (stats[level] !== undefined) {
        stats[level]++;
      }
    }

    return stats;

  } catch (e) {
    Logger.log('Erreur getLegacyLogsStats : ' + e.toString());
    return { total: 0, INFO: 0, WARN: 0, ERROR: 0, SUCCESS: 0 };
  }
}
