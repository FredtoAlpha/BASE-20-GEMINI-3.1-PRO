/**
 * ===================================================================
 * APP.SHEETSDATA.JS - ACCÈS GOOGLE SHEETS
 * ===================================================================
 *
 * Module contenant toutes les fonctions d'accès aux données Google Sheets.
 * Responsabilités: lecture, écriture, gestion des onglets.
 *
 * ARCHITECTURE PHASE 5 - Refactoring progressif
 * Extraction depuis Orchestration_V14I.js
 *
 * Date: 26 novembre 2025
 * Version: 1.0.0
 * ===================================================================
 */

// ===================================================================
// GESTION DES ONGLETS (CRÉATION/ACCÈS)
// ===================================================================

/**
 * Obtient ou crée un onglet par nom
 *
 * @param {string} name - Nom de l'onglet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} L'onglet
 *
 * @example
 * const sheet = getOrCreateSheet_('6°1CACHE');
 */
function getOrCreateSheet_(name) {
  const ss = getActiveSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    logLine('INFO', '📄 Création onglet: ' + name);
    sh = ss.insertSheet(name);
  }
  return sh;
}

/**
 * Obtient ou crée un onglet par nom exact et le rend visible
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Le spreadsheet
 * @param {string} name - Nom exact de l'onglet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} L'onglet
 *
 * @example
 * const sheet = getOrCreateSheetByExactName_(ss, '6°1CACHE');
 */
function getOrCreateSheetByExactName_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  try {
    if (sheet.isSheetHidden && sheet.isSheetHidden()) {
      sheet.showSheet();
    }
  } catch (e) {
    // Ignorer si l'API ne supporte pas isSheetHidden
  }
  return sheet;
}

/**
 * Assure qu'une colonne existe dans un onglet (crée si absente)
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - L'onglet
 * @param {string} headerName - Nom de la colonne
 * @returns {number} Index de la colonne (0-based)
 *
 * @example
 * const idx = ensureColumn_(sheet, 'MOBILITE');
 */
function ensureColumn_(sheet, headerName) {
  const rng = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1);
  const headers = rng.getValues()[0];
  let idx = headers.indexOf(headerName);
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue(headerName);
    SpreadsheetApp.flush();
  }
  return idx; // 0-based
}

// ===================================================================
// LECTURE DE DONNÉES
// ===================================================================

/**
 * Lit les élèves depuis un onglet
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - L'onglet à lire
 * @returns {Array<Object>} Liste des élèves (objets avec clés = headers)
 *
 * @example
 * const eleves = readElevesFromSheet_(sheet);
 * // → [{Nom: 'Dupont', Prénom: 'Jean', ...}, ...]
 */
function readElevesFromSheet_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const eleves = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Ligne vide

    const eleve = {};
    headers.forEach((h, j) => {
      eleve[h] = row[j];
    });
    eleves.push(eleve);
  }

  return eleves;
}

/**
 * Lit les quotas OPTIONS/LV2 depuis _STRUCTURE
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - L'onglet _STRUCTURE
 * @returns {Object} Quotas {classe: {option: quota}}
 *
 * @example
 * const quotas = readQuotasFromStructure_(sheet);
 * // → { '6°1': { ITA: 6, CHAV: 10 }, '6°2': { ESP: 8 } }
 */
function readQuotasFromStructure_(sheet) {
  const quotas = {};

  try {
    const data = sheet.getDataRange().getValues();

    // ✅ CORRECTION : Recherche dynamique de l'en-tête (tolère lignes de garde/metadata)
    let headerRow = -1;
    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      // Chercher une ligne contenant CLASSE_DEST ou CLASSE_ORIGINE
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] || '').trim().toUpperCase();
        if (cell === 'CLASSE_DEST' || cell === 'CLASSE_ORIGINE') {
          headerRow = i;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1) {
      logLine('WARN', '⚠️ En-têtes non trouvés dans _STRUCTURE (cherché dans les 20 premières lignes)');
      return quotas;
    }

    logLine('INFO', '✅ En-tête trouvé à la ligne ' + (headerRow + 1));

    const headers = data[headerRow];

    // ✅ Trouver la colonne CLASSE_DEST et OPTIONS
    const classeCol = headers.indexOf('CLASSE_DEST');
    const optionsCol = headers.indexOf('OPTIONS');

    logLine('INFO', '🔍 readQuotasFromStructure: classeCol=' + classeCol + ', optionsCol=' + optionsCol);

    if (classeCol === -1 || optionsCol === -1) {
      logLine('WARN', '⚠️ Colonnes CLASSE_DEST ou OPTIONS introuvables dans _STRUCTURE');
      return quotas;
    }

    // Parcourir les lignes (à partir de headerRow + 1)
    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      const classe = String(row[classeCol] || '').trim();
      if (!classe) continue;

      const optionsStr = String(row[optionsCol] || '').trim();
      logLine('INFO', '🔍 readQuotasFromStructure: ' + classe + ' → OPTIONS="' + optionsStr + '"');

      quotas[classe] = {};

      // Parser le format "ITA=6, CHAV=10" ou "ITA=6 | CHAV=10"
      const parts = optionsStr.split(/[,|]/);
      parts.forEach(function(part) {
        const match = part.trim().match(/^([A-ZÉÈÀ]+)\s*=\s*(\d+)/);
        if (match) {
          const option = match[1];
          const quota = parseInt(match[2]) || 0;
          quotas[classe][option] = quota;
          logLine('INFO', '  ✅ ' + classe + '.' + option + ' = ' + quota);
        }
      });
    }

  } catch (e) {
    logLine('WARN', 'Erreur lecture quotas depuis _STRUCTURE : ' + e.message);
  }

  return quotas;
}

/**
 * Lit les effectifs cibles depuis _STRUCTURE
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - L'onglet _STRUCTURE
 * @returns {Object} Effectifs {classe: effectif}
 *
 * @example
 * const targets = readTargetsFromStructure_(sheet);
 * // → { '6°1': 26, '6°2': 25, '5°1': 27 }
 */
function readTargetsFromStructure_(sheet) {
  const targets = {};

  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // ✅ Trouver la colonne CLASSE_DEST et EFFECTIF
    const classeCol = headers.indexOf('CLASSE_DEST');
    const effectifCol = headers.indexOf('EFFECTIF');

    logLine('INFO', '🔍 readTargetsFromStructure: classeCol=' + classeCol + ', effectifCol=' + effectifCol);

    if (classeCol === -1 || effectifCol === -1) {
      logLine('WARN', '⚠️ Colonnes CLASSE_DEST ou EFFECTIF introuvables dans _STRUCTURE');
      return targets;
    }

    // Parcourir les lignes
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const classe = String(row[classeCol] || '').trim();
      if (!classe) continue;

      const effectif = parseInt(row[effectifCol]) || 25; // Fallback 25
      targets[classe] = effectif;

      logLine('INFO', '  ✅ ' + classe + ' effectif cible = ' + effectif);
    }

  } catch (e) {
    logLine('WARN', 'Erreur lecture effectifs depuis _STRUCTURE : ' + e.message);
  }

  return targets;
}

// ===================================================================
// ÉCRITURE DE DONNÉES
// ===================================================================

/**
 * Écrit des valeurs dans un onglet avec vérification
 *
 * @param {string} sheetName - Nom de l'onglet
 * @param {number} rangeStartRow - Ligne de départ (1-based)
 * @param {number} rangeStartCol - Colonne de départ (1-based)
 * @param {Array<Array>} values - Tableau 2D de valeurs
 * @param {number} headerRow - Ligne d'en-tête (0 si pas d'en-tête, 1 si ligne 1)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} L'onglet
 *
 * @example
 * writeAndVerify_('6°1TEST', 2, 1, [[1, 2], [3, 4]], 1);
 */
function writeAndVerify_(sheetName, rangeStartRow, rangeStartCol, values, headerRow) {
  const sh = getOrCreateSheet_(sheetName);

  if (headerRow) {
    // On ne touche pas l'entête ; on efface le dessous
    const lastRow = sh.getLastRow();
    if (lastRow > headerRow) {
      sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getMaxColumns()).clearContent();
    }
  } else {
    sh.clearContents();
  }

  if (values && values.length && values[0] && values[0].length) {
    sh.getRange(rangeStartRow, rangeStartCol, values.length, values[0].length).setValues(values);
  }

  SpreadsheetApp.flush();
  Utilities.sleep(100);

  // ✅ Vérification: au moins 1 ligne écrite sous l'entête si headerRow=1
  const rows = sh.getLastRow();
  const ok = headerRow ? (rows > headerRow) : (rows > 0);
  if (!ok) {
    throw new Error('WRITE_FAILED: rien d\'écrit dans ' + sheetName);
  }

  logLine('INFO', '✅ Écriture vérifiée dans ' + sheetName + ' (' + (rows - (headerRow || 0)) + ' lignes)');
  return sh;
}

/**
 * Écrit une liste d'élèves dans un onglet
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Le spreadsheet
 * @param {string} sheetName - Nom de l'onglet
 * @param {Array<Object>} eleves - Liste des élèves
 *
 * @example
 * writeElevesToSheet_(ss, '6°1CACHE', [{Nom: 'Dupont', Prénom: 'Jean'}, ...]);
 */
function writeElevesToSheet_(ss, sheetName, eleves) {
  const sheet = getOrCreateSheetByExactName_(ss, sheetName);
  sheet.clearContents();

  if (eleves.length === 0) {
    ss.setActiveSheet(sheet);
    SpreadsheetApp.flush();
    return;
  }

  // Récupérer les headers du premier élève
  const headers = Object.keys(eleves[0]);

  // Écrire les headers (ligne 1)
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Écrire les données
  const rows = eleves.map(e => headers.map(h => e[h] || ''));
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  ss.setActiveSheet(sheet);
  SpreadsheetApp.flush();
}

/**
 * Écrit directement dans un onglet CACHE
 *
 * @param {Object} ctx - Contexte (avec ss)
 * @param {string} baseClass - Nom de base de la classe (ex: '6°1')
 * @param {Array<Array>} values - Tableau 2D de valeurs
 *
 * @example
 * writeToCache_(ctx, '6°1', [['Nom', 'Prénom'], ['Dupont', 'Jean']]);
 */
function writeToCache_(ctx, baseClass, values) {
  const name = baseClass + 'CACHE';
  const sheet = getOrCreateSheetByExactName_(ctx.ss, name);
  sheet.clearContents();
  if (values && values.length) {
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
  ctx.ss.setActiveSheet(sheet);
  SpreadsheetApp.flush();
}

/**
 * Purge les onglets CACHE (garde les en-têtes)
 *
 * @param {Object} ctx - Contexte (avec cacheSheets, ss)
 *
 * @example
 * clearSheets_(ctx);
 */
function clearSheets_(ctx) {
  for (const sheetName of ctx.cacheSheets) {
    const sheet = ctx.ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
  }
}

/**
 * Écrit tous les états de classes dans les onglets CACHE
 * Purge d'abord, puis écrit, puis vérifie l'unicité des IDs
 *
 * @param {Object} ctx - Contexte (avec ss, cacheSheets)
 * @param {Object} classesState - État {classe: [eleves]}
 *
 * @example
 * writeAllClassesToCACHE_(ctx, { '6°1': [...], '6°2': [...] });
 */
function writeAllClassesToCACHE_(ctx, classesState) {
  // 1. Purger les feuilles CACHE
  clearSheets_(ctx);

  // 2. Écrire chaque classe dans son onglet CACHE
  Object.keys(classesState).forEach(function(classe) {
    const eleves = classesState[classe];
    const cacheName = classe + 'CACHE';

    if (eleves.length === 0) {
      logLine('WARN', '⚠️ Classe ' + classe + ' vide, onglet CACHE non écrit');
      return;
    }

    writeElevesToSheet_(ctx.ss, cacheName, eleves);
    logLine('INFO', '✅ Écrit ' + eleves.length + ' élèves dans ' + cacheName);
  });

  // 3. ✅ AUDIT : Vérifier l'unicité des IDs dans tous les CACHE
  const allIds = {};
  let duplicates = [];

  Object.keys(classesState).forEach(function(classe) {
    const eleves = classesState[classe];
    eleves.forEach(function(e) {
      const id = e.ID || e.Id || e.id;
      if (!id) return;

      if (allIds[id]) {
        duplicates.push({
          id: id,
          classe1: allIds[id],
          classe2: classe
        });
      } else {
        allIds[id] = classe;
      }
    });
  });

  if (duplicates.length > 0) {
    logLine('ERROR', '❌ IDs DUPLIQUÉS détectés dans CACHE :');
    duplicates.forEach(function(dup) {
      logLine('ERROR', '  - ID ' + dup.id + ' présent dans ' + dup.classe1 + ' ET ' + dup.classe2);
    });
    throw new Error('IDS_DUPLIQUES: ' + duplicates.length + ' élèves en double dans CACHE');
  }

  logLine('INFO', '✅ Tous les IDs sont uniques dans CACHE (' + Object.keys(allIds).length + ' élèves)');
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
