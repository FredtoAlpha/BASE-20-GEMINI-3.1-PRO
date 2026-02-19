/**
 * ===================================================================
 * BACKEND_FINALISATION.GS - CRÉATION ET FORMATAGE ONGLETS FIN
 * ===================================================================
 * Module responsable de la création des onglets FIN avec formatage
 * Couleurs, bordures, styles pédagogiques
 * Extrait du Code.gs originel (Lignes 340-417)
 * ===================================================================
 */

// Configuration du formatage FIN
const FIN_FORMATTING_CONFIG = {
  headerColor: '#2c3e50',
  headerFontColor: '#ffffff',
  gridLineColor: '#ecf0f1',
  alternateRowColor: '#f8f9fa',
  teteLineColor: '#3498db',
  niv1LineColor: '#e74c3c',
  normalLineColor: '#ffffff',
  columnWidth: 15,
  fontSize: 11,
  // Couleurs par LV2 (couleurs des pays)
  lv2Colors: {
    'ESP': '#FFB347',     // Orange (Espagne)
    'ITA': '#d5f5e3',     // Vert personnalisé (Italie)
    'ALL': '#FFED4E',     // Jaune (Allemagne)
    'PT': '#32CD32',      // Vert (Portugal)
    'OR': '#FFD700'       // Or (Option OR)
  },
  // Couleurs par OPT (avec meilleur contraste)
  optColors: {
    'CHAV': '#8B4789',    // Violet plus foncé (CHAV) - meilleur contraste
    'LATIN': '#e8f8f5',   // Vert d'eau (LATIN)
    'CHINOIS': '#C41E3A', // Rouge cardinal (CHINOIS)
    'GREC': '#f6ca9d'     // Orange clair (GREC)
  }
};

/**
 * Formate un onglet FIN avec couleurs et styles pédagogiques
 * @param {Sheet} sheet - L'onglet à formater
 * @param {Array} rowData - Les données des élèves
 * @param {Array} headers - Les en-têtes
 */
function formatFinSheet(sheet, rowData, headers) {
  if (!sheet || rowData.length === 0) {
    Logger.log('[WARN] formatFinSheet : Données invalides');
    return;
  }

  try {
    // 0. CACHER LES COLONNES A, B ET C
    sheet.hideColumns(1, 3);
    
    // 1. FORMATAGE DES EN-TÊTES
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(FIN_FORMATTING_CONFIG.headerColor);
    headerRange.setFontColor(FIN_FORMATTING_CONFIG.headerFontColor);
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(FIN_FORMATTING_CONFIG.fontSize + 1); // Header plus grand
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');

    // 2. HAUTEUR DES LIGNES
    sheet.setRowHeight(1, 25); // En-tête plus haut

    // 3. LARGEUR DES COLONNES
    for (let col = 1; col <= headers.length; col++) {
      const width = getOptimalColumnWidth(headers[col - 1]);
      sheet.setColumnWidth(col, width);
    }

    // 4. FORMATAGE DES DONNÉES ÉLÈVES
    const idColIndex = headers.indexOf('ID_ELEVE') + 1;
    const comColIndex = headers.indexOf('COM') + 1;
    const traColIndex = headers.indexOf('TRA') + 1;
    const sexeColIndex = headers.indexOf('SEXE') + 1;
    const lv2ColIndex = headers.indexOf('LV2') + 1;
    const optColIndex = headers.indexOf('OPT') + 1;

    for (let i = 0; i < rowData.length; i++) {
      const rowNum = i + 2; // +1 pour en-tête, +1 pour indices
      const row = rowData[i];

      // Déterminer la couleur de la ligne selon LV2/OPT
      let backgroundColor = FIN_FORMATTING_CONFIG.normalLineColor;
      let fontColor = '#000000';

      if (row && row.length > 0) {
        // Lire LV2 et OPT
        const lv2Value = lv2ColIndex > 0 ? String(row[lv2ColIndex - 1] || '').trim().toUpperCase() : '';
        const optValue = optColIndex > 0 ? String(row[optColIndex - 1] || '').trim().toUpperCase() : '';
        
        // Priorité 1 : Couleur par OPT (si présent)
        if (optValue && FIN_FORMATTING_CONFIG.optColors[optValue]) {
          backgroundColor = FIN_FORMATTING_CONFIG.optColors[optValue];
        }
        // Priorité 2 : Couleur par LV2 (si présent et pas d'OPT)
        else if (lv2Value && FIN_FORMATTING_CONFIG.lv2Colors[lv2Value]) {
          backgroundColor = FIN_FORMATTING_CONFIG.lv2Colors[lv2Value];
        }
        // Priorité 3 : Tête de classe ou Niv1 (logique existante)
        else if (row.length > Math.max(comColIndex, traColIndex)) {
          const COM = Number(row[comColIndex - 1]) || 2.5;
          const TRA = Number(row[traColIndex - 1]) || 2.5;

          // Tête de classe (Bleu clair)
          if (COM >= 4 || TRA >= 4 || (COM + TRA) / 2 >= 3.5) {
            backgroundColor = FIN_FORMATTING_CONFIG.teteLineColor;
            fontColor = '#ffffff';
          }
          // Élève en difficulté (Rose/Rouge)
          else if (COM <= 1 || TRA <= 1) {
            backgroundColor = FIN_FORMATTING_CONFIG.niv1LineColor;
            fontColor = '#ffffff';
          }
          // Alternance de lignes (pour lisibilité)
          else if (i % 2 === 0) {
            backgroundColor = FIN_FORMATTING_CONFIG.alternateRowColor;
          }
        }
      }

      // Appliquer le formatage à la ligne
      const rowRange = sheet.getRange(rowNum, 1, 1, headers.length);
      rowRange.setBackground(backgroundColor);
      rowRange.setFontColor(fontColor);
      rowRange.setFontWeight('bold'); // ✅ TOUT EN GRAS
      rowRange.setFontSize(FIN_FORMATTING_CONFIG.fontSize); // ✅ TAILLE POLICE
      rowRange.setVerticalAlignment('middle');

      // Centrer les colonnes numériques
      if (comColIndex > 0) {
        sheet.getRange(rowNum, comColIndex).setHorizontalAlignment('center');
      }
      if (traColIndex > 0) {
        sheet.getRange(rowNum, traColIndex).setHorizontalAlignment('center');
      }
      if (sexeColIndex > 0) {
        sheet.getRange(rowNum, sexeColIndex).setHorizontalAlignment('center');
      }
    }

    // 5. AJOUTER DES BORDURES
    const allDataRange = sheet.getRange(1, 1, rowData.length + 1, headers.length);
    allDataRange.setBorder(true, true, true, true, true, true);

    SpreadsheetApp.flush();
    Logger.log(`[SUCCESS] Formatage FIN appliqué à ${sheet.getName()}`);

  } catch (e) {
    Logger.log(`[ERROR] Erreur formatage FIN : ${e.toString()}`);
  }
}

/**
 * Détermine la largeur optimale d'une colonne
 * @param {string} headerName - Nom de l'en-tête
 * @returns {number} Largeur en pixels
 */
function getOptimalColumnWidth(headerName) {
  const headerLength = String(headerName).length;

  if (headerName === 'ID_ELEVE' || headerName === 'SEXE') return 80;
  if (headerName === 'NOM' || headerName === 'PRENOM') return 120;
  if (headerName === 'COM' || headerName === 'TRA' || headerName === 'PART') return 70;

  return Math.max(100, headerLength * 12);
}

/**
 * Crée les onglets FIN à partir des dispositions
 * @param {Object} disposition - Les résultats de répartition
 * @param {string} mode - 'finalize' ou 'test'
 * @returns {Object} Résultat de création
 */
function finalizeClasses(disposition, mode = 'finalize') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = {
    created: [],
    failed: [],
    formatted: []
  };

  if (!disposition || typeof disposition !== 'object') {
    Logger.log('[ERROR] Disposition invalide pour finalisation');
    return { ok: false, error: 'Disposition invalide', results: results };
  }

  try {
    for (const className in disposition) {
      const classData = disposition[className];
      if (!classData || !classData.students || !Array.isArray(classData.students)) {
        results.failed.push(className);
        continue;
      }

      const finSheetName = `${className}FIN`;

      // 1. CRÉER OU OBTENIR L'ONGLET FIN
      let finSheet = ss.getSheetByName(finSheetName);
      if (!finSheet) {
        finSheet = ss.insertSheet(finSheetName);
        Logger.log(`[INFO] Onglet ${finSheetName} créé`);
      } else {
        finSheet.clearContents();
        Logger.log(`[INFO] Onglet ${finSheetName} vidé`);
      }

      // 2. ÉCRIRE LES DONNÉES
      if (classData.students.length > 0) {
        const headersRow = classData.headers || [];
        const allRows = [headersRow, ...classData.students];

        finSheet.getRange(1, 1, allRows.length, headersRow.length).setValues(allRows);
        results.created.push(finSheetName);

        // 3. APPLIQUER LE FORMATAGE
        formatFinSheet(finSheet, classData.students, headersRow);
        results.formatted.push(finSheetName);
      }
    }

    SpreadsheetApp.flush();
    Logger.log(`[SUCCESS] Finalisation complète : ${results.created.length} onglets créés`);

    return {
      ok: true,
      results: results
    };

  } catch (e) {
    Logger.log(`[ERROR] Erreur finalisation : ${e.toString()}`);
    return {
      ok: false,
      error: e.toString(),
      results: results
    };
  }
}

/**
 * Applique des styles de légende pour expliquer le formatage
 * @param {Sheet} sheet - L'onglet FIN
 */
function applyLegend(sheet) {
  try {
    const legendRow = sheet.getLastRow() + 3;

    sheet.getRange(legendRow, 1).setValue('LÉGENDE :');
    sheet.getRange(legendRow, 1).setFontWeight('bold');

    // Tête de classe
    const tetelRow = legendRow + 1;
    sheet.getRange(tetelRow, 1).setBackground(FIN_FORMATTING_CONFIG.teteLineColor);
    sheet.getRange(tetelRow, 1).setFontColor('#ffffff');
    sheet.getRange(tetelRow, 1).setValue('Tête');
    sheet.getRange(tetelRow, 2).setValue('COM ≥ 4 ou TRA ≥ 4');

    // Élève en difficulté
    const niv1Row = legendRow + 2;
    sheet.getRange(niv1Row, 1).setBackground(FIN_FORMATTING_CONFIG.niv1LineColor);
    sheet.getRange(niv1Row, 1).setFontColor('#ffffff');
    sheet.getRange(niv1Row, 1).setValue('Niv1');
    sheet.getRange(niv1Row, 2).setValue('COM ≤ 1 ou TRA ≤ 1');

    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`[WARN] Erreur application légende : ${e.toString()}`);
  }
}

/**
 * Supprime les onglets FIN existants
 * @returns {number} Nombre d'onglets supprimés
 */
function deleteFinSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deletedCount = 0;

  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName().endsWith('FIN')) {
      ss.deleteSheet(sheet);
      deletedCount++;
      Logger.log(`[INFO] Onglet ${sheet.getName()} supprimé`);
    }
  });

  return deletedCount;
}

/**
 * Récupère tous les onglets FIN existants
 * @returns {Array} Liste des noms d'onglets FIN
 */
function getFinSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(s => s.getName().endsWith('FIN'))
    .map(s => s.getName());
}

/**
 * Applique une protection read-only sur un onglet FIN
 * @param {string} sheetName - Nom de l'onglet
 * @returns {boolean} Succès
 */
function protectFinSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`[WARN] Onglet ${sheetName} introuvable pour protection`);
      return false;
    }

    const protection = sheet.protect()
      .setDescription('Onglet FIN - Lecture seule');

    const ui = SpreadsheetApp.getUi();
    protection.removeEditors([Session.getEffectiveUser().getEmail()]);

    Logger.log(`[SUCCESS] Onglet ${sheetName} protégé`);
    return true;
  } catch (e) {
    Logger.log(`[WARN] Impossible de protéger ${sheetName} : ${e.toString()}`);
    return false;
  }
}
