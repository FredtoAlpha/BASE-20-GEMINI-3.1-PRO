// ... (début du fichier ConsolePilotage_Server.gs) ...

/**
 * Fournit le contexte de pont à l'InterfaceV2 et le supprime ensuite.
 * C'est la fonction appelée par l'InterfaceV2 à son initialisation.
 * @returns {Object} Un objet contenant {success: Boolean, context: Object|null}.
 */
/**
 * Copie les onglets ...TEST vers ...DEF.
 * C'est l'action finale et irréversible de la console.
 */
function finalizeProcess() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testSheets = ss.getSheets().filter(s => s.getName().endsWith('TEST'));

    if (testSheets.length === 0) {
      throw new Error("Aucun onglet ...TEST à finaliser.");
    }

    let finalizedCount = 0;
    testSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const finalName = sheetName.replace(/TEST$/, 'DEF');

      // Supprimer l'ancien onglet DEF s'il existe
      const oldDefSheet = ss.getSheetByName(finalName);
      if (oldDefSheet) {
        ss.deleteSheet(oldDefSheet);
      }

      // Copier l'onglet TEST vers le nouvel onglet DEF
      const newDefSheet = sheet.copyTo(ss);
      newDefSheet.setName(finalName);

      // Rendre la feuille visible et la protéger (facultatif)
      newDefSheet.showSheet();
      
      // ✅ FORMATER L'ONGLET DEF (même formatage que TEST/FIN)
      formatDefSheet_ConsolePilotage(newDefSheet);

      finalizedCount++;
    });

    return { success: true, message: `${finalizedCount} classe(s) ont été finalisées avec succès.` };
  } catch (e) {
    console.error(`Erreur dans finalizeProcess: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * Formate un onglet DEF avec les mêmes règles que TEST/FIN
 * - Cache colonnes A et B
 * - Police en gras partout
 * - Taille police augmentée
 * - Couleurs par LV2/OPT
 * 
 * @param {Sheet} sheet - L'onglet DEF à formater
 */
function formatDefSheet_ConsolePilotage(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return;
  
  try {
    const CONFIG = {
      fontSize: 11,
      lv2Colors: {
        'ESP': '#FFB347',     // Orange (Espagne)
        'ITA': '#d5f5e3',     // Vert personnalisé (Italie)
        'ALL': '#FFED4E',     // Jaune (Allemagne)
        'PT': '#32CD32',      // Vert (Portugal)
        'OR': '#FFD700'       // Or
      },
      optColors: {
        'CHAV': '#8B4789',    // Violet plus foncé - meilleur contraste
        'LATIN': '#e8f8f5',   // Vert d'eau
        'CHINOIS': '#C41E3A', // Rouge cardinal
        'GREC': '#f6ca9d'     // Orange clair
      }
    };
    
    // 1. CACHER COLONNES A, B ET C
    sheet.hideColumns(1, 3);
    
    // 2. TOUT EN GRAS + TAILLE POLICE
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      const allRange = sheet.getRange(1, 1, lastRow, lastCol);
      allRange.setFontWeight('bold');
      allRange.setFontSize(CONFIG.fontSize);
      
      // En-tête plus grand
      sheet.getRange(1, 1, 1, lastCol).setFontSize(CONFIG.fontSize + 1);
    }
    
    // 3. APPLIQUER COULEURS PAR LV2/OPT
    if (lastRow > 1) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idxLV2 = headers.indexOf('LV2');
      const idxOPT = headers.indexOf('OPT');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowNum = i + 1;
        
        const lv2Value = idxLV2 >= 0 ? String(row[idxLV2] || '').trim().toUpperCase() : '';
        const optValue = idxOPT >= 0 ? String(row[idxOPT] || '').trim().toUpperCase() : '';
        
        let backgroundColor = null;
        
        // Priorité 1 : Couleur par OPT
        if (optValue && CONFIG.optColors[optValue]) {
          backgroundColor = CONFIG.optColors[optValue];
        }
        // Priorité 2 : Couleur par LV2
        else if (lv2Value && CONFIG.lv2Colors[lv2Value]) {
          backgroundColor = CONFIG.lv2Colors[lv2Value];
        }
        
        // Appliquer la couleur
        if (backgroundColor) {
          sheet.getRange(rowNum, 1, 1, headers.length).setBackground(backgroundColor);
        }
      }
    }
    
    // 4. FIGER LA PREMIÈRE LIGNE
    sheet.setFrozenRows(1);
    
    console.log('[INFO] Onglet DEF ' + sheet.getName() + ' formaté avec succès');
    
  } catch (e) {
    console.log('[WARN] Erreur formatage DEF ' + sheet.getName() + ': ' + e.message);
  }
}

/**
 * Stores the necessary context for InterfaceV2 to initialize.
 * @param {string} mode - The mode to load (e.g., 'TEST').
 */
function setBridgeContext(mode, sourceSheetName) {
  try {
    const context = {
      mode: mode,
      sourceSheetName: sourceSheetName,
      timestamp: new Date().toISOString()
    };
    PropertiesService.getUserProperties().setProperty('JULES_CONTEXT', JSON.stringify(context));
    return { success: true };
  } catch (e) {
    console.error(`Erreur dans setBridgeContext: ${e.message}`);
    return { success: false, error: e.message };
  }
}
