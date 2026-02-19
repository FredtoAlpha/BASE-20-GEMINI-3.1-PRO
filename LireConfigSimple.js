/**
 * ============================================================
 * FONCTION GÉNÉRIQUE: Lit TOUS les paramètres depuis _CONFIG
 * ============================================================
 * Cette fonction lit dynamiquement tous les paramètres de _CONFIG
 * sans aucune valeur codée en dur. Elle retourne un objet avec
 * tous les paramètres trouvés.
 */
function lireTousLesParametresConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("_CONFIG");
    
    if (!configSheet) {
      Logger.log("❌ lireTousLesParametresConfig: Feuille _CONFIG introuvable");
      return {};
    }
    
    const data = configSheet.getDataRange().getValues();
    Logger.log(`📋 lireTousLesParametresConfig: ${data.length} lignes lues`);
    
    // Chercher les en-têtes (PARAMETRE/PARAMÈTRE et VALEUR)
    let paramCol = -1, valCol = -1, startRow = 0;
    
    for (let i = 0; i < Math.min(3, data.length); i++) {
      for (let j = 0; j < data[i].length; j++) {
        const cell = String(data[i][j] || '').trim().toUpperCase();
        if (cell === "PARAMETRE" || cell === "PARAMÈTRE") paramCol = j;
        if (cell === "VALEUR") valCol = j;
      }
      
      if (paramCol !== -1 && valCol !== -1) {
        startRow = i + 1;
        Logger.log(`✅ En-têtes trouvés à la ligne ${i+1}: paramCol=${paramCol}, valCol=${valCol}`);
        break;
      }
    }
    
    if (paramCol === -1 || valCol === -1) {
      Logger.log("❌ En-têtes PARAMETRE/VALEUR non trouvés");
      return {};
    }
    
    // Lire TOUS les paramètres dynamiquement
    const params = {};
    
    for (let i = startRow; i < data.length; i++) {
      const paramKey = String(data[i][paramCol] || '').trim();
      let val = data[i][valCol];
      
      // Ignorer les lignes vides
      if (!paramKey) continue;
      
      // Conversion automatique des types
      if (typeof val === 'string') {
        const trimmed = val.trim();
        
        // Booléen
        if (trimmed.toLowerCase() === 'true') {
          val = true;
        } else if (trimmed.toLowerCase() === 'false') {
          val = false;
        }
        // Nombre
        else if (trimmed !== '' && !isNaN(Number(trimmed))) {
          val = Number(trimmed);
        }
        // Sinon garder la chaîne
      }
      
      // Stocker avec la clé exacte (préserver la casse)
      params[paramKey] = val;
      Logger.log(`  📌 ${paramKey} = ${JSON.stringify(val)} (${typeof val})`);
    }
    
    Logger.log(`📊 Total de ${Object.keys(params).length} paramètres lus depuis _CONFIG`);
    return params;
    
  } catch (e) {
    Logger.log(`❌ Erreur lireTousLesParametresConfig: ${e.message}`);
    return {};
  }
}

/**
 * Fonction de compatibilité pour l'ancien code
 */
function lireNbClassesDepuisConfig() {
  const params = lireTousLesParametresConfig();
  return {
    NB_SOURCES: params.NB_SOURCES || 6,
    NB_DEST: params.NB_DEST || 6
  };
}

/**
 * Test de la fonction
 */
function TEST_LIRE_TOUS_PARAMS() {
  Logger.log("\n=== TEST: lireTousLesParametresConfig ===");
  const params = lireTousLesParametresConfig();
  
  Logger.log("\n📋 Paramètres lus:");
  for (const key in params) {
    Logger.log(`  ${key} = ${JSON.stringify(params[key])}`);
  }
}
