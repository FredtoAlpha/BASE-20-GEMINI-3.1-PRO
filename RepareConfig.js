/**
 * =====================================================================
 * RÉPARATION AUTOMATIQUE DE _CONFIG DEPUIS CONSOLIDATION
 * =====================================================================
 * Scanne CONSOLIDATION et reconstruit _CONFIG automatiquement
 * Sans détruire les données existantes !
 */

function repareConfigDepuisConsolidation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const conso = ss.getSheetByName("CONSOLIDATION");
    
    if (!conso || conso.getLastRow() <= 1) {
      return {
        success: false,
        error: "CONSOLIDATION est vide ou n'existe pas. Impossible de réparer."
      };
    }
    
    Logger.log("🔧 RÉPARATION DE _CONFIG DEPUIS CONSOLIDATION...");
    
    // Lire les données de CONSOLIDATION
    const data = conso.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    Logger.log(`📊 ${rows.length} lignes trouvées dans CONSOLIDATION`);
    
    // Trouver les indices des colonnes
    const idx = {
      SOURCE: headers.indexOf('SOURCE'),
      LV2: headers.indexOf('LV2'),
      OPT: headers.indexOf('OPT'),
      DISPOSITIF: headers.indexOf('DISPOSITIF')
    };
    
    Logger.log(`📋 Colonnes détectées: SOURCE=${idx.SOURCE}, LV2=${idx.LV2}, OPT=${idx.OPT}, DISPOSITIF=${idx.DISPOSITIF}`);
    
    // Extraire les valeurs uniques
    const sourcesSet = new Set();
    const lv2Set = new Set();
    const optSet = new Set();
    const dispoSet = new Set();
    
    rows.forEach(row => {
      if (idx.SOURCE >= 0 && row[idx.SOURCE]) sourcesSet.add(String(row[idx.SOURCE]).trim());
      if (idx.LV2 >= 0 && row[idx.LV2]) lv2Set.add(String(row[idx.LV2]).trim());
      if (idx.OPT >= 0 && row[idx.OPT]) optSet.add(String(row[idx.OPT]).trim());
      if (idx.DISPOSITIF >= 0 && row[idx.DISPOSITIF]) dispoSet.add(String(row[idx.DISPOSITIF]).trim());
    });
    
    // Convertir en listes
    const sources = Array.from(sourcesSet).filter(s => s && s !== "");
    const lv2List = Array.from(lv2Set).filter(s => s && s !== "");
    const optList = Array.from(optSet).filter(s => s && s !== "");
    const dispoList = Array.from(dispoSet).filter(s => s && s !== "");
    
    Logger.log(`✅ SOURCES: ${sources.length} trouvées: ${sources.join(', ')}`);
    Logger.log(`✅ LV2: ${lv2List.length} trouvées: ${lv2List.join(', ')}`);
    Logger.log(`✅ OPTIONS: ${optList.length} trouvées: ${optList.join(', ')}`);
    Logger.log(`✅ DISPOSITIFS: ${dispoList.length} trouvés: ${dispoList.join(', ')}`);
    
    // Détecter le niveau depuis la première source
    let niveau = "6e"; // Défaut
    if (sources.length > 0) {
      const firstSource = sources[0];
      // Extraire le niveau (ex: "6°1" → "6e")
      const match = firstSource.match(/^(\d+)[°e]/i);
      if (match) {
        niveau = match[1] + "e";
      }
    }
    
    // Nombre de classes cibles = sources
    const nbSources = sources.length;
    const nbDest = nbSources; // Par défaut, même nombre
    
    Logger.log(`🎯 NIVEAU détecté: ${niveau}`);
    Logger.log(`🎯 NB_SOURCES: ${nbSources}`);
    Logger.log(`🎯 NB_DEST: ${nbDest}`);
    
    // Reconstruire _CONFIG
    let configSheet = ss.getSheetByName("_CONFIG");
    if (!configSheet) {
      configSheet = ss.insertSheet("_CONFIG");
      Logger.log("📄 _CONFIG créé");
    }
    
    configSheet.clear();
    
    // Écrire les nouvelles valeurs
    const configData = [
      ["Paramètre", "Valeur"],
      ["NIVEAU", niveau],
      ["NB_SOURCES", nbSources],
      ["NB_DEST", nbDest],
      ["LV2", lv2List.join(', ')],
      ["OPT", optList.join(', ')],
      ["DISPOSITIF", dispoList.join(', ')],
      ["DATE_REPAIR", new Date().toISOString()]
    ];
    
    configSheet.getRange(1, 1, configData.length, 2).setValues(configData);
    
    // Formater
    configSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4a5568').setFontColor('#ffffff');
    configSheet.setColumnWidth(1, 200);
    configSheet.setColumnWidth(2, 400);
    
    Logger.log("✅ _CONFIG réparé avec succès !");
    
    return {
      success: true,
      message: `✅ _CONFIG réparé !\n\n🔹 NIVEAU: ${niveau}\n🔹 NB_DEST: ${nbDest}\n🔹 LV2: ${lv2List.join(', ')}\n🔹 OPT: ${optList.join(', ')}\n🔹 DISPOSITIF: ${dispoList.join(', ')}\n\n${rows.length} élèves analysés`,
      config: {
        NIVEAU: niveau,
        NB_SOURCES: nbSources,
        NB_DEST: nbDest,
        LV2: lv2List.join(', '),
        OPT: optList.join(', '),
        DISPOSITIF: dispoList.join(', ')
      }
    };
    
  } catch (e) {
    Logger.log(`❌ ERREUR: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Wrapper pour l'appeler depuis l'interface
 */
function v3_repareConfig() {
  return repareConfigDepuisConsolidation();
}
