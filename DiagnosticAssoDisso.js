/**
 * ═══════════════════════════════════════════════════════════════════════════
 * DIAGNOSTIC ASSO/DISSO - Vérifier l'ordre des colonnes
 * ═══════════════════════════════════════════════════════════════════════════
 */

function diagnosticAssoDisso() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const consoSheet = ss.getSheetByName('CONSOLIDATION');
  
  if (!consoSheet) {
    Logger.log("❌ CONSOLIDATION n'existe pas");
    return;
  }
  
  const data = consoSheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  Logger.log("🔍 DIAGNOSTIC ASSO/DISSO");
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
  
  Logger.log("📋 En-têtes CONSOLIDATION:");
  headers.forEach((h, i) => {
    Logger.log(`  Colonne ${i + 1}: ${h}`);
  });
  
  const assoIdx = headers.indexOf('ASSO');
  const dissoIdx = headers.indexOf('DISSO');
  
  Logger.log(`\n📍 Indices des colonnes:`);
  Logger.log(`  ASSO: ${assoIdx} (colonne ${assoIdx + 1})`);
  Logger.log(`  DISSO: ${dissoIdx} (colonne ${dissoIdx + 1})`);
  
  if (assoIdx === -1 || dissoIdx === -1) {
    Logger.log("\n❌ Une ou les deux colonnes sont manquantes!");
    return;
  }
  
  // Analyser les données
  Logger.log(`\n📊 Analyse des données (${data.length - 1} élèves):`);
  
  let countAsso = 0;
  let countDisso = 0;
  const assoValues = [];
  const dissoValues = [];
  
  for (let i = 1; i < data.length; i++) {
    const asso = String(data[i][assoIdx] || '').trim();
    const disso = String(data[i][dissoIdx] || '').trim();
    
    if (asso) {
      countAsso++;
      if (!assoValues.includes(asso)) assoValues.push(asso);
    }
    if (disso) {
      countDisso++;
      if (!dissoValues.includes(disso)) dissoValues.push(disso);
    }
  }
  
  Logger.log(`  ASSO: ${countAsso} élèves, ${assoValues.length} codes différents`);
  Logger.log(`    Codes: ${assoValues.join(', ') || 'aucun'}`);
  
  Logger.log(`  DISSO: ${countDisso} élèves, ${dissoValues.length} codes différents`);
  Logger.log(`    Codes: ${dissoValues.join(', ') || 'aucun'}`);
  
  // Vérifier les onglets sources
  Logger.log(`\n📁 Vérification des onglets sources:`);
  const sourceSheets = ss.getSheets().filter(s => /.+°\d+$/.test(s.getName()));
  
  sourceSheets.forEach(sheet => {
    const sData = sheet.getDataRange().getValues();
    const sHeaders = sData[0];
    const sAssoIdx = sHeaders.indexOf('ASSO');
    const sDissoIdx = sHeaders.indexOf('DISSO');
    
    Logger.log(`\n  ${sheet.getName()}:`);
    Logger.log(`    ASSO: colonne ${sAssoIdx + 1}`);
    Logger.log(`    DISSO: colonne ${sDissoIdx + 1}`);
    
    // Afficher quelques exemples
    for (let i = 1; i < Math.min(4, sData.length); i++) {
      const asso = sData[i][sAssoIdx] || '';
      const disso = sData[i][sDissoIdx] || '';
      if (asso || disso) {
        Logger.log(`    Ligne ${i + 1}: ASSO="${asso}", DISSO="${disso}"`);
      }
    }
  });
  
  Logger.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
}

/**
 * Inverser les colonnes ASSO et DISSO si elles sont mal placées
 */
function inverserAssoDisso() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Inverser ASSO et DISSO',
    'Cette action va échanger le contenu des colonnes ASSO et DISSO dans tous les onglets.\n\nÊtes-vous sûr?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    Logger.log("❌ Annulé par l'utilisateur");
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. CONSOLIDATION
  const consoSheet = ss.getSheetByName('CONSOLIDATION');
  if (consoSheet) {
    Logger.log("🔄 Inversion dans CONSOLIDATION...");
    inverserColonnes(consoSheet, 'ASSO', 'DISSO');
  }
  
  // 2. Onglets sources
  const sourceSheets = ss.getSheets().filter(s => /.+°\d+$/.test(s.getName()));
  sourceSheets.forEach(sheet => {
    Logger.log(`🔄 Inversion dans ${sheet.getName()}...`);
    inverserColonnes(sheet, 'ASSO', 'DISSO');
  });
  
  SpreadsheetApp.flush();
  
  ui.alert('✅ Succès', 'Les colonnes ASSO et DISSO ont été inversées.', ui.ButtonSet.OK);
  Logger.log("✅ Inversion terminée!");
}

function inverserColonnes(sheet, col1Name, col2Name) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idx1 = headers.indexOf(col1Name);
  const idx2 = headers.indexOf(col2Name);
  
  if (idx1 === -1 || idx2 === -1) {
    Logger.log(`  ⚠️ Colonnes ${col1Name} ou ${col2Name} non trouvées`);
    return;
  }
  
  // Échanger les valeurs pour chaque ligne
  for (let i = 1; i < data.length; i++) {
    const temp = data[i][idx1];
    data[i][idx1] = data[i][idx2];
    data[i][idx2] = temp;
  }
  
  // Écrire les données modifiées
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}
