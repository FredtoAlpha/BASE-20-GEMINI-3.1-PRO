/**
 * BANDEAU DE MAINTENANCE
 * 
 * Ce fichier est en cours de maintenance. Tous les utilitaires sont centralisés dans Utils.js.
 * Veuillez vous référer à ce fichier pour toute modification ou ajout de fonctionnalités.
 */

/**
 * Fonction principale pour ajouter/mettre à jour les listes déroulantes et le formatage conditionnel,
 * injecter la formule ARRAYFORMULA NOM_PRENOM et ajuster finement les largeurs de colonnes.
 * (Version universelle - adaptée pour gérer dynamiquement les options de LV2 et OPT)
 */
function ajouterListesDeroulantes() {
  Logger.log("====== Début ajouterListesDeroulantes (version universelle) ======");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  /* ------------------------------------------------------------------
   * 1️⃣  Lecture / normalisation de la configuration ------------------
   * ----------------------------------------------------------------*/
  let config = {};
  
  // Initialisation de la configuration par défaut
  config.NIVEAU = "";
  config.LV2 = [];
  config.OPT = [];
  config.LV2_AVAILABLE = [];
  config.OPT_AVAILABLE = [];

  try {
    // Essayer de lire la config via la fonction getConfig()
    const systemConfig = getConfig();
    if (systemConfig) {
      config = systemConfig;
    }
  } catch (e) {
    Logger.log(`Info: getConfig() non disponible ou erreur: ${e.message}`);
  }

  // Récupération de la configuration depuis _CONFIG
  try {
    const cfgSheet = ss.getSheetByName("_CONFIG");
    if (cfgSheet) {
      const data = cfgSheet.getDataRange().getValues();
      data.forEach(row => {
        const [cle, valeurs] = row;
        if (cle === "LV2" && valeurs) {
          config.LV2 = String(valeurs).split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
          if (!config.LV2_AVAILABLE || !config.LV2_AVAILABLE.length) {
            config.LV2_AVAILABLE = config.LV2;
          }
          Logger.log(`LV2 récupérées depuis _CONFIG : ${config.LV2.join(',')}`);
        }
        if (cle === "OPT" && valeurs) {
          config.OPT = String(valeurs).split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
          if (!config.OPT_AVAILABLE || !config.OPT_AVAILABLE.length) {
            config.OPT_AVAILABLE = config.OPT;
          }
          Logger.log(`OPT récupérées depuis _CONFIG : ${config.OPT.join(',')}`);
        }
      });
    }
  } catch (e) {
    Logger.log(`Erreur lecture _CONFIG : ${e.message}`);
  }

  // Récupération de la configuration depuis _STRUCTURE si nécessaire
  if ((!config.OPT || !config.OPT.length) || (!config.LV2 || !config.LV2.length)) {
    try {
      const structureSheet = ss.getSheetByName("_STRUCTURE");
      if (structureSheet) {
        Logger.log("Extraction des options depuis _STRUCTURE...");
        const structData = structureSheet.getDataRange().getValues();
        const optHeader = structData[0].indexOf("OPTIONS");
        
        if (optHeader !== -1) {
          const optValues = structData.slice(1)
                            .map(row => row[optHeader])
                            .filter(val => val && typeof val === 'string')
                            .map(val => val.includes("=") ? val.split("=")[0].trim() : val.trim())
                            .filter(val => val);
          
          // Extraction des OPT uniques
          if (!config.OPT || !config.OPT.length) {
            config.OPT = [...new Set(optValues)];
            config.OPT_AVAILABLE = config.OPT;
            Logger.log(`OPT récupérées depuis _STRUCTURE : ${config.OPT.join(',')}`);
          }
          
          // Si on trouve des LV2 aussi
          const possibleLV2 = optValues.filter(v => ["ALL", "ESP", "ITA", "CHI"].includes(v));
          if (possibleLV2.length > 0 && (!config.LV2 || !config.LV2.length)) {
            config.LV2 = [...new Set(possibleLV2)];
            config.LV2_AVAILABLE = config.LV2;
            Logger.log(`LV2 récupérées depuis _STRUCTURE : ${config.LV2.join(',')}`);
          }
        }
      }
    } catch (e) {
      Logger.log(`Erreur lecture _STRUCTURE : ${e.message}`);
    }
  }

  // Valeurs par défaut si rien n'a été trouvé
  if (!config.LV2 || !config.LV2.length) {
    config.LV2 = ["ALL", "ESP", "ITA", "CHI"];
    config.LV2_AVAILABLE = config.LV2;
    Logger.log("Utilisation des LV2 par défaut: ALL,ESP,ITA,CHI");
  }

  if (!config.OPT || !config.OPT.length) {
    config.OPT = ["CHAV", "LATIN", "GREC"];
    config.OPT_AVAILABLE = config.OPT;
    Logger.log("Utilisation des OPT par défaut: CHAV,LATIN,GREC");
  }

  // Garantir que les listes AVAILABLE sont remplies
  if (!config.LV2_AVAILABLE || !config.LV2_AVAILABLE.length) {
    config.LV2_AVAILABLE = config.LV2;
  }
  if (!config.OPT_AVAILABLE || !config.OPT_AVAILABLE.length) {
    config.OPT_AVAILABLE = config.OPT;
  }

  Logger.log(`Configuration finale → NIVEAU=${config.NIVEAU || 'Non défini'} | LV2=[${config.LV2_AVAILABLE.join(',')}] | OPT=[${config.OPT_AVAILABLE.join(',')}]`);

  /* ------------------------------------------------------------------
   * 2️⃣  Récupération des feuilles à traiter -------------------------
   * ----------------------------------------------------------------*/
  let sourceSheets = [];
  try {
    sourceSheets = getSourceSheets();
    Logger.log(`Feuilles sources trouvées : ${sourceSheets.length}`);
  } catch (e) {
    // En cas d'erreur, tenter une détection basique des onglets sources
    Logger.log(`Erreur getSourceSheets : ${e.message}, tentative de détection basique...`);
    try {
      // Recherche basique des onglets qui pourraient être des sources (ex. 6°1, 5°2, etc.)
      ss.getSheets().forEach(sheet => {
        const sheetName = sheet.getName();
        if (!["_CONFIG", "_STRUCTURE", "CONSOLIDATION"].includes(sheetName) && 
            !sheetName.startsWith("_") && 
            sheet.getLastRow() > 1) {
          sourceSheets.push(sheet);
        }
      });
    } catch (err) {
      Logger.log(`Erreur détection basique : ${err.message}`);
    }
  }

  const sheetsToProcess = [...sourceSheets];
  const consolidationSheetName = "CONSOLIDATION";
  try {
    // Récupérer le nom de l'onglet CONSOLIDATION depuis CONFIG si disponible
    if (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.CONSOLIDATION) {
      consolidationSheetName = CONFIG.SHEETS.CONSOLIDATION;
    }
  } catch (e) {
    Logger.log("INFO: CONFIG.SHEETS.CONSOLIDATION non disponible, utilisation de 'CONSOLIDATION'");
  }

  const consolidationSheet = ss.getSheetByName(consolidationSheetName);
  if (consolidationSheet) {
    sheetsToProcess.push(consolidationSheet);
    Logger.log(`Onglet ${consolidationSheetName} ajouté à la liste.`);
  } else {
    Logger.log(`WARN : Onglet ${consolidationSheetName} non trouvé.`);
  }

  if (!sheetsToProcess.length) {
    Logger.log("Aucune feuille à traiter.");
    ui.alert("Aucune feuille à traiter n'a été trouvée.");
    return;
  }

  Logger.log(`Application des listes et formatages sur : ${sheetsToProcess.map(s => s.getName()).join(', ')}`);

  /* ------------------------------------------------------------------
   * 3️⃣  Préparation des listes et règles de validation --------------
   * ----------------------------------------------------------------*/
  const ruleSEXE = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'M', 'F'], true)
    .setAllowInvalid(false)
    .build();

  const lv2List = ['', ...config.LV2_AVAILABLE];
  const ruleLV2 = lv2List.length > 1 ? 
                   SpreadsheetApp.newDataValidation()
                   .requireValueInList(lv2List, true)
                   .setAllowInvalid(false)
                   .build() : null;
  
  const optList = ['', ...config.OPT_AVAILABLE];
  const ruleOPT = optList.length > 1 ? 
                   SpreadsheetApp.newDataValidation()
                   .requireValueInList(optList, true)
                   .setAllowInvalid(false)
                   .build() : null;
  
  const ruleCRIT = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', '1', '2', '3', '4'], true)
    .setAllowInvalid(false)
    .build();
  
  const ruleDISPO = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'ULIS', 'GEVASCO', 'PAP', 'PPRE', 'UPE2A', 'Autres'], true)
    .setAllowInvalid(true)
    .build();

  // Création de palettes de couleurs dynamiques pour LV2 et OPT
  const generateColorMap = (options, baseColors) => {
    const colorMap = {};
    options.forEach((opt, index) => {
      if (baseColors[opt]) {
        colorMap[opt] = baseColors[opt];
      } else {
        // Couleurs par défaut si non définies spécifiquement
        const defaultPalette = ['#d6eaf8', '#ebf5fb', '#e8f8f5', '#fef9e7', '#f4ecf7', '#eaeded'];
        colorMap[opt] = defaultPalette[index % defaultPalette.length];
      }
    });
    return colorMap;
  };

  // Couleurs de base pour les options connues
  const lv2BaseColors = { ESP:'#eb984e', ITA:'#73c6b6', ALL:'#aed6f1', CHI:'#f7dc6f' };
  const optBaseColors = { CHAV:'#d7bde2', LATIN:'#f9e79f', GREC:'#abebc6', LLCA:'#f5cba7' };
  
  // Générer les palettes complètes
  const lv2Colors = generateColorMap(config.LV2_AVAILABLE, lv2BaseColors);
  const optColors = generateColorMap(config.OPT_AVAILABLE, optBaseColors);

  /* ------------------------------------------------------------------
   * 4️⃣  Boucle sur chaque feuille ------------------------------------
   * ----------------------------------------------------------------*/
  sheetsToProcess.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`--- Traitement Onglet : ${sheetName} ---`);
    try {
      const maxRows = sheet.getMaxRows();
      const lastCol = sheet.getLastColumn();
      if (maxRows <= 1 || lastCol === 0) {
        Logger.log(" → Onglet vide ou sans données, ignoré.");
        return; // continue
      }

      const numRows = maxRows - 1;
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      Logger.log(`   En‑têtes : [${headers.join(', ')}]`);
      const findCol = name => getColumnIndexByName(sheet, name);

      // --------------------------------------------------------------
      // 4.1  Validation + MFC (on nettoie d'abord)
      // --------------------------------------------------------------
      sheet.clearConditionalFormatRules();
      const cfrules = [];
      let rulesAppliedCount = 0;

      // SEXE ---------------------------------------------------------
      const colSEXE = findCol('SEXE');
      if (colSEXE) {
        const range = sheet.getRange(2, colSEXE, numRows, 1);
        range.setDataValidation(ruleSEXE);
        sheet.setColumnWidth(colSEXE, 60);
        const colorM = '#85c1e9', colorF = '#f5b7b1';
        cfrules.push(
          SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('M').setBackground(colorM).setBold(true).setRanges([range]).build(),
          SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('F').setBackground(colorF).setBold(true).setRanges([range]).build()
        );
        rulesAppliedCount += 2;
      }

      // LV2 ----------------------------------------------------------
      const colLV2 = findCol('LV2');
      if (colLV2) {
        const range = sheet.getRange(2, colLV2, numRows, 1);
        ruleLV2 ? range.setDataValidation(ruleLV2) : range.clearDataValidations();
        
        // Appliquer le formatage conditionnel pour chaque LV2
        config.LV2_AVAILABLE.forEach(lv2 => {
          const bg = lv2Colors[lv2] || '#d6eaf8';
          cfrules.push(
            SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(lv2).setBackground(bg).setBold(true).setRanges([range]).build()
          );
          rulesAppliedCount++;
        });
      }

      // OPT ----------------------------------------------------------
      const colOPT = findCol('OPT');
      if (colOPT) {
        const range = sheet.getRange(2, colOPT, numRows, 1);
        ruleOPT ? range.setDataValidation(ruleOPT) : range.clearDataValidations();
        
        // Appliquer le formatage conditionnel pour chaque OPT
        config.OPT_AVAILABLE.forEach(opt => {
          const bg = optColors[opt] || '#d6eaf8';
          cfrules.push(
            SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(opt).setBackground(bg).setBold(true).setRanges([range]).build()
          );
          rulesAppliedCount++;
        });
      }

      // Critères -----------------------------------------------------
      ['COM','TRA','PART','ABS'].forEach(colName => {
        const idx = findCol(colName);
        if (idx) {
          const range = sheet.getRange(2, idx, numRows, 1);
          range.setDataValidation(ruleCRIT);
          const critColors = [
            {v:'1', bg:'#e74c3c', fc:'#ffffff'},
            {v:'2', bg:'#f9e79f', fc:'#000000'},
            {v:'3', bg:'#d5f5e3', fc:'#000000'},
            {v:'4', bg:'#1e8449', fc:'#ffffff'}
          ];
          critColors.forEach(c => {
            cfrules.push(
              SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(c.v).setBackground(c.bg).setFontColor(c.fc).setBold(true).setRanges([range]).build()
            );
            rulesAppliedCount++;
          });
        }
      });

      // DISPO --------------------------------------------------------
      const colDISPO = findCol('DISPO');
      if (colDISPO) {
        const range = sheet.getRange(2, colDISPO, numRows, 1);
        range.setDataValidation(ruleDISPO);
        const dispoColors = { ULIS:'#f4f6f7', GEVASCO:'#d6dbdf', PAP:'#d5f5e3', PPRE:'#ebf5fb', UPE2A:'#fdebd0' };
        Object.entries(dispoColors).forEach(([val, bg]) => {
          cfrules.push(
            SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(val).setBackground(bg).setBold(true).setRanges([range]).build()
          );
          rulesAppliedCount++;
        });
      }

      // Appliquer les règles MFC
      if (cfrules.length) sheet.setConditionalFormatRules(cfrules);
      Logger.log(`   ${rulesAppliedCount} règles de formatage appliquées sur ${sheetName}.`);

      // --------------------------------------------------------------
      // 4.2  Suppression de la gestion des formules NOM_PRENOM --------
      // --------------------------------------------------------------
      // Cette section a été nettoyée pour éviter les problèmes avec les formules
      // La génération des valeurs NOM_PRENOM et ID_ELEVE est maintenant gérée
      // par la fonction genererNomPrenomEtID() qui crée des valeurs directes
      // plutôt que des formules

      // Pour générer NOM_PRENOM et ID_ELEVE, utilisez l'option "Générer NOM_PRENOM et ID_ELEVE"
      // dans le menu "Outils de Traitement"

      Logger.log("Note: Formules NOM_PRENOM et ID_ELEVE non gérées - utilisez le menu dédié à la place");


      // --------------------------------------------------------------
      // 4.3  Largeurs de colonne précises -----------------------------
      // --------------------------------------------------------------
      ajusterLargeurColonnes(sheet);

      // --------------------------------------------------------------
      // 4.4  Figer la première ligne ---------------------------------
      // --------------------------------------------------------------
      sheet.setFrozenRows(1);

    } catch (err) {
      Logger.log(`ERREUR dans ${sheetName} : ${err} | ${err.stack}`);
      ui.alert(`Erreur sur l'onglet ${sheetName} : ${err.message}`);
    }
  });

  Logger.log("====== Fin ajouterListesDeroulantes (version universelle) ======");
  try { Utils.logAction("Ajout/MàJ Listes & Formatages (version universelle)"); } catch (_) {}
}

/**********************************************************************
 * UTILITAIRES                                                        *
 *********************************************************************/

/**
 * Convertit un numéro de colonne en lettre (1 → A, 28 → AB …)
 */
function columnToLetter(column) {
  let letter = '';
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Renvoie l'index (1‑based) d'une colonne à partir de son en‑tête
 */
function getColumnIndexByName(sheet, columnName) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.indexOf(columnName) + 1;
  } catch (e) {
    Logger.log(`Erreur getColumnIndexByName(${columnName}) : ${e}`);
    return 0;
  }
}

/**
 * Ajuste les largeurs des colonnes selon un dictionnaire (px)
 */
function ajusterLargeurColonnes(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const widths = {
    ID_ELEVE: 100, NOM: 120, PRENOM: 120, NOM_PRENOM: 180,
    SEXE: 60, LV2: 55, OPT: 65,
    COM: 50, TRA: 50, PART: 50, ABS: 50,
    DISPO: 85, ASSO: 70, DISSO: 70, SOURCE: 60
  };
  headers.forEach((h, i) => {
    if (widths[h]) sheet.setColumnWidth(i + 1, widths[h]);
  });
}
