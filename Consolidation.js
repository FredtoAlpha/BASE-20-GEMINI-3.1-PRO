/**
 * Retourne la liste des onglets sources (feuilles nommees X°Y : 4°1, 5°2, etc.)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet[]}
 */
function getSourceSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(function(s) { return /.+°\d+$/.test(s.getName()); });
}

/**
 * Vérifie l'intégrité des données consolidées
 * - Vérifie que chaque élève a un ID unique
 * - Vérifie que les champs obligatoires sont remplis
 * - Ignore les colonnes G (OPT), L, M et N lors de la vérification
 */
function verifierDonnees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Vérifier que l'onglet CONSOLIDATION existe
  const consolidationSheet = ss.getSheetByName("CONSOLIDATION");
  if (!consolidationSheet) {
    Logger.log("❌ L'onglet CONSOLIDATION n'existe pas. Veuillez d'abord exécuter la consolidation.");
    return "Onglet CONSOLIDATION manquant";
  }

  // Vérifier d'abord les onglets sources
  const sourceSheets = getSourceSheets();
  if (sourceSheets.length === 0) {
    Logger.log("❌ Aucun onglet source trouvé. Veuillez vérifier votre structure.");
    return "Aucun onglet source introuvable";
  }

  // Liste des problèmes pour tous les onglets
  let problemesGlobaux = [];
  let totalEleves = 0;

  // Vérifier chaque onglet source
  for (const sheet of sourceSheets) {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      // Onglet vide ou juste l'en-tête, passer au suivant
      continue;
    }

    const headerRow = data[0];

    // Colonnes à vérifier (A, B, C, D, E, F, H, I, J, K et O)
    // A=ID_ELEVE, B=NOM, C=PRENOM, D=NOM_PRENOM, E=SEXE, F=LV2, H=COM, I=TRA, J=PART, K=ABS, O=?
    const requiredColumns = ["ID_ELEVE", "NOM", "PRENOM", "NOM_PRENOM", "SEXE", "LV2"];
    const additionalColumns = ["COM", "TRA", "PART", "ABS"];

    const indexes = {};
    requiredColumns.forEach(col => {
      indexes[col] = headerRow.indexOf(col);
    });

    additionalColumns.forEach(col => {
      indexes[col] = headerRow.indexOf(col);
    });

    // Vérifier si les colonnes requises existent
    const missingColumns = [];
    requiredColumns.forEach(col => {
      if (indexes[col] === -1) {
        missingColumns.push(col);
      }
    });

    if (missingColumns.length > 0) {
      Logger.log(`❌ Colonnes manquantes dans ${sheetName}: ${missingColumns.join(", ")}`);
      return `Colonnes manquantes dans ${sheetName}: ${missingColumns.join(", ")}`;
    }

    // Vérifier les données (ignorer l'en-tête)
    const problemes = [];
    const idsUtilises = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;

      // Ignorer les lignes vides
      if (!row[indexes.NOM] && !row[indexes.PRENOM]) {
        continue;
      }

      totalEleves++;

      // Vérifier ID
      if (!row[indexes.ID_ELEVE]) {
        problemes.push(`Ligne ${rowNum}: ID manquant pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
      } else if (idsUtilises[row[indexes.ID_ELEVE]]) {
        problemes.push(`Ligne ${rowNum}: ID en double "${row[indexes.ID_ELEVE]}" (déjà utilisé ligne ${idsUtilises[row[indexes.ID_ELEVE]]})`);
      } else {
        idsUtilises[row[indexes.ID_ELEVE]] = rowNum;
      }

      // Vérifier les champs obligatoires (NOM, PRENOM, SEXE, LV2)
      for (const col of ["NOM", "PRENOM", "SEXE", "LV2"]) {
        if (indexes[col] !== -1 && !row[indexes[col]]) {
          problemes.push(`Ligne ${rowNum}: "${col}" manquant pour ${row[indexes.NOM] || ""} ${row[indexes.PRENOM] || ""}`);
        }
      }

      // Vérifier que NOM_PRENOM est correctement formé (si présent)
      if (indexes.NOM_PRENOM !== -1 && row[indexes.NOM] && row[indexes.PRENOM]) {
        const expectedNomPrenom = `${row[indexes.NOM]} ${row[indexes.PRENOM]}`;
        if (row[indexes.NOM_PRENOM] !== expectedNomPrenom) {
          problemes.push(`Ligne ${rowNum}: NOM_PRENOM incorrect "${row[indexes.NOM_PRENOM]}" (devrait être "${expectedNomPrenom}")`);
        }
      }

      // Vérifier les critères (COM, TRA, PART, ABS) s'ils existent
      additionalColumns.forEach(col => {
        if (indexes[col] !== -1) {
          const valeur = row[indexes[col]];
          if (valeur === "") {
            problemes.push(`Ligne ${rowNum}: "${col}" manquant pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
          } else if (typeof valeur === 'number' && (valeur < 1 || valeur > 4)) {
            problemes.push(`Ligne ${rowNum}: "${col}" invalide (${valeur}) pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
          }
        }
      });

      // Nous n'effectuons pas de vérification sur la colonne G (OPT), L, M et N
    }

    // Ajouter les problèmes de cet onglet à la liste globale
    if (problemes.length > 0) {
      problemesGlobaux.push(`Onglet ${sheetName}:`);
      problemesGlobaux = problemesGlobaux.concat(problemes);
      problemesGlobaux.push(""); // Ligne vide entre les onglets
    }
  }

  // Vérifier également l'onglet CONSOLIDATION s'il existe (déjà vérifié au début)
  if (consolidationSheet) {
    const data = consolidationSheet.getDataRange().getValues();

    if (data.length > 1) {
      const headerRow = data[0];

      // Vérifier uniquement ID_ELEVE, NOM, PRENOM, SEXE, LV2 dans CONSOLIDATION
      const requiredColumns = ["ID_ELEVE", "NOM", "PRENOM", "SEXE", "LV2"];

      const indexes = {};
      requiredColumns.forEach(col => {
        indexes[col] = headerRow.indexOf(col);
      });

      // Vérifier si les colonnes requises existent
      const missingColumns = [];
      requiredColumns.forEach(col => {
        if (indexes[col] === -1) {
          missingColumns.push(col);
        }
      });

      if (missingColumns.length > 0) {
        Logger.log(`❌ Colonnes manquantes dans CONSOLIDATION: ${missingColumns.join(", ")}`);
        return `Colonnes manquantes dans CONSOLIDATION: ${missingColumns.join(", ")}`;
      }

      // Vérifier les données (ignorer l'en-tête)
      const problemes = [];
      const idsUtilises = {};
      const idsSuffixesTraites = new Set(); // Pour traquer les IDs avec suffixes déjà traités

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowNum = i + 1;

        // Ignorer les lignes vides
        if (!row[indexes.NOM] && !row[indexes.PRENOM]) {
          continue;
        }

        // Vérifier ID
        if (!row[indexes.ID_ELEVE]) {
          problemes.push(`Ligne ${rowNum}: ID manquant pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
        } else {
          const id = String(row[indexes.ID_ELEVE]);

          // Ignorer les vérifications de doublons pour les IDs avec suffixes (_1, _2, etc.)
          if (id.includes('_')) {
            // Extraire la partie base de l'ID (avant le _)
            const idBase = id.split('_')[0];

            // Si c'est la première fois qu'on voit cet ID avec suffixe, on l'accepte
            if (!idsSuffixesTraites.has(id)) {
              idsSuffixesTraites.add(id);
            } else {
              // Sinon, c'est un doublon d'un ID déjà avec suffixe
              problemes.push(`Ligne ${rowNum}: ID en double "${id}" (déjà utilisé ligne ${idsUtilises[id]})`);
            }
          } else if (idsUtilises[id]) {
            // ID en double (sans suffixe)
            problemes.push(`Ligne ${rowNum}: ID en double "${id}" (déjà utilisé ligne ${idsUtilises[id]})`);
          }

          // Enregistrer l'ID utilisé
          idsUtilises[id] = rowNum;
        }

        // Vérifier les champs obligatoires (NOM, PRENOM, SEXE, LV2)
        for (const col of ["NOM", "PRENOM", "SEXE", "LV2"]) {
          if (!row[indexes[col]]) {
            problemes.push(`Ligne ${rowNum}: "${col}" manquant pour ${row[indexes.NOM] || ""} ${row[indexes.PRENOM] || ""}`);
          }
        }
      }

      // Ajouter les problèmes de CONSOLIDATION à la liste globale
      if (problemes.length > 0) {
        problemesGlobaux.push("Onglet CONSOLIDATION:");
        problemesGlobaux = problemesGlobaux.concat(problemes);
      }
    }
  }

  // Afficher le résultat
  if (problemesGlobaux.length === 0) {
    Logger.log(`✅ Vérification terminée : Aucun problème détecté. Total d'élèves: ${totalEleves}`);
    return "Aucun problème détecté";
  } else {
    // Logger tous les problèmes
    Logger.log(`⚠️ ${problemesGlobaux.length} problème(s) détecté(s):`);
    problemesGlobaux.forEach(pb => Logger.log(pb));
    return `${problemesGlobaux.length} problème(s) détecté(s)`;
  }
}
/**
 * Consolide les données des onglets sources vers l'onglet CONSOLIDATION
 */
function consoliderDonnees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Flush pour s'assurer que les ecritures precedentes sont persistees
  SpreadsheetApp.flush();

  // Trouver les onglets sources avec le pattern universel
  const allSheets = ss.getSheets();
  const sourceSheets = allSheets.filter(s => /.+°\d+$/.test(s.getName()));

  Logger.log(`🔍 CONSOLIDATION: ${sourceSheets.length} onglets sources détectés: ${sourceSheets.map(s => s.getName()).join(', ')}`);

  if (sourceSheets.length === 0) {
    Logger.log("❌ Aucun onglet source trouvé");
    return "Aucun onglet source trouvé";
  }

  // Récupérer la liste des options valides depuis _CONFIG ou _STRUCTURE
  let optionsValides = [];
  try {
    // D'abord essayer _CONFIG
    const configSheet = ss.getSheetByName("_CONFIG");
    if (configSheet) {
      const data = configSheet.getDataRange().getValues();
      for (const row of data) {
        if (row[0] === "OPT" && row[1]) {
          optionsValides = String(row[1]).split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
          Logger.log(`Options récupérées depuis _CONFIG: ${optionsValides.join(',')}`);
          break;
        }
      }
    }

    // Si rien trouvé, essayer _STRUCTURE
    if (optionsValides.length === 0) {
      const structureSheet = ss.getSheetByName("_STRUCTURE");
      if (structureSheet) {
        const data = structureSheet.getDataRange().getValues();
        const optCol = data[0].indexOf("OPTIONS");
        if (optCol !== -1) {
          const optValues = data.slice(1)
            .map(row => row[optCol])
            .filter(val => val && typeof val === 'string')
            .map(val => val.includes("=") ? val.split("=")[0].trim() : val.trim())
            .filter(val => val);
          optionsValides = [...new Set(optValues)];
          Logger.log(`Options récupérées depuis _STRUCTURE: ${optionsValides.join(',')}`);
        }
      }
    }

    // Si toujours rien, utiliser les valeurs par défaut
    if (optionsValides.length === 0) {
      optionsValides = ["CHAV", "LATIN"];
      Logger.log("Utilisation des options par défaut: CHAV, LATIN");
    }
  } catch (e) {
    Logger.log(`Erreur récupération options: ${e.message}`);
    optionsValides = ["CHAV", "LATIN"]; // Valeurs par défaut en cas d'erreur
  }

  // Récupérer ou créer l'onglet CONSOLIDATION
  let consolidationSheet = ss.getSheetByName("CONSOLIDATION");
  if (!consolidationSheet) {
    Logger.log("⚠️ CONSOLIDATION n'existe pas, création...");
    consolidationSheet = ss.insertSheet("CONSOLIDATION");
  }

  // Nettoyer CONSOLIDATION
  Logger.log("🗑️ Nettoyage de CONSOLIDATION...");
  consolidationSheet.clear();

  // Creer l'en-tete depuis le premier onglet source (ecrire les headers BRUTS)
  const firstSource = sourceSheets[0];
  const firstHeaders = firstSource.getRange(1, 1, 1, firstSource.getLastColumn()).getValues()[0];
  // Ecrire les headers bruts dans CONSOLIDATION (preserve la casse d'origine)
  consolidationSheet.getRange(1, 1, 1, firstHeaders.length).setValues([firstHeaders]);

  // Appliquer le formatage à l'en-tête
  const headerRange = consolidationSheet.getRange(1, 1, 1, firstHeaders.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a5568');
  headerRange.setFontColor('#ffffff');

  // NORMALISER les en-tetes pour comparaison robuste (trim + uppercase)
  const headers = firstHeaders.map(h => String(h || '').trim().toUpperCase());
  const idIndex = headers.indexOf("ID_ELEVE");
  const sourceIndex = headers.indexOf("SOURCE");
  const optIndex = headers.indexOf("OPT");

  Logger.log(`📋 En-têtes normalisés: ${headers.join(', ')}`);

  if (idIndex === -1) {
    Logger.log("❌ Colonne ID_ELEVE manquante dans les sources. Headers: " + headers.join(', '));
    return "Colonne ID_ELEVE manquante";
  }

  // Collecter d'abord toutes les données
  const toutesLesDonnees = [];
  const idsUtilises = new Set();

  for (const sheet of sourceSheets) {
    const sheetName = sheet.getName();
    const lastRowSource = Math.max(sheet.getLastRow(), 1);
    if (lastRowSource <= 1) continue; // Onglet vide

    // Lire les en-têtes de CETTE feuille source spécifique - NORMALISEES
    const sourceHeadersRaw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sourceHeaders = sourceHeadersRaw.map(h => String(h || '').trim().toUpperCase());
    const sourceData = sheet.getRange(2, 1, lastRowSource - 1, sheet.getLastColumn()).getValues();

    Logger.log(`  Source ${sheetName}: ${sourceData.length} lignes, headers: [${sourceHeaders.join(' | ')}]`);

    // Créer un mapping des colonnes pour cette feuille (comparaison normalisée)
    const colMap = {};
    headers.forEach((h, destIdx) => {
      const srcIdx = sourceHeaders.indexOf(h);
      if (srcIdx !== -1) {
        colMap[destIdx] = srcIdx;
      }
    });

    // Filtrer et mapper les données
    sourceData.forEach(sourceRow => {
      // Vérifier si la ligne est vide (basé sur NOM/PRENOM dans la source)
      const srcNomIdx = sourceHeaders.indexOf("NOM");
      const srcPrenomIdx = sourceHeaders.indexOf("PRENOM");

      if (srcNomIdx !== -1 && srcPrenomIdx !== -1) {
        var nomVal = String(sourceRow[srcNomIdx] || '').trim();
        var prenomVal = String(sourceRow[srcPrenomIdx] || '').trim();
        if (!nomVal && !prenomVal) return; // Skip empty row
      }

      // Créer la nouvelle ligne consolidée
      const newRow = new Array(headers.length).fill("");

      // Remplir les colonnes mappées
      Object.keys(colMap).forEach(destIdx => {
        newRow[destIdx] = sourceRow[colMap[destIdx]];
      });

      // Logique spécifique (ID, Source, etc.)
      // ... (rest of the logic adapted to use newRow)

      // Si pas d'ID, en generer un (robuste: traite null/undefined/whitespace)
      var rawId = (newRow[idIndex] === null || newRow[idIndex] === undefined) ? '' : String(newRow[idIndex]).trim();
      if (!rawId) {
        newRow[idIndex] = `${sheetName}${(toutesLesDonnees.length + 1).toString().padStart(3, '0')}`;
      }

      // Assigner la source
      if (sourceIndex !== -1) newRow[sourceIndex] = sheetName;

      // Generer NOM_PRENOM si manquant (utilise indices normalises)
      const nomIndex = headers.indexOf("NOM");
      const prenomIndex = headers.indexOf("PRENOM");
      const nomPrenomIndex = headers.indexOf("NOM_PRENOM");

      if (nomIndex !== -1 && prenomIndex !== -1 && nomPrenomIndex !== -1) {
        var npExisting = String(newRow[nomPrenomIndex] || '').trim();
        if (!npExisting && newRow[nomIndex] && newRow[prenomIndex]) {
          newRow[nomPrenomIndex] = `${String(newRow[nomIndex]).trim()} ${String(newRow[prenomIndex]).trim()}`;
        }
      }

      // Nettoyer OPT
      if (optIndex !== -1 && newRow[optIndex]) {
        var optVal = String(newRow[optIndex]).trim().toUpperCase();
        if (!optionsValides.includes(optVal)) {
          newRow[optIndex] = "";
        }
      }

      // Gestion des IDs uniques
      let idOriginal = String(newRow[idIndex]).trim();
      newRow[idIndex] = idOriginal;
      let compteur = 1;
      while (idsUtilises.has(newRow[idIndex])) {
        newRow[idIndex] = `${idOriginal}_${compteur}`;
        compteur++;
      }

      idsUtilises.add(newRow[idIndex]);
      toutesLesDonnees.push(newRow);
    });
  }

  // Écrire toutes les données dans CONSOLIDATION
  Logger.log(`📝 Écriture de ${toutesLesDonnees.length} lignes dans CONSOLIDATION...`);
  if (toutesLesDonnees.length > 0) {
    consolidationSheet.getRange(2, 1, toutesLesDonnees.length, headers.length).setValues(toutesLesDonnees);
    Logger.log(`✅ ${toutesLesDonnees.length} élèves écrits avec succès`);
  } else {
    Logger.log(`⚠️ Aucune donnée à consolider`);
    return "Aucune donnée à consolider";
  }

  // Formater et trier
  Logger.log(`🎨 Formatage et tri...`);
  if (toutesLesDonnees.length > 0) {
    // Créer un filtre
    consolidationSheet.getRange(1, 1, toutesLesDonnees.length + 1, headers.length).createFilter();

    // Trier par NOM, PRENOM
    const nomIndex = headers.indexOf("NOM") + 1; // +1 car getRange est 1-indexé
    const prenomIndex = headers.indexOf("PRENOM") + 1;
    if (nomIndex > 0 && prenomIndex > 0) {
      consolidationSheet.getRange(2, 1, toutesLesDonnees.length, headers.length)
        .sort([{ column: nomIndex, ascending: true }, { column: prenomIndex, ascending: true }]);
    }
  }

  // Mettre en forme pour faciliter la lecture
  consolidationSheet.setFrozenRows(1);

  // Pour les listes déroulantes
  try {
    if (typeof ajouterListesDeroulantes === 'function') {
      ajouterListesDeroulantes();
    }
  } catch (e) {
    Logger.log("⚠️ Fonction ajouterListesDeroulantes non disponible");
  }

  const message = `✅ Consolidation terminée : ${toutesLesDonnees.length} élèves consolidés depuis ${sourceSheets.length} sources`;
  Logger.log(`\n🎉 ${message}`);
  return message;
}
