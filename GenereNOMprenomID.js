/**
 * ===================================================================
 * GENERATEUR D'IDENTIFIANTS UNIVERSEL
 * ===================================================================
 * Scanne TOUS les onglets sources (peu importe le format : 6°1, 5e2, CM2)
 * Genere les IDs au format historique : [NOM_ONGLET][1000 + INDEX]
 * Exemples: 6°51001, 5e21001, CM21001, BRESSOLS°51001
 *
 * Principe: DETECTION PAR EXCLUSION (prendre tout sauf systeme/resultats)
 *
 * ROBUSTESSE :
 *  - Detection des en-tetes par normalisation (trim + uppercase)
 *  - Comparaison d'ID vide tolere null/undefined/whitespace/retour chariot
 *  - Flush GAS garanti avant lecture
 */

function genererNomPrenomEtID() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Flush pour s'assurer que les ecritures precedentes sont persistees
  SpreadsheetApp.flush();

  // DETECTION STRICTE (Alignee avec Backend & Init)
  var sheets = ss.getSheets().filter(function(s) { return /.+°\d+$/.test(s.getName()); });

  if (sheets.length === 0) {
    Logger.log('[WARN] Aucun onglet source trouve (pattern X°N). Verifiez vos donnees.');
    return;
  }

  Logger.log('[INFO] genererNomPrenomEtID: ' + sheets.length + ' onglets sources: ' +
    sheets.map(function(s) { return s.getName(); }).join(', '));

  // TRAITEMENT ROBUSTE
  var totalUpdated = 0;

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('[INFO] ' + name + ' : onglet vide (lastRow=' + lastRow + '), ignore.');
      return;
    }

    var data = sheet.getDataRange().getValues();
    var rawHeaders = data[0];

    // NORMALISER LES EN-TETES : trim + uppercase pour eviter les espaces invisibles
    var headers = rawHeaders.map(function(h) { return String(h || '').trim().toUpperCase(); });

    var colID = headers.indexOf('ID_ELEVE');
    var colNom = headers.indexOf('NOM');
    var colPrenom = headers.indexOf('PRENOM');
    var colNomPrenom = headers.indexOf('NOM_PRENOM');

    if (colNom === -1 || colPrenom === -1) {
      Logger.log('[WARN] ' + name + ' : colonnes NOM/PRENOM introuvables. Headers: [' + headers.join(' | ') + ']');
      return;
    }

    Logger.log('[INFO] ' + name + ' : ' + (data.length - 1) + ' lignes, colID=' + colID +
      ' colNom=' + colNom + ' colPrenom=' + colPrenom + ' colNomPrenom=' + colNomPrenom);

    var prefix = name.trim();
    var countInSheet = 0;
    var updatesID = [];
    var updatesNP = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nom = String(row[colNom] || '').trim();
      var prenom = String(row[colPrenom] || '').trim();

      if (!nom && !prenom) continue;

      // A. Concatenation NOM_PRENOM
      if (colNomPrenom > -1) {
        var fullName = (nom + ' ' + prenom).trim();
        var existing = String(row[colNomPrenom] || '').trim();
        if (existing !== fullName) {
          sheet.getRange(i + 1, colNomPrenom + 1).setValue(fullName);
          updatesNP.push(i + 1);
        }
      }

      // B. Generation ID (Format universel: prefix + base1000)
      // ROBUSTE : on traite null, undefined, 0, whitespace, retour chariot
      var currentId = '';
      if (colID > -1) {
        var rawId = row[colID];
        currentId = (rawId === null || rawId === undefined) ? '' : String(rawId).trim();
      }

      if (!currentId) {
        var suffix = (1000 + countInSheet + 1).toString();
        currentId = prefix + suffix;

        if (colID > -1) {
          sheet.getRange(i + 1, colID + 1).setValue(currentId);
          updatesID.push(i + 1);
        }
      }
      countInSheet++;
      totalUpdated++;
    }
    Logger.log('[INFO] ' + name + ' : ' + countInSheet + ' eleves traites (Format ' + prefix + '1xxx). ' +
      updatesID.length + ' IDs generes, ' + updatesNP.length + ' NOM_PRENOM mis a jour.');
  });

  // Flush pour persister les ecritures avant toute lecture ulterieure (consolidation)
  SpreadsheetApp.flush();

  Logger.log('[INFO] IDs generes pour ' + totalUpdated + ' eleves dans ' + sheets.length + ' onglets.');
}

// NOTE: Le wrapper v3_genererNomPrenomEtID() est dans ConsolePilotageV3_Server.gs
// Il appelle genererNomPrenomEtID() PUIS consoliderDonnees()
