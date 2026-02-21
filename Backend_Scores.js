/**
 * ===================================================================
 * BACKEND_SCORES.JS - MODULE DE CALCUL DES SCORES ÉLÈVES
 * ===================================================================
 * Calcule 4 scores (ABS, COM, TRA, PART) à partir des exports Pronote
 * et les injecte dans les colonnes des onglets sources élèves.
 *
 * S'intègre dans l'architecture existante :
 * - Lit les exports Pronote depuis des onglets DATA_* temporaires
 * - Calcule les scores sur l'échelle 0-5 (compatible Backend_Eleves.js)
 * - Écrit les résultats dans les colonnes COM, TRA, PART, ABS
 *   des onglets sources (°1, °2, etc.)
 *
 * ARCHITECTURE DES ONGLETS PRONOTE :
 * - DATA_ABS       → Export Pronote des absences
 * - DATA_INCIDENTS  → Export Pronote des incidents/sanctions
 * - DATA_PUNITIONS  → Export Pronote des punitions
 * - DATA_NOTES      → Export Pronote des notes/moyennes
 *
 * @version 1.0.0
 * ===================================================================
 */

// =============================================================================
// CONFIGURATION DU MODULE SCORES
// =============================================================================

const SCORES_CONFIG = {
  // Échelle de l'app : 0-5 (Backend_Eleves.js utilise validateScore 0-5)
  // Les seuils ci-dessous produisent des scores de 1 à 4,
  // puis on les mappe sur 0-5 via mapScore_()
  SEUILS_ABS: {
    DJ: [
      { score: 4, min: 0, max: 5 },
      { score: 3, min: 6, max: 13 },
      { score: 2, min: 14, max: 25 },
      { score: 1, min: 26, max: 999 }
    ],
    NJ: [
      { score: 4, min: 0, max: 0 },
      { score: 3, min: 1, max: 2 },
      { score: 2, min: 3, max: 5 },
      { score: 1, min: 6, max: 999 }
    ],
    poidsDJ: 0.6,
    poidsNJ: 0.4
  },
  SEUILS_COM: [
    { score: 4, min: 0, max: 0 },
    { score: 3, min: 1, max: 5 },
    { score: 2, min: 6, max: 20 },
    { score: 1, min: 21, max: 999 }
  ],
  SEUILS_TRA: [
    { score: 4, min: 15, max: 20 },
    { score: 3, min: 12, max: 14.99 },
    { score: 2, min: 8, max: 11.99 },
    { score: 1, min: 0, max: 7.99 }
  ],
  SEUILS_PART: [
    { score: 4, min: 15, max: 20 },
    { score: 3, min: 12, max: 14.99 },
    { score: 2, min: 8, max: 11.99 },
    { score: 1, min: 0, max: 7.99 }
  ],
  // Pondérations par matière pour le score TRA
  // cols = indices de colonnes (0-indexed) dans l'export Pronote notes
  MATIERES_TRA: [
    { nom: 'Français', cols: [8], coeff: 4.5 },
    { nom: 'Maths', cols: [11], coeff: 3.5 },
    { nom: 'Histoire-Géo', cols: [9], coeff: 3.0 },
    { nom: 'Anglais', cols: [2], coeff: 3.0 },
    { nom: 'Espagnol/It.', cols: [12], coeff: 2.5 },
    { nom: 'EPS', cols: [6], coeff: 2.0 },
    { nom: 'Phys.-Chimie', cols: [19, 20], coeff: 1.5 },
    { nom: 'SVT', cols: [17, 18], coeff: 1.5 },
    { nom: 'Technologie', cols: [15, 16], coeff: 1.5 },
    { nom: 'Arts Pla.', cols: [5], coeff: 1.0 },
    { nom: 'Musique', cols: [7], coeff: 1.0 },
    { nom: 'Latin', cols: [21], coeff: 1.0 }
  ],
  // Colonnes de l'export notes pour le score PART (oral)
  COL_ORAL_ANG: 4,   // index 0-based
  COL_ORAL_LV2: 14   // index 0-based
};

// =============================================================================
// FONCTIONS SERVEUR V3 — Adaptateurs pour Console Pilotage V3
// =============================================================================

/**
 * Initialise les onglets DATA_* pour recevoir les exports Pronote.
 * Appelé depuis la Console V3 phase Scores.
 * @returns {Object} {success, message, tabs}
 */
function v3_initScoresSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var onglets = ['DATA_ABS', 'DATA_INCIDENTS', 'DATA_PUNITIONS', 'DATA_NOTES'];
    var created = [];

    onglets.forEach(function(nom) {
      if (!ss.getSheetByName(nom)) {
        ss.insertSheet(nom);
        created.push(nom);
      }
    });

    // ── En-têtes calqués sur le format export Pronote ──
    // Chaque tableau reproduit l'ordre exact des colonnes attendues par
    // calculerScoreABS_, calculerScoreCOM_, calculerScoreTRA_, calculerScorePART_

    var headers = {
      'DATA_ABS': {
        // 2 lignes d'en-tête Pronote, données à partir de la ligne 3
        row1: ['Nom', 'Classe', 'Nb retards', 'Nb retards NJ', 'H retard',
               'Nb abs', 'Nb abs NJ', 'H absence', 'H absence NJ',
               'DJ Bulletin', 'Justifiée'],
        instruction: 'Export Pronote ABSENCES — collez les données à partir de la ligne 3'
      },
      'DATA_INCIDENTS': {
        // 2 lignes d'en-tête Pronote, données à partir de la ligne 3
        row1: ['Classe', 'Nom', 'Prénom', 'Date', 'Objet',
               'Type', 'Protagoniste', 'Gravité'],
        instruction: 'Export Pronote INCIDENTS — collez les données à partir de la ligne 3'
      },
      'DATA_PUNITIONS': {
        // 1 seule ligne d'en-tête, données à partir de la ligne 2
        row1: ['Nb', 'Nom', 'Classe'],
        instruction: null // pas de ligne d'instruction séparée, 1 seul header row
      },
      'DATA_NOTES': {
        // 2 lignes d'en-tête Pronote, données à partir de la ligne 3
        // Les indices doivent correspondre à MATIERES_TRA et COL_ORAL_*
        row1: ['Nom', 'Classe',
               'Anglais', '(Réservé)', 'Oral Anglais',
               'Arts Pla.', 'EPS', 'Musique',
               'Français', 'Hist-Géo', '(Réservé)',
               'Maths', 'Espagnol/It.', '(Réservé)', 'Oral LV2',
               'Technologie', 'Techno(2)', 'SVT', 'SVT(2)',
               'Phys-Chimie', 'PhysCh(2)', 'Latin'],
        instruction: 'Export Pronote NOTES — collez les données à partir de la ligne 3'
      }
    };

    var headerBg = '#1a237e';
    var headerColor = '#ffffff';
    var instrBg = '#e8eaf6';
    var instrColor = '#283593';
    var requiredBg = '#fff9c4'; // jaune clair pour colonnes critiques

    // Indices critiques (0-based) lus par le moteur
    var criticalCols = {
      'DATA_ABS': [0, 1, 9, 10],
      'DATA_INCIDENTS': [0, 1, 7],
      'DATA_PUNITIONS': [0, 1, 2],
      'DATA_NOTES': [0, 1, 2, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 17, 19, 21]
    };

    for (var nom in headers) {
      var ws = ss.getSheetByName(nom);
      if (!ws) continue;
      // N'écraser que si l'onglet est vide ou ne contient que l'ancienne instruction
      if (ws.getLastRow() > 2) continue;

      var h = headers[nom];
      var cols = h.row1.length;

      if (nom === 'DATA_PUNITIONS') {
        // 1 seule ligne d'en-tête
        ws.getRange(1, 1, 1, cols).setValues([h.row1]);
        ws.getRange(1, 1, 1, cols)
          .setFontWeight('bold').setFontColor(headerColor)
          .setBackground(headerBg).setHorizontalAlignment('center');
      } else {
        // Ligne 1 = instruction, Ligne 2 = en-têtes
        ws.getRange(1, 1).setValue(h.instruction);
        ws.getRange(1, 1, 1, cols)
          .merge().setFontStyle('italic').setFontColor(instrColor)
          .setBackground(instrBg).setHorizontalAlignment('center');

        ws.getRange(2, 1, 1, cols).setValues([h.row1]);
        ws.getRange(2, 1, 1, cols)
          .setFontWeight('bold').setFontColor(headerColor)
          .setBackground(headerBg).setHorizontalAlignment('center');
      }

      // Surligner les colonnes critiques
      var critical = criticalCols[nom] || [];
      var headerRow = (nom === 'DATA_PUNITIONS') ? 1 : 2;
      for (var c = 0; c < critical.length; c++) {
        if (critical[c] < cols) {
          ws.getRange(headerRow, critical[c] + 1)
            .setBackground(requiredBg).setFontColor('#1a237e');
        }
      }

      // Figer la/les ligne(s) d'en-tête
      ws.setFrozenRows(nom === 'DATA_PUNITIONS' ? 1 : 2);
    }

    return {
      success: true,
      message: created.length > 0
        ? 'Onglets créés : ' + created.join(', ')
        : 'Tous les onglets DATA existent déjà.',
      tabs: onglets.map(function(nom) {
        var ws = ss.getSheetByName(nom);
        return {
          name: nom,
          rows: ws ? Math.max(0, ws.getLastRow() - 2) : 0,
          hasData: ws ? ws.getLastRow() > 2 : false
        };
      })
    };

  } catch (e) {
    Logger.log('Erreur v3_initScoresSheets: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Vérifie l'état des onglets DATA et retourne le statut.
 * @returns {Object} {success, tabs: [{name, rows, hasData}]}
 */
function v3_getScoresStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var onglets = ['DATA_ABS', 'DATA_INCIDENTS', 'DATA_PUNITIONS', 'DATA_NOTES'];

    return {
      success: true,
      tabs: onglets.map(function(nom) {
        var ws = ss.getSheetByName(nom);
        var rows = ws ? Math.max(0, ws.getLastRow() - 2) : 0;
        return {
          name: nom,
          rows: rows,
          hasData: rows > 0,
          exists: !!ws
        };
      })
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Calcule TOUS les scores et les injecte dans les onglets sources.
 * Point d'entrée principal depuis la Console V3.
 * @returns {Object} {success, results: {abs, com, tra, part}, injected}
 */
function v3_calculerTousScores() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Calculer chaque score
    var absResults = calculerScoreABS_(ss);
    var comResults = calculerScoreCOM_(ss);
    var traResults = calculerScoreTRA_(ss);
    var partResults = calculerScorePART_(ss);

    // Fusionner tous les résultats par nom d'élève
    var fusion = fusionnerScores_(absResults, comResults, traResults, partResults);

    // Injecter dans les onglets sources
    var injected = injecterScoresDansOngletsSources_(ss, fusion);

    return {
      success: true,
      results: {
        abs: { count: absResults.length, ok: true },
        com: { count: comResults.length, ok: true },
        tra: { count: traResults.length, ok: true },
        part: { count: partResults.length, ok: true }
      },
      injected: injected,
      totalEleves: Object.keys(fusion).length
    };

  } catch (e) {
    Logger.log('Erreur v3_calculerTousScores: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Calcule un seul type de score.
 * @param {string} type - 'ABS', 'COM', 'TRA' ou 'PART'
 * @returns {Object} {success, count, results}
 */
function v3_calculerScore(type) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var results;

    switch (type) {
      case 'ABS': results = calculerScoreABS_(ss); break;
      case 'COM': results = calculerScoreCOM_(ss); break;
      case 'TRA': results = calculerScoreTRA_(ss); break;
      case 'PART': results = calculerScorePART_(ss); break;
      default: return { success: false, error: 'Type de score inconnu: ' + type };
    }

    return {
      success: true,
      type: type,
      count: results.length,
      results: results.slice(0, 20) // Aperçu des 20 premiers
    };

  } catch (e) {
    Logger.log('Erreur v3_calculerScore(' + type + '): ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Récupère un aperçu des scores calculés (pour affichage dans la Console V3).
 * @returns {Object} {success, preview: [{nom, classe, abs, com, tra, part}]}
 */
function v3_getScoresPreview() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var absResults = calculerScoreABS_(ss);
    var comResults = calculerScoreCOM_(ss);
    var traResults = calculerScoreTRA_(ss);
    var partResults = calculerScorePART_(ss);

    var fusion = fusionnerScores_(absResults, comResults, traResults, partResults);

    var preview = [];
    for (var nom in fusion) {
      var e = fusion[nom];
      preview.push({
        nom: nom,
        classe: e.classe,
        abs: e.scoreABS !== null ? mapScore_(e.scoreABS) : null,
        com: e.scoreCOM !== null ? mapScore_(e.scoreCOM) : null,
        tra: e.scoreTRA !== null ? mapScore_(e.scoreTRA) : null,
        part: e.scorePART !== null ? mapScore_(e.scorePART) : null
      });
    }

    preview.sort(function(a, b) {
      return (a.classe || '').localeCompare(b.classe || '') || a.nom.localeCompare(b.nom);
    });

    return {
      success: true,
      totalEleves: preview.length,
      preview: preview.slice(0, 50) // 50 premiers pour l'aperçu
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// MODULE ABS — Score d'assiduité
// =============================================================================

function calculerScoreABS_(ss) {
  var wsData = ss.getSheetByName('DATA_ABS');
  if (!wsData || wsData.getLastRow() < 3) return [];

  var data = wsData.getDataRange().getValues();
  var eleves = {};
  var seuils = SCORES_CONFIG.SEUILS_ABS;

  for (var i = 2; i < data.length; i++) {
    var nom = String(data[i][0]).trim();
    if (!nom) continue;
    var classe = String(data[i][1]).trim();
    var djBulletin = parseNotePronote_(data[i][9]);
    var justifiee = String(data[i][10]).trim();

    if (!eleves[nom]) {
      eleves[nom] = { classe: classe, djTotal: 0, nonJustifiees: 0 };
    }
    if (djBulletin !== null) eleves[nom].djTotal += djBulletin;
    if (justifiee === 'Non') eleves[nom].nonJustifiees++;
  }

  var resultats = [];
  for (var nom in eleves) {
    var e = eleves[nom];
    var scoreDJ = attribuerScoreParSeuil_(e.djTotal, seuils.DJ);
    var scoreNJ = attribuerScoreParSeuil_(e.nonJustifiees, seuils.NJ);
    var scoreABS = Math.ceil(scoreDJ * seuils.poidsDJ + scoreNJ * seuils.poidsNJ);

    resultats.push({
      nom: nom, classe: e.classe,
      dj: Math.round(e.djTotal * 10) / 10,
      nj: e.nonJustifiees,
      scoreABS: scoreABS
    });
  }

  return resultats;
}

// =============================================================================
// MODULE COM — Score de comportement
// =============================================================================

function calculerScoreCOM_(ss) {
  var wsInc = ss.getSheetByName('DATA_INCIDENTS');
  var wsPun = ss.getSheetByName('DATA_PUNITIONS');

  if ((!wsInc || wsInc.getLastRow() < 3) && (!wsPun || wsPun.getLastRow() < 2)) return [];

  var seuils = SCORES_CONFIG.SEUILS_COM;

  // Lire les punitions
  var punitions = {};
  if (wsPun && wsPun.getLastRow() >= 2) {
    var dataPun = wsPun.getDataRange().getValues();
    for (var i = 1; i < dataPun.length; i++) {
      var nom = String(dataPun[i][1]).trim();
      if (!nom) continue;
      var nb = parseInt(dataPun[i][0]) || 0;
      var classe = String(dataPun[i][2]).trim();
      punitions[nom] = { nb: nb, classe: classe };
    }
  }

  // Lire les incidents
  var incidents = {};
  if (wsInc && wsInc.getLastRow() >= 3) {
    var dataInc = wsInc.getDataRange().getValues();
    for (var i = 2; i < dataInc.length; i++) {
      var nom = String(dataInc[i][1]).trim();
      if (!nom) continue;
      var classe = String(dataInc[i][0]).trim();
      var gravStr = String(dataInc[i][7]).trim();
      var grav = 1;
      if (gravStr && gravStr.indexOf('/') > -1) {
        grav = parseInt(gravStr.split('/')[0]) || 1;
      }

      if (!incidents[nom]) {
        incidents[nom] = { classe: '', nbInc: 0, ptsGrav: 0 };
      }
      if (classe) incidents[nom].classe = classe;
      incidents[nom].nbInc++;
      incidents[nom].ptsGrav += grav;
    }
  }

  // Fusionner punitions + incidents
  var tousNoms = {};
  for (var nom in punitions) tousNoms[nom] = true;
  for (var nom in incidents) tousNoms[nom] = true;

  var resultats = [];
  for (var nom in tousNoms) {
    var ptsPun = punitions[nom] ? punitions[nom].nb : 0;
    var ptsInc = incidents[nom] ? incidents[nom].ptsGrav * 3 : 0;
    var total = ptsPun + ptsInc;
    var classe = (punitions[nom] ? punitions[nom].classe : '') || (incidents[nom] ? incidents[nom].classe : '');
    var scoreCOM = attribuerScoreParSeuil_(total, seuils);

    resultats.push({ nom: nom, classe: classe, total: total, scoreCOM: scoreCOM });
  }

  return resultats;
}

// =============================================================================
// MODULE TRA — Score de travail (moyenne pondérée)
// =============================================================================

function calculerScoreTRA_(ss) {
  var wsData = ss.getSheetByName('DATA_NOTES');
  if (!wsData || wsData.getLastRow() < 3) return [];

  var seuils = SCORES_CONFIG.SEUILS_TRA;
  var matieres = SCORES_CONFIG.MATIERES_TRA;
  var data = wsData.getDataRange().getValues();
  var resultats = [];

  for (var i = 2; i < data.length; i++) {
    var nom = String(data[i][0]).trim();
    if (!nom) continue;
    var classe = String(data[i][1]).trim();

    var totalPts = 0;
    var totalCoeff = 0;

    for (var m = 0; m < matieres.length; m++) {
      var note = null;
      for (var c = 0; c < matieres[m].cols.length; c++) {
        var colIdx = matieres[m].cols[c];
        if (colIdx < data[i].length) {
          var n = parseNotePronote_(data[i][colIdx]);
          if (n !== null) { note = n; break; }
        }
      }
      if (note !== null) {
        totalPts += note * matieres[m].coeff;
        totalCoeff += matieres[m].coeff;
      }
    }

    var moyPond = totalCoeff > 0 ? Math.round(totalPts / totalCoeff * 100) / 100 : null;
    var scoreTRA = moyPond !== null ? attribuerScoreParSeuil_(moyPond, seuils) : null;

    resultats.push({ nom: nom, classe: classe, moyPond: moyPond, scoreTRA: scoreTRA });
  }

  return resultats;
}

// =============================================================================
// MODULE PART — Score de participation orale
// =============================================================================

function calculerScorePART_(ss) {
  var wsData = ss.getSheetByName('DATA_NOTES');
  if (!wsData || wsData.getLastRow() < 3) return [];

  var seuils = SCORES_CONFIG.SEUILS_PART;
  var data = wsData.getDataRange().getValues();
  var resultats = [];

  for (var i = 2; i < data.length; i++) {
    var nom = String(data[i][0]).trim();
    if (!nom) continue;
    var classe = String(data[i][1]).trim();

    var oralAng = parseNotePronote_(data[i][SCORES_CONFIG.COL_ORAL_ANG]);
    var oralLV2 = parseNotePronote_(data[i][SCORES_CONFIG.COL_ORAL_LV2]);

    var notes = [];
    if (oralAng !== null) notes.push(oralAng);
    if (oralLV2 !== null) notes.push(oralLV2);

    var moyOral = notes.length > 0
      ? Math.round(notes.reduce(function(a, b) { return a + b; }, 0) / notes.length * 100) / 100
      : null;
    var scorePART = moyOral !== null ? attribuerScoreParSeuil_(moyOral, seuils) : null;

    resultats.push({ nom: nom, classe: classe, moyOral: moyOral, scorePART: scorePART });
  }

  return resultats;
}

// =============================================================================
// FUSION ET INJECTION DANS LES ONGLETS SOURCES
// =============================================================================

/**
 * Fusionne les résultats des 4 modules en un seul objet par élève.
 */
function fusionnerScores_(absResults, comResults, traResults, partResults) {
  var fusion = {};

  absResults.forEach(function(r) {
    if (!fusion[r.nom]) fusion[r.nom] = { classe: r.classe };
    fusion[r.nom].scoreABS = r.scoreABS;
    fusion[r.nom].dj = r.dj;
  });

  comResults.forEach(function(r) {
    if (!fusion[r.nom]) fusion[r.nom] = { classe: r.classe };
    fusion[r.nom].scoreCOM = r.scoreCOM;
  });

  traResults.forEach(function(r) {
    if (!fusion[r.nom]) fusion[r.nom] = { classe: r.classe };
    fusion[r.nom].scoreTRA = r.scoreTRA;
    fusion[r.nom].moyPond = r.moyPond;
  });

  partResults.forEach(function(r) {
    if (!fusion[r.nom]) fusion[r.nom] = { classe: r.classe };
    fusion[r.nom].scorePART = r.scorePART;
  });

  return fusion;
}

/**
 * Injecte les scores calculés dans les colonnes COM, TRA, PART, ABSENCE
 * des onglets sources (ceux qui matchent le pattern °digit).
 * Fait le matching par NOM ou NOM_PRENOM.
 *
 * @returns {Object} {updated, notFound, total}
 */
function injecterScoresDansOngletsSources_(ss, fusion) {
  var sheets = ss.getSheets().filter(function(s) {
    return /.+°\d+$/.test(s.getName());
  });

  var totalUpdated = 0;
  var totalNotFound = 0;

  sheets.forEach(function(sheet) {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var headers = data[0];
    var idxNom = headers.indexOf('NOM');
    var idxPrenom = headers.indexOf('PRENOM');
    var idxNomPrenom = headers.indexOf('NOM_PRENOM');
    var idxCOM = headers.indexOf('COM');
    var idxTRA = headers.indexOf('TRA');
    var idxPART = headers.indexOf('PART');
    var idxABS = headers.indexOf('ABSENCE');
    if (idxABS === -1) idxABS = headers.indexOf('ABS');

    if (idxCOM === -1 && idxTRA === -1 && idxPART === -1 && idxABS === -1) return;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      // Construire les clés de matching possibles
      var nomPrenom = idxNomPrenom >= 0 ? String(row[idxNomPrenom]).trim() : '';
      var nom = idxNom >= 0 ? String(row[idxNom]).trim() : '';
      var prenom = idxPrenom >= 0 ? String(row[idxPrenom]).trim() : '';

      // Chercher dans la fusion par différentes clés
      var match = null;
      if (nomPrenom && fusion[nomPrenom]) {
        match = fusion[nomPrenom];
      } else if (nom && fusion[nom]) {
        match = fusion[nom];
      } else {
        // Essayer NOM Prénom combiné
        var fullName = (nom + ' ' + prenom).trim();
        if (fullName && fusion[fullName]) {
          match = fusion[fullName];
        }
      }

      if (match) {
        var rowNum = i + 1; // 1-indexed pour getRange
        if (idxCOM >= 0 && match.scoreCOM !== undefined) {
          sheet.getRange(rowNum, idxCOM + 1).setValue(mapScore_(match.scoreCOM));
        }
        if (idxTRA >= 0 && match.scoreTRA !== undefined) {
          sheet.getRange(rowNum, idxTRA + 1).setValue(mapScore_(match.scoreTRA));
        }
        if (idxPART >= 0 && match.scorePART !== undefined) {
          sheet.getRange(rowNum, idxPART + 1).setValue(mapScore_(match.scorePART));
        }
        if (idxABS >= 0 && match.scoreABS !== undefined) {
          sheet.getRange(rowNum, idxABS + 1).setValue(mapScore_(match.scoreABS));
        }
        totalUpdated++;
      } else {
        totalNotFound++;
      }
    }
  });

  SpreadsheetApp.flush();

  return {
    updated: totalUpdated,
    notFound: totalNotFound,
    total: totalUpdated + totalNotFound
  };
}

// =============================================================================
// FONCTIONS UTILITAIRES
// =============================================================================

/**
 * Mappe un score 1-4 (échelle Pronote) vers 0-5 (échelle app).
 * 1 → 1, 2 → 2.5, 3 → 3.5, 4 → 5
 */
function mapScore_(score14) {
  if (score14 === null || score14 === undefined) return 2.5; // défaut
  var map = { 1: 1, 2: 2.5, 3: 3.5, 4: 5 };
  return map[score14] !== undefined ? map[score14] : 2.5;
}

/**
 * Parse une note depuis un export Pronote.
 * Gère les virgules françaises, "Abs", "Disp", etc.
 */
function parseNotePronote_(val) {
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  if (s === '' || s === 'Abs' || s === 'Disp' || s === 'NE' || s === 'NN' || s === '—') return null;
  s = s.replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/**
 * Attribue un score basé sur une valeur et des seuils min/max.
 */
function attribuerScoreParSeuil_(valeur, seuils) {
  for (var i = 0; i < seuils.length; i++) {
    if (valeur >= seuils[i].min && valeur <= seuils[i].max) {
      return seuils[i].score;
    }
  }
  return 1;
}

// =============================================================================
// NETTOYAGE — Suppression des onglets DATA temporaires
// =============================================================================

/**
 * Supprime les onglets DATA_* après injection des scores.
 * @returns {Object} {success, deleted}
 */
function v3_cleanupScoresSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var onglets = ['DATA_ABS', 'DATA_INCIDENTS', 'DATA_PUNITIONS', 'DATA_NOTES'];
    var deleted = [];

    onglets.forEach(function(nom) {
      var ws = ss.getSheetByName(nom);
      if (ws) {
        ss.deleteSheet(ws);
        deleted.push(nom);
      }
    });

    return { success: true, deleted: deleted };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
