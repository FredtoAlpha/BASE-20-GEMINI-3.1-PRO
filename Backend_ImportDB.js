/**
 * ===================================================================
 * BACKEND_IMPORTDB.JS - IMPORT BASE DE DONNEES PRONOTE
 * ===================================================================
 * Permet l'import d'un fichier CSV unique contenant toutes les donnees
 * eleves (identite + donnees brutes) depuis Pronote.
 * Remplit les onglets sources et calcule/injecte les scores
 * en une seule operation.
 *
 * COLONNES DU TEMPLATE :
 *   Identite : CLASSE, NOM, PRENOM, SEXE, LV2, OPT, DISPO
 *   Absences : DJ, NJ
 *   Comportement : NB_PUNITIONS, NB_INCIDENTS
 *   Notes : FRANCAIS, MATHS, HIST_GEO, ANGLAIS, LV2_MOY, EPS,
 *           PHYS_CHIMIE, SVT, TECHNO, ARTS_PLAST, MUSIQUE, LATIN
 *   Oral : ORAL_ANGLAIS, ORAL_LV2
 *
 * @version 1.0.0
 * ===================================================================
 */

// =============================================================================
// TEMPLATE CSV
// =============================================================================

/**
 * Genere le contenu CSV du template d'import Pronote.
 * @returns {Object} {success, csv, filename}
 */
function v3_generateImportTemplate() {
  try {
    var sep = ';';
    var headers = [
      'CLASSE', 'NOM', 'PRENOM', 'SEXE', 'LV2', 'OPT', 'DISPO',
      'DJ', 'NJ',
      'NB_PUNITIONS', 'NB_INCIDENTS',
      'FRANCAIS', 'MATHS', 'HIST_GEO', 'ANGLAIS', 'LV2_MOY',
      'EPS', 'PHYS_CHIMIE', 'SVT', 'TECHNO', 'ARTS_PLAST', 'MUSIQUE', 'LATIN',
      'ORAL_ANGLAIS', 'ORAL_LV2'
    ];

    var example1 = [
      '4\u00b01', 'DUPONT', 'Marie', 'F', 'ESP', 'LATIN', '',
      '3', '0',
      '0', '0',
      '14,5', '12', '13', '15', '14',
      '16', '11', '12', '13', '14', '15', '16',
      '14', '13'
    ];

    var example2 = [
      '4\u00b01', 'MARTIN', 'Lucas', 'M', 'ALL', '', '',
      '8', '2',
      '1', '0',
      '10', '9', '11', '12', '13',
      '14', '10', '11', '12', '13', '11', '',
      '13', '14'
    ];

    var lines = [
      headers.join(sep),
      example1.join(sep),
      example2.join(sep)
    ];

    return {
      success: true,
      csv: '\uFEFF' + lines.join('\n'),
      filename: 'modele_import_pronote.csv'
    };
  } catch (e) {
    Logger.log('Erreur v3_generateImportTemplate: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// IMPORT PRINCIPAL
// =============================================================================

/**
 * Importe les donnees CSV dans les onglets sources et calcule les scores.
 * @param {string[]} headers - En-tetes du CSV (normalises)
 * @param {Array[]} rows - Lignes de donnees (tableaux de valeurs)
 * @returns {Object} {success, summary}
 */
function v3_importDatabase(headers, rows) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log('=== IMPORT BASE PRONOTE ===');
    Logger.log('Headers recus: ' + headers.join(' | '));
    Logger.log('Nombre de lignes: ' + rows.length);

    // 1. Mapper les colonnes du CSV
    var colMap = mapImportColumns_(headers);
    Logger.log('Colonnes detectees: ' + JSON.stringify(colMap));

    if (colMap.CLASSE === -1) {
      throw new Error('Colonne CLASSE introuvable dans le fichier. Colonnes disponibles: ' + headers.join(', '));
    }
    if (colMap.NOM === -1) {
      throw new Error('Colonne NOM introuvable dans le fichier. Colonnes disponibles: ' + headers.join(', '));
    }

    // 2. Detecter les colonnes de notes (pour TRA et PART)
    var gradeColumns = detectGradeColumns_(headers);
    var oralColumns = detectOralColumns_(headers);
    Logger.log('Matieres detectees: ' + gradeColumns.map(function(g) { return g.name; }).join(', '));
    Logger.log('Colonnes oral detectees: ' + oralColumns.length);

    // 3. Grouper les eleves par classe
    var classeGroups = {};
    var totalStudents = 0;
    var skippedRows = 0;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var classe = String(row[colMap.CLASSE] || '').trim();
      var nom = String(row[colMap.NOM] || '').trim();
      if (!classe || !nom) {
        skippedRows++;
        continue;
      }

      if (!classeGroups[classe]) classeGroups[classe] = [];
      classeGroups[classe].push(row);
      totalStudents++;
    }

    var classesProcessed = Object.keys(classeGroups);
    Logger.log('Classes trouvees: ' + classesProcessed.join(', '));
    Logger.log('Total eleves: ' + totalStudents + ' (ignores: ' + skippedRows + ')');

    // 4. Headers des onglets sources
    var sourceHeaders = [
      'ID_ELEVE', 'NOM', 'PRENOM', 'NOM_PRENOM', 'SEXE', 'LV2', 'OPT',
      'COM', 'TRA', 'PART', 'ABS', 'DISPO', 'ASSO', 'DISSO', 'SOURCE'
    ];

    // 5. Preparer les regles de validation (listes deroulantes 1-4)
    var ruleCRIT = SpreadsheetApp.newDataValidation()
      .requireValueInList(['', '1', '2', '3', '4'], true)
      .setAllowInvalid(false)
      .build();

    var scoresInjected = 0;
    var sheetsCreated = 0;
    var sheetsUpdated = 0;

    // 6. Ecrire dans chaque onglet source
    for (var classeName in classeGroups) {
      var students = classeGroups[classeName];
      var sheet = ss.getSheetByName(classeName);

      if (!sheet) {
        sheet = ss.insertSheet(classeName);
        sheetsCreated++;
        Logger.log('Onglet cree: ' + classeName);
      } else {
        sheet.clear();
        sheetsUpdated++;
        Logger.log('Onglet mis a jour: ' + classeName);
      }

      // Headers
      sheet.getRange(1, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
      sheet.getRange(1, 1, 1, sourceHeaders.length)
        .setFontWeight('bold')
        .setBackground('#d9ead3')
        .setFontSize(10);
      sheet.setFrozenRows(1);

      // Construire les lignes eleves
      var studentRows = [];
      for (var s = 0; s < students.length; s++) {
        var sRow = students[s];
        var nomVal = String(sRow[colMap.NOM] || '').trim();
        var prenomVal = colMap.PRENOM >= 0 ? String(sRow[colMap.PRENOM] || '').trim() : '';
        var nomPrenom = nomVal + (prenomVal ? ' ' + prenomVal : '');
        var sexe = colMap.SEXE >= 0 ? String(sRow[colMap.SEXE] || '').trim().toUpperCase() : '';
        var lv2 = colMap.LV2 >= 0 ? String(sRow[colMap.LV2] || '').trim().toUpperCase() : '';
        var opt = colMap.OPT >= 0 ? String(sRow[colMap.OPT] || '').trim().toUpperCase() : '';
        var dispo = colMap.DISPO >= 0 ? String(sRow[colMap.DISPO] || '').trim().toUpperCase() : '';

        // Calculer les scores
        var scoreABS = calcImportScoreABS_(sRow, colMap);
        var scoreCOM = calcImportScoreCOM_(sRow, colMap);
        var scoreTRA = calcImportScoreTRA_(sRow, gradeColumns);
        var scorePART = calcImportScorePART_(sRow, oralColumns);

        if (scoreABS !== null || scoreCOM !== null || scoreTRA !== null || scorePART !== null) {
          scoresInjected++;
        }

        // ID auto-genere
        var id = classeName + String(s + 1).padStart(3, '0');

        studentRows.push([
          id,
          nomVal,
          prenomVal,
          nomPrenom,
          sexe,
          lv2,
          opt,
          scoreCOM !== null ? String(scoreCOM) : '',
          scoreTRA !== null ? String(scoreTRA) : '',
          scorePART !== null ? String(scorePART) : '',
          scoreABS !== null ? String(scoreABS) : '',
          dispo,
          '', // ASSO
          '', // DISSO
          classeName // SOURCE
        ]);
      }

      // Ecrire les donnees
      if (studentRows.length > 0) {
        sheet.getRange(2, 1, studentRows.length, sourceHeaders.length).setValues(studentRows);

        // Appliquer les listes deroulantes 1-4 sur COM(8), TRA(9), PART(10), ABS(11)
        [8, 9, 10, 11].forEach(function(col) {
          sheet.getRange(2, col, studentRows.length, 1).setDataValidation(ruleCRIT);
        });
      }

      // Ajuster les largeurs de colonnes
      var widths = {
        1: 100, 2: 120, 3: 120, 4: 180,
        5: 60, 6: 55, 7: 65,
        8: 50, 9: 50, 10: 50, 11: 50,
        12: 85, 13: 70, 14: 70, 15: 60
      };
      for (var col in widths) {
        sheet.setColumnWidth(parseInt(col), widths[col]);
      }
    }

    // 7. Appliquer les listes deroulantes completes (formatage conditionnel etc.)
    try {
      ajouterListesDeroulantes();
      Logger.log('Listes deroulantes appliquees');
    } catch (e) {
      Logger.log('Info: ajouterListesDeroulantes non disponible ou erreur: ' + e.message);
    }

    // 8. Generer NOM_PRENOM + IDs + Consolidation
    try {
      genererNomPrenomEtID();
      Logger.log('NOM_PRENOM et IDs generes');
    } catch (e) {
      Logger.log('Info: genererNomPrenomEtID erreur: ' + e.message);
    }

    var consolResult = '';
    try {
      consolResult = consoliderDonnees();
      Logger.log('Consolidation: ' + consolResult);
    } catch (e) {
      Logger.log('Info: consoliderDonnees erreur: ' + e.message);
      consolResult = 'Consolidation non disponible';
    }

    Logger.log('=== IMPORT TERMINE ===');
    Logger.log('Eleves importes: ' + totalStudents);
    Logger.log('Scores injectes: ' + scoresInjected);
    Logger.log('Onglets crees: ' + sheetsCreated + ', mis a jour: ' + sheetsUpdated);

    return {
      success: true,
      summary: {
        totalStudents: totalStudents,
        skippedRows: skippedRows,
        classesProcessed: classesProcessed.length,
        classesList: classesProcessed,
        sheetsCreated: sheetsCreated,
        sheetsUpdated: sheetsUpdated,
        scoresInjected: scoresInjected,
        gradesDetected: gradeColumns.map(function(g) { return g.name; }),
        consolidation: consolResult
      }
    };

  } catch (e) {
    Logger.log('Erreur v3_importDatabase: ' + e.toString());
    Logger.log(e.stack);
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// DETECTION DES COLONNES
// =============================================================================

/**
 * Mappe les en-tetes du CSV aux indices de colonnes.
 * Utilise du pattern matching pour accepter les noms Pronote natifs
 * en plus des noms du template.
 * @param {string[]} headers - En-tetes normalises (UPPERCASE, trimmed)
 * @returns {Object} Map nom → indice (0-based, -1 si non trouve)
 */
function mapImportColumns_(headers) {
  var h = headers.map(function(s) { return String(s).trim().toUpperCase(); });

  var mapping = {
    CLASSE:        findImportCol_(h, ['CLASSE', 'CLASS']),
    NOM:           findImportCol_(h, ['^NOM$', '^NOM[^_]']),
    PRENOM:        findImportCol_(h, ['PRENOM', 'PR[EE\u00c9]NOM']),
    SEXE:          findImportCol_(h, ['SEXE', 'GENRE']),
    LV2:           findImportCol_(h, ['^LV2$', 'LANGUE']),
    OPT:           findImportCol_(h, ['^OPT$', 'OPTION']),
    DISPO:         findImportCol_(h, ['DISPO', 'DISPOSITIF']),
    DJ:            findImportCol_(h, ['^DJ$', 'DEMI.?JOURN']),
    NJ:            findImportCol_(h, ['^NJ$', 'NON.?JUST']),
    NB_PUNITIONS:  findImportCol_(h, ['NB.?PUN', 'PUNITION']),
    NB_INCIDENTS:  findImportCol_(h, ['NB.?INC', 'INCIDENT'])
  };

  // Fallback pour NOM : si la regex stricte ne matche pas, essayer 'NOM'
  if (mapping.NOM === -1) {
    for (var i = 0; i < h.length; i++) {
      if (h[i] === 'NOM') { mapping.NOM = i; break; }
    }
  }

  return mapping;
}

/**
 * Cherche l'indice de la premiere colonne matchant un pattern.
 * @param {string[]} headers
 * @param {string[]} patterns
 * @returns {number} indice 0-based, -1 si non trouve
 */
function findImportCol_(headers, patterns) {
  for (var p = 0; p < patterns.length; p++) {
    var re = new RegExp(patterns[p], 'i');
    for (var c = 0; c < headers.length; c++) {
      if (re.test(headers[c])) return c;
    }
  }
  return -1;
}

/**
 * Detecte les colonnes de notes pour le score TRA.
 * Combine les patterns du template ET les patterns Pronote natifs.
 * @param {string[]} headers
 * @returns {Array<{name, col, coeff}>}
 */
function detectGradeColumns_(headers) {
  var h = headers.map(function(s) { return String(s).trim().toUpperCase(); });

  var matieres = [
    { name: 'Francais',    patterns: ['FRANCAIS', 'FRANC', 'FRAN[C\u00c7]'],             coeff: 4.5 },
    { name: 'Maths',       patterns: ['MATHS', 'MATH'],                                   coeff: 3.5 },
    { name: 'Hist-Geo',    patterns: ['HIST.?GEO', 'HI.?GE', 'HIST.*G[EE\u00c9]O', 'HG'], coeff: 3.0 },
    { name: 'Anglais',     patterns: ['ANGLAIS', 'ANG.*MOY', 'AGL.*MOY'],                  coeff: 3.0 },
    { name: 'LV2 Moy',     patterns: ['LV2.?MOY', 'ESP.*MOY', 'ALL.*MOY', 'ITA.*MOY'],    coeff: 2.5 },
    { name: 'EPS',         patterns: ['^EPS$', '^EPS[^A-Z]'],                              coeff: 2.0 },
    { name: 'Phys-Chimie', patterns: ['PHYS.?CHIMIE', 'PH.?CH', 'SC.?PH'],                coeff: 1.5 },
    { name: 'SVT',         patterns: ['^SVT$', '^SVT[^A-Z]'],                              coeff: 1.5 },
    { name: 'Techno',      patterns: ['TECHNO'],                                            coeff: 1.5 },
    { name: 'Arts Plast',  patterns: ['ARTS.?PLAST', 'A.?PLA'],                            coeff: 1.0 },
    { name: 'Musique',     patterns: ['MUSIQUE', 'EDMUS', '^MUS$'],                        coeff: 1.0 },
    { name: 'Latin',       patterns: ['^LATIN$', '^LAT$', 'LCALA'],                        coeff: 1.0 }
  ];

  var resolved = [];
  for (var m = 0; m < matieres.length; m++) {
    var mat = matieres[m];
    var col = findImportCol_(h, mat.patterns);
    if (col >= 0) {
      resolved.push({ name: mat.name, col: col, coeff: mat.coeff });
    }
  }

  return resolved;
}

/**
 * Detecte les colonnes de notes orales pour le score PART.
 * @param {string[]} headers
 * @returns {Array<{name, col}>}
 */
function detectOralColumns_(headers) {
  var h = headers.map(function(s) { return String(s).trim().toUpperCase(); });

  var oraux = [
    { name: 'Oral Anglais', patterns: ['ORAL.?ANGLAIS', 'ORAL.?ANG', 'ANG.*ORAL', 'AGL.*ORAL'] },
    { name: 'Oral LV2',     patterns: ['ORAL.?LV2', 'ESP.*ORAL', 'ALL.*ORAL', 'ITA.*ORAL', 'ORAL.*LV2'] }
  ];

  var resolved = [];
  for (var o = 0; o < oraux.length; o++) {
    var oral = oraux[o];
    var col = findImportCol_(h, oral.patterns);
    if (col >= 0) {
      resolved.push({ name: oral.name, col: col });
    }
  }

  return resolved;
}

// =============================================================================
// CALCUL DES SCORES DEPUIS UNE LIGNE IMPORTEE
// =============================================================================

/**
 * Parse une valeur numerique depuis le CSV (gere virgule francaise).
 * @param {*} val
 * @returns {number|null}
 */
function parseImportNote_(val) {
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  if (s === '' || s === '-' || s === 'Abs' || s === 'Disp' || s === 'NE' || s === 'NN') return null;
  s = s.replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/**
 * Score ABS depuis une ligne CSV.
 * DJ = demi-journees d'absence, NJ = nombre d'absences non justifiees.
 */
function calcImportScoreABS_(row, colMap) {
  if (colMap.DJ === -1) return null;

  var dj = parseImportNote_(row[colMap.DJ]);
  if (dj === null) return null;

  var nj = colMap.NJ >= 0 ? parseImportNote_(row[colMap.NJ]) : 0;
  if (nj === null) nj = 0;

  var seuils = SCORES_CONFIG.SEUILS_ABS;
  var scoreDJ = attribuerScoreParSeuil_(dj, seuils.DJ);
  var scoreNJ = attribuerScoreParSeuil_(nj, seuils.NJ);

  return Math.ceil(scoreDJ * seuils.poidsDJ + scoreNJ * seuils.poidsNJ);
}

/**
 * Score COM depuis une ligne CSV.
 * NB_PUNITIONS * 1 + NB_INCIDENTS * 3 → seuils COM.
 */
function calcImportScoreCOM_(row, colMap) {
  var hasPun = colMap.NB_PUNITIONS >= 0;
  var hasInc = colMap.NB_INCIDENTS >= 0;
  if (!hasPun && !hasInc) return null;

  var nbPun = hasPun ? (parseImportNote_(row[colMap.NB_PUNITIONS]) || 0) : 0;
  var nbInc = hasInc ? (parseImportNote_(row[colMap.NB_INCIDENTS]) || 0) : 0;

  // Meme formule que calculerScoreCOM_ : punitions*1 + incidents*gravity*3
  // Dans le template, gravity = 1 par defaut (donnee agregee)
  var total = nbPun + (nbInc * 3);

  return attribuerScoreParSeuil_(total, SCORES_CONFIG.SEUILS_COM);
}

/**
 * Score TRA depuis une ligne CSV.
 * Moyenne ponderee des notes de matieres.
 */
function calcImportScoreTRA_(row, gradeColumns) {
  if (!gradeColumns || gradeColumns.length === 0) return null;

  var totalPts = 0;
  var totalCoeff = 0;

  for (var i = 0; i < gradeColumns.length; i++) {
    var gc = gradeColumns[i];
    var note = parseImportNote_(row[gc.col]);
    if (note !== null) {
      totalPts += note * gc.coeff;
      totalCoeff += gc.coeff;
    }
  }

  if (totalCoeff === 0) return null;

  var moyPond = Math.round(totalPts / totalCoeff * 100) / 100;
  return attribuerScoreParSeuil_(moyPond, SCORES_CONFIG.SEUILS_TRA);
}

/**
 * Score PART depuis une ligne CSV.
 * Moyenne des notes orales.
 */
function calcImportScorePART_(row, oralColumns) {
  if (!oralColumns || oralColumns.length === 0) return null;

  var notes = [];
  for (var i = 0; i < oralColumns.length; i++) {
    var note = parseImportNote_(row[oralColumns[i].col]);
    if (note !== null) notes.push(note);
  }

  if (notes.length === 0) return null;

  var moyOral = Math.round(
    notes.reduce(function(a, b) { return a + b; }, 0) / notes.length * 100
  ) / 100;

  return attribuerScoreParSeuil_(moyOral, SCORES_CONFIG.SEUILS_PART);
}
