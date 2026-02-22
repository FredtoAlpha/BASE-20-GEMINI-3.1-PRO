/**
 * ===================================================================
 * BACKEND_IMPORTDB.JS - MODULE IMPORT MULTI-PASTE PRONOTE
 * ===================================================================
 * Architecture multi-collage : l'utilisateur copie depuis Pronote
 * et colle dans la Console V3. Le systeme parse, calcule les scores
 * et peuple les onglets sources.
 *
 * 4 ETAPES DE COLLAGE :
 *   1. Liste Eleves    -> NOM, PRENOM, SEXE, LV2, OPT, CLASSE
 *   2. Notes/Moyennes  -> TRA + PART (1 collage par classe)
 *   3. Absences         -> ABS (recap avec justifications)
 *   4. Comportement     -> COM (punitions + observations)
 *
 * COMPILATION : fusionne toutes les donnees, calcule les scores 1-4,
 * peuple les onglets sources avec listes deroulantes.
 *
 * @version 2.0.0 - Architecture multi-paste
 * ===================================================================
 */

// =============================================================================
// UTILITAIRES DE PARSING
// =============================================================================

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
 * Parse une valeur numerique (gere virgule francaise, Abs, Disp, etc.)
 */
function parseNote_(val) {
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  if (s === '' || s === '-' || s === 'Abs' || s === 'Disp' || s === 'NE' || s === 'NN' || s === 'N.Not') return null;
  s = s.replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/**
 * Normalise un nom : MAJUSCULES, trim, supprime accents
 */
function normaliserNom_(s) {
  if (!s) return '';
  return String(s).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

/**
 * Cle de fusion NOM+PRENOM pour matcher les eleves entre les pastes
 */
function cleEleve_(nom, prenom) {
  return normaliserNom_(nom) + '|' + normaliserNom_(prenom);
}

/**
 * Parse le sexe Pronote : feminin -> F, masculin -> M
 */
function parseSexe_(val) {
  if (!val) return '';
  var s = String(val).trim();
  if (s === '\u2640' || s === 'F' || s === 'f') return 'F';
  if (s === '\u2642' || s === 'M' || s === 'm') return 'M';
  return s.toUpperCase();
}

/**
 * Parse "Toutes les options" de Pronote pour extraire LV2 et OPT
 * Ex: "ANGLAIS LV1, ESPAGNOL LV2, LATIN" -> {lv2: 'ESP', opt: 'LATIN'}
 */
function parseOptions_(optionsStr) {
  var result = { lv2: '', opt: '' };
  if (!optionsStr) return result;

  var parts = String(optionsStr).split(',').map(function(p) { return p.trim().toUpperCase(); });

  var languesLV2 = {
    'ESPAGNOL': 'ESP', 'ALLEMAND': 'ALL', 'ITALIEN': 'ITA',
    'CHINOIS': 'CHI', 'PORTUGAIS': 'PT', 'ARABE': 'ARA',
    'RUSSE': 'RUS', 'JAPONAIS': 'JAP'
  };

  var optionsConnues = ['LATIN', 'GREC', 'CHAV', 'CHINOIS', 'LCALA', 'LLCA', 'EURO'];

  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];

    // Detecter LV2 : "ESPAGNOL LV2"
    if (p.indexOf('LV2') >= 0) {
      for (var langue in languesLV2) {
        if (p.indexOf(langue) >= 0) {
          result.lv2 = languesLV2[langue];
          break;
        }
      }
      if (!result.lv2) {
        var mot = p.replace(/\s*LV2.*/, '').trim();
        if (languesLV2[mot]) result.lv2 = languesLV2[mot];
        else if (mot.length <= 4) result.lv2 = mot;
      }
      continue;
    }

    // Ignorer LV1
    if (p.indexOf('LV1') >= 0) continue;

    // Detecter option
    for (var j = 0; j < optionsConnues.length; j++) {
      if (p.indexOf(optionsConnues[j]) >= 0) {
        result.opt = optionsConnues[j];
        break;
      }
    }
  }

  return result;
}

// =============================================================================
// 1. PARSING LISTE ELEVES
// =============================================================================

/**
 * Parse le collage de la liste eleves Pronote.
 * Format attendu (TSV) :
 *   Nom | Prenom | Ne(e) le | S | Classe | Prj. d'acc. | Toutes les options
 *
 * @param {Array[]} rows - Tableau 2D (lignes x colonnes) du TSV parse cote client
 * @returns {Object} {success, eleves, count, classes}
 */
function v3_parseListeEleves(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes. Collez la liste eleves depuis Pronote.' };
    }

    var headerRow = 0;
    var maxText = 0;
    for (var r = 0; r < Math.min(rows.length, 3); r++) {
      var textCount = 0;
      for (var c = 0; c < rows[r].length; c++) {
        var v = String(rows[r][c] || '').trim();
        if (v && isNaN(v)) textCount++;
      }
      if (textCount > maxText) { maxText = textCount; headerRow = r; }
    }

    var headers = rows[headerRow].map(function(h) { return String(h || '').trim().toUpperCase(); });
    Logger.log('Liste Eleves - Headers: ' + headers.join(' | '));

    var colNom = findImportCol_(headers, ['^NOM$', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[E\u00c9]NOM', 'PRENOM']);
    var colSexe = findImportCol_(headers, ['^S$', '^S\\.$', 'SEXE']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colOptions = findImportCol_(headers, ['TOUTES.*OPT', 'OPTIONS']);

    if (colNom === -1) return { success: false, error: 'Colonne NOM introuvable. Headers: ' + headers.join(', ') };
    if (colClasse === -1) return { success: false, error: 'Colonne CLASSE introuvable. Headers: ' + headers.join(', ') };

    Logger.log('Colonnes: NOM=' + colNom + ' PRENOM=' + colPrenom + ' SEXE=' + colSexe + ' CLASSE=' + colClasse + ' OPT=' + colOptions);

    var eleves = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var row = rows[i];
      var nom = String(row[colNom] || '').trim();
      if (!nom) continue;

      var prenom = colPrenom >= 0 ? String(row[colPrenom] || '').trim() : '';
      var sexe = colSexe >= 0 ? parseSexe_(row[colSexe]) : '';
      var classe = String(row[colClasse] || '').trim();
      if (!classe) continue;

      var opts = { lv2: '', opt: '' };
      if (colOptions >= 0) {
        opts = parseOptions_(row[colOptions]);
      }

      eleves.push({
        nom: nom.toUpperCase(),
        prenom: prenom,
        sexe: sexe,
        classe: classe,
        lv2: opts.lv2,
        opt: opts.opt
      });
    }

    var classes = {};
    eleves.forEach(function(e) { classes[e.classe] = (classes[e.classe] || 0) + 1; });

    Logger.log('Liste Eleves: ' + eleves.length + ' eleves dans ' + Object.keys(classes).length + ' classes');

    return {
      success: true,
      eleves: eleves,
      count: eleves.length,
      classes: classes
    };

  } catch (e) {
    Logger.log('Erreur v3_parseListeEleves: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// 2. PARSING NOTES / MOYENNES (PAR CLASSE)
// =============================================================================

/**
 * Parse le collage des notes/moyennes Pronote pour UNE classe.
 * Gere les headers repetes (AGL1 x3 = ecrit, oral, moyenne).
 *
 * @param {Array[]} rows - Tableau 2D du TSV parse
 * @returns {Object} {success, notes, count, classe, headersDetected, oralDetected}
 */
function v3_parseNotesMoyennes(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes pour les notes.' };
    }

    var headerRow = 0;
    var maxText = 0;
    for (var r = 0; r < Math.min(rows.length, 3); r++) {
      var textCount = 0;
      for (var c = 0; c < rows[r].length; c++) {
        var v = String(rows[r][c] || '').trim();
        if (v && isNaN(v.replace(',', '.'))) textCount++;
      }
      if (textCount > maxText) { maxText = textCount; headerRow = r; }
    }

    var rawHeaders = rows[headerRow].map(function(h) { return String(h || '').trim().toUpperCase(); });
    Logger.log('Notes - Headers bruts: ' + rawHeaders.join(' | '));

    var colNom = findImportCol_(rawHeaders, ['^NOM$', 'NOM']);
    var colPrenom = findImportCol_(rawHeaders, ['PR[E\u00c9]NOM', 'PRENOM']);
    var colClasse = findImportCol_(rawHeaders, ['CLASSE']);

    // DETECTION POSITION-BASED POUR HEADERS REPETES
    var matieresConfig = [
      { id: 'FRANC', patterns: ['FRANC', 'FRAN'], coeff: 4.5 },
      { id: 'MATH',  patterns: ['MATH'], coeff: 3.5 },
      { id: 'HG',    patterns: ['HI.?GE', 'HG', 'H.G'], coeff: 3.0, useMoy: true },
      { id: 'ANG',   patterns: ['AGL1', 'ANG', 'ANGLAIS'], coeff: 3.0, hasOral: true },
      { id: 'LV2',   patterns: ['ESP2', 'ALL2', 'ITA2', 'ESP', 'ALL', 'ITA'], coeff: 2.5, hasOral: true },
      { id: 'EPS',   patterns: ['^EPS$', 'EPS'], coeff: 2.0 },
      { id: 'PHCH',  patterns: ['PH.?CH', 'SC.?PH', 'PHYS'], coeff: 1.5 },
      { id: 'SVT',   patterns: ['^SVT$', 'SVT'], coeff: 1.5 },
      { id: 'TECH',  patterns: ['TECHN'], coeff: 1.5 },
      { id: 'APLA',  patterns: ['A.?PLA', 'ARTS'], coeff: 1.0 },
      { id: 'MUS',   patterns: ['EDMUS', 'MUS'], coeff: 1.0 },
      { id: 'LAT',   patterns: ['^LAT', 'LCALA'], coeff: 1.0 }
    ];

    var gradeMap = {};
    var oralMap = {};
    var detectedSubjects = [];

    for (var m = 0; m < matieresConfig.length; m++) {
      var mat = matieresConfig[m];
      var matchedCols = [];

      for (var p = 0; p < mat.patterns.length; p++) {
        var re = new RegExp(mat.patterns[p], 'i');
        for (var ci = 0; ci < rawHeaders.length; ci++) {
          if (re.test(rawHeaders[ci]) && matchedCols.indexOf(ci) === -1) {
            matchedCols.push(ci);
          }
        }
        if (matchedCols.length > 0) break;
      }

      if (matchedCols.length === 0) continue;

      // Position-based: 3 cols = ecrit(0), oral(1), moyenne(2)
      // 2 cols = sous-matiere(0), moyenne(1)
      // 1 col = moyenne
      var colMoy, colOral;

      if (matchedCols.length >= 3) {
        colMoy = matchedCols[2];
        colOral = matchedCols[1];
      } else if (matchedCols.length === 2) {
        colMoy = matchedCols[1];
        colOral = null;
      } else {
        colMoy = matchedCols[0];
        colOral = null;
      }

      gradeMap[mat.id] = { col: colMoy, coeff: mat.coeff };
      detectedSubjects.push(mat.id + '(col' + colMoy + ')');

      if (mat.hasOral && colOral !== null) {
        oralMap[mat.id] = colOral;
      }
    }

    Logger.log('Matieres detectees: ' + detectedSubjects.join(', '));
    Logger.log('Oraux detectes: ' + JSON.stringify(oralMap));

    // PARSER LES LIGNES ELEVES
    var notes = [];
    var classeDetected = '';

    for (var i = headerRow + 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.length < 3) continue;

      var nom = '', prenom = '', classe = '';

      if (colNom >= 0) nom = String(row[colNom] || '').trim();
      if (colPrenom >= 0) prenom = String(row[colPrenom] || '').trim();
      if (colClasse >= 0) classe = String(row[colClasse] || '').trim();

      // Si pas de colonne NOM explicite, chercher dans les cellules
      if (!nom) {
        for (var ci2 = 0; ci2 < Math.min(row.length, 6); ci2++) {
          var cellVal = String(row[ci2] || '').trim();
          if (cellVal && cellVal.length >= 2 && isNaN(cellVal.replace(',', '.'))
              && cellVal !== '\u2642' && cellVal !== '\u2640' && cellVal !== 'M' && cellVal !== 'F') {
            var parts = cellVal.split(/\s+/);
            if (parts.length >= 2) {
              nom = parts[0];
              prenom = parts.slice(1).join(' ');
            } else {
              nom = cellVal;
            }
            break;
          }
        }
      }

      if (!nom) continue;
      if (classe) classeDetected = classe;

      var moyennes = {};
      for (var gid in gradeMap) {
        var note = parseNote_(row[gradeMap[gid].col]);
        if (note !== null) moyennes[gid] = note;
      }

      var oraux = {};
      for (var oid in oralMap) {
        var noteOral = parseNote_(row[oralMap[oid]]);
        if (noteOral !== null) oraux[oid] = noteOral;
      }

      notes.push({
        nom: nom.toUpperCase(),
        prenom: prenom,
        classe: classe || classeDetected,
        moyennes: moyennes,
        oraux: oraux
      });
    }

    Logger.log('Notes: ' + notes.length + ' eleves parses, classe=' + classeDetected);

    return {
      success: true,
      notes: notes,
      count: notes.length,
      classe: classeDetected,
      headersDetected: detectedSubjects,
      oralDetected: Object.keys(oralMap)
    };

  } catch (e) {
    Logger.log('Erreur v3_parseNotesMoyennes: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// 3. PARSING ABSENCES
// =============================================================================

function v3_parseAbsences(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes pour les absences.' };
    }

    var headerRow = 0;
    var maxText = 0;
    for (var r = 0; r < Math.min(rows.length, 3); r++) {
      var textCount = 0;
      for (var c = 0; c < rows[r].length; c++) {
        var v = String(rows[r][c] || '').trim();
        if (v && isNaN(v)) textCount++;
      }
      if (textCount > maxText) { maxText = textCount; headerRow = r; }
    }

    var headers = rows[headerRow].map(function(h) { return String(h || '').trim().toUpperCase(); });
    Logger.log('Absences - Headers: ' + headers.join(' | '));

    var colNom = findImportCol_(headers, ['^NOM$', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[E\u00c9]NOM', 'PRENOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colDJ = findImportCol_(headers, ['^DJ$', 'DEMI.?JOURN', 'NB.*ABS', 'TOTAL']);
    var colJustif = findImportCol_(headers, ['JUSTIFI', 'JUST\\.?']);
    var colSante = findImportCol_(headers, ['SANT[E\u00c9]', 'SANTE']);
    var colNJ = findImportCol_(headers, ['NON.?JUST', '^NJ$', 'INJUST']);

    // Format: recap (1 ligne/eleve) ou evenementiel (1 ligne/evenement)
    var colDate = findImportCol_(headers, ['DATE', 'DU', 'DEBUT']);
    var isRecap = colDate === -1;

    var absences;
    if (isRecap) {
      absences = [];
      for (var i = headerRow + 1; i < rows.length; i++) {
        var row = rows[i];
        var nom = colNom >= 0 ? String(row[colNom] || '').trim() : '';
        if (!nom) continue;
        var prenom = colPrenom >= 0 ? String(row[colPrenom] || '').trim() : '';
        var classe = colClasse >= 0 ? String(row[colClasse] || '').trim() : '';
        var dj = colDJ >= 0 ? (parseNote_(row[colDJ]) || 0) : 0;
        var nj = 0;
        if (colNJ >= 0) {
          nj = parseNote_(row[colNJ]) || 0;
        } else if (colJustif >= 0) {
          var justif = parseNote_(row[colJustif]) || 0;
          nj = Math.max(0, dj - justif);
        }
        absences.push({ nom: nom.toUpperCase(), prenom: prenom, classe: classe, dj: dj, nj: nj });
      }
    } else {
      // Evenementiel: agreger par eleve
      var perStudent = {};
      for (var i2 = headerRow + 1; i2 < rows.length; i2++) {
        var row2 = rows[i2];
        var nom2 = colNom >= 0 ? String(row2[colNom] || '').trim().toUpperCase() : '';
        if (!nom2) continue;
        var prenom2 = colPrenom >= 0 ? String(row2[colPrenom] || '').trim() : '';
        var classe2 = colClasse >= 0 ? String(row2[colClasse] || '').trim() : '';
        var cle = cleEleve_(nom2, prenom2);
        if (!perStudent[cle]) {
          perStudent[cle] = { nom: nom2, prenom: prenom2, classe: classe2, dj: 0, nj: 0 };
        }
        var djVal = colDJ >= 0 ? (parseNote_(row2[colDJ]) || 1) : 1;
        perStudent[cle].dj += djVal;
        var justifVal = colJustif >= 0 ? String(row2[colJustif] || '').trim().toUpperCase() : '';
        var isJustifie = justifVal === 'OUI' || justifVal === 'O' || justifVal === 'X' || justifVal === '1';
        if (!isJustifie) perStudent[cle].nj += djVal;
        if (classe2) perStudent[cle].classe = classe2;
      }
      absences = [];
      for (var key in perStudent) absences.push(perStudent[key]);
    }

    Logger.log('Absences: ' + absences.length + ' eleves parses (format ' + (isRecap ? 'recap' : 'evenements') + ')');

    return {
      success: true,
      absences: absences,
      count: absences.length,
      format: isRecap ? 'recap' : 'events'
    };

  } catch (e) {
    Logger.log('Erreur v3_parseAbsences: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// 4. PARSING COMPORTEMENT (PUNITIONS + OBSERVATIONS)
// =============================================================================

function v3_parsePunitions(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes pour les punitions.' };
    }

    var headerRow = 0;
    var maxText = 0;
    for (var r = 0; r < Math.min(rows.length, 3); r++) {
      var textCount = 0;
      for (var c = 0; c < rows[r].length; c++) {
        var v = String(rows[r][c] || '').trim();
        if (v && isNaN(v)) textCount++;
      }
      if (textCount > maxText) { maxText = textCount; headerRow = r; }
    }

    var headers = rows[headerRow].map(function(h) { return String(h || '').trim().toUpperCase(); });
    Logger.log('Punitions - Headers: ' + headers.join(' | '));

    var colNom = findImportCol_(headers, ['^NOM$', 'NOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colNb = findImportCol_(headers, ['^NB', 'NOMBRE', 'QT', 'QUANT', 'TOTAL', 'PUNITION']);

    if (colNb === -1) {
      for (var c = 0; c < headers.length; c++) {
        var testVal = rows.length > headerRow + 1 ? String(rows[headerRow + 1][c] || '').trim() : '';
        if (testVal && !isNaN(testVal)) { colNb = c; break; }
      }
    }

    var punitions = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var row = rows[i];
      var nom = '', prenom = '', nb = 0, classe = '';

      if (colNb >= 0) nb = parseNote_(row[colNb]) || 0;

      if (colNom >= 0) {
        var nomVal = String(row[colNom] || '').trim();
        if (!nomVal) continue;
        var parts = nomVal.split(/\s+/);
        if (parts.length >= 2) {
          nom = parts[0].toUpperCase();
          prenom = parts.slice(1).join(' ');
        } else {
          nom = nomVal.toUpperCase();
        }
      } else {
        for (var ci = 0; ci < row.length; ci++) {
          var cv = String(row[ci] || '').trim();
          if (cv && cv.length >= 2 && isNaN(cv)) {
            var parts2 = cv.split(/\s+/);
            nom = parts2[0].toUpperCase();
            prenom = parts2.length > 1 ? parts2.slice(1).join(' ') : '';
            break;
          }
        }
      }

      if (!nom) continue;
      if (colClasse >= 0) classe = String(row[colClasse] || '').trim();

      punitions.push({ nom: nom, prenom: prenom, classe: classe, nb: nb });
    }

    Logger.log('Punitions: ' + punitions.length + ' eleves parses');
    return { success: true, punitions: punitions, count: punitions.length };

  } catch (e) {
    Logger.log('Erreur v3_parsePunitions: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function v3_parseObservations(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes pour les observations.' };
    }

    var headerRow = 0;
    var maxText = 0;
    for (var r = 0; r < Math.min(rows.length, 5); r++) {
      var textCount = 0;
      for (var c = 0; c < rows[r].length; c++) {
        var v = String(rows[r][c] || '').trim();
        if (v && isNaN(v)) textCount++;
      }
      if (textCount > maxText) { maxText = textCount; headerRow = r; }
    }

    var headers = rows[headerRow].map(function(h) { return String(h || '').trim().toUpperCase(); });
    Logger.log('Observations - Headers: ' + headers.join(' | '));

    var colNom = findImportCol_(headers, ['^NOM$', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[E\u00c9]NOM', 'PRENOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colObs = findImportCol_(headers, ['OBSERVATION']);
    var colDefCarnet = findImportCol_(headers, ['D[E\u00c9]FAUT.*CARNET', 'CARNET']);
    var colLecon = findImportCol_(headers, ['LE[C\u00c7]ON.*NON', 'LECON']);
    var colOubli = findImportCol_(headers, ['OUBLI', 'MAT[E\u00c9]RIEL']);
    var colTravail = findImportCol_(headers, ['TRAVAIL.*NON', 'TRAVAIL']);
    var colEncourage = findImportCol_(headers, ['ENCOURAGEMENT', 'ENCOUR']);

    var observations = [];
    var currentClasse = '';

    for (var i = headerRow + 1; i < rows.length; i++) {
      var row = rows[i];
      var firstCell = String(row[0] || '').trim();

      // Ligne resume de classe (ex: triangle 4E 1)
      if (firstCell.match(/^[\u25B2\u25BC\u25BA\u25CF\u25A0\u25A1]/)) {
        currentClasse = firstCell.replace(/^[\u25B2\u25BC\u25BA\u25CF\u25A0\u25A1]\s*/, '').trim();
        continue;
      }

      var nom = colNom >= 0 ? String(row[colNom] || '').trim() : '';
      if (!nom) continue;
      if (nom.match(/^[\u25B2\u25BC\u25BA\u25CF\u25A0\u25A1]/) || nom.length < 2) continue;

      var prenom = colPrenom >= 0 ? String(row[colPrenom] || '').trim() : '';
      var classe = colClasse >= 0 ? String(row[colClasse] || '').trim() : currentClasse;

      var nbIncidents = 0;
      if (colDefCarnet >= 0) nbIncidents += (parseNote_(row[colDefCarnet]) || 0);
      if (colLecon >= 0)     nbIncidents += (parseNote_(row[colLecon]) || 0);
      if (colOubli >= 0)     nbIncidents += (parseNote_(row[colOubli]) || 0);
      if (colTravail >= 0)   nbIncidents += (parseNote_(row[colTravail]) || 0);

      var nbObs = colObs >= 0 ? (parseNote_(row[colObs]) || 0) : 0;
      var nbEncourage = colEncourage >= 0 ? (parseNote_(row[colEncourage]) || 0) : 0;

      observations.push({
        nom: nom.toUpperCase(),
        prenom: prenom,
        classe: classe,
        nbObs: nbObs,
        nbIncidents: nbIncidents,
        nbEncourage: nbEncourage
      });
    }

    Logger.log('Observations: ' + observations.length + ' eleves parses');
    return { success: true, observations: observations, count: observations.length };

  } catch (e) {
    Logger.log('Erreur v3_parseObservations: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// COMPILATION : FUSION + SCORES + ECRITURE ONGLETS SOURCES
// =============================================================================

/**
 * Compile toutes les donnees importees et cree les onglets sources peuples.
 *
 * @param {Object} data - Toutes les donnees parsees
 * @param {Object[]} data.eleves
 * @param {Object[]} data.notes
 * @param {Object[]} data.absences
 * @param {Object[]} data.punitions
 * @param {Object[]} data.observations
 * @returns {Object} {success, summary}
 */
function v3_compileImport(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('=== COMPILATION IMPORT MULTI-PASTE ===');

    if (!data || !data.eleves || data.eleves.length === 0) {
      return { success: false, error: 'Aucun eleve dans la liste. Collez d\'abord la liste eleves (etape 1).' };
    }

    // 1. CONSTRUIRE LA MAP ELEVES (cle = NOM|PRENOM)
    var studentMap = {};
    var classeGroups = {};

    for (var i = 0; i < data.eleves.length; i++) {
      var e = data.eleves[i];
      var cle = cleEleve_(e.nom, e.prenom);
      studentMap[cle] = {
        nom: e.nom, prenom: e.prenom, sexe: e.sexe || '', classe: e.classe,
        lv2: e.lv2 || '', opt: e.opt || '',
        moyennes: {}, oraux: {},
        dj: 0, nj: 0,
        nbPunitions: 0, nbObservations: 0, nbIncidents: 0, nbEncourage: 0
      };
      if (!classeGroups[e.classe]) classeGroups[e.classe] = [];
      classeGroups[e.classe].push(cle);
    }

    Logger.log('Eleves mappes: ' + Object.keys(studentMap).length);
    Logger.log('Classes: ' + Object.keys(classeGroups).join(', '));

    // 2. FUSIONNER LES NOTES
    var notesMatched = 0;
    if (data.notes && data.notes.length > 0) {
      for (var n = 0; n < data.notes.length; n++) {
        var noteData = data.notes[n];
        var cle2 = findMatchingStudent_(studentMap, noteData.nom, noteData.prenom);
        if (cle2) {
          studentMap[cle2].moyennes = noteData.moyennes || {};
          studentMap[cle2].oraux = noteData.oraux || {};
          notesMatched++;
        }
      }
      Logger.log('Notes fusionnees: ' + notesMatched + '/' + data.notes.length);
    }

    // 3. FUSIONNER LES ABSENCES
    var absMatched = 0;
    if (data.absences && data.absences.length > 0) {
      for (var a = 0; a < data.absences.length; a++) {
        var absData = data.absences[a];
        var cle3 = findMatchingStudent_(studentMap, absData.nom, absData.prenom);
        if (cle3) {
          studentMap[cle3].dj = absData.dj || 0;
          studentMap[cle3].nj = absData.nj || 0;
          absMatched++;
        }
      }
      Logger.log('Absences fusionnees: ' + absMatched + '/' + data.absences.length);
    }

    // 4. FUSIONNER PUNITIONS
    var punMatched = 0;
    if (data.punitions && data.punitions.length > 0) {
      for (var p = 0; p < data.punitions.length; p++) {
        var punData = data.punitions[p];
        var cle4 = findMatchingStudent_(studentMap, punData.nom, punData.prenom);
        if (cle4) {
          studentMap[cle4].nbPunitions += (punData.nb || 0);
          punMatched++;
        }
      }
      Logger.log('Punitions fusionnees: ' + punMatched + '/' + data.punitions.length);
    }

    // 5. FUSIONNER OBSERVATIONS
    var obsMatched = 0;
    if (data.observations && data.observations.length > 0) {
      for (var o = 0; o < data.observations.length; o++) {
        var obsData = data.observations[o];
        var cle5 = findMatchingStudent_(studentMap, obsData.nom, obsData.prenom);
        if (cle5) {
          studentMap[cle5].nbObservations += (obsData.nbObs || 0);
          studentMap[cle5].nbIncidents += (obsData.nbIncidents || 0);
          studentMap[cle5].nbEncourage += (obsData.nbEncourage || 0);
          obsMatched++;
        }
      }
      Logger.log('Observations fusionnees: ' + obsMatched + '/' + data.observations.length);
    }

    // 6. CALCULER LES SCORES
    var scoresCount = 0;
    for (var cle6 in studentMap) {
      var st = studentMap[cle6];
      st.scoreTRA = calcScoreTRA_import_(st.moyennes);
      st.scorePART = calcScorePART_import_(st.oraux);
      st.scoreABS = calcScoreABS_import_(st.dj, st.nj);
      st.scoreCOM = calcScoreCOM_import_(st.nbPunitions, st.nbIncidents, st.nbObservations);
      if (st.scoreTRA || st.scorePART || st.scoreABS || st.scoreCOM) scoresCount++;
    }

    Logger.log('Scores calcules: ' + scoresCount);

    // 7. ECRIRE DANS LES ONGLETS SOURCES
    var sourceHeaders = [
      'ID_ELEVE', 'NOM', 'PRENOM', 'NOM_PRENOM', 'SEXE', 'LV2', 'OPT',
      'COM', 'TRA', 'PART', 'ABS', 'DISPO', 'ASSO', 'DISSO', 'SOURCE'
    ];

    var ruleCRIT = SpreadsheetApp.newDataValidation()
      .requireValueInList(['', '1', '2', '3', '4'], true)
      .setAllowInvalid(false)
      .build();

    var sheetsCreated = 0;
    var sheetsUpdated = 0;

    for (var classeName in classeGroups) {
      var studentKeys = classeGroups[classeName];
      var sheet = ss.getSheetByName(classeName);

      if (!sheet) {
        sheet = ss.insertSheet(classeName);
        sheetsCreated++;
      } else {
        sheet.clear();
        sheetsUpdated++;
      }

      sheet.getRange(1, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
      sheet.getRange(1, 1, 1, sourceHeaders.length)
        .setFontWeight('bold').setBackground('#d9ead3').setFontSize(10);
      sheet.setFrozenRows(1);

      var studentRows = [];
      for (var si = 0; si < studentKeys.length; si++) {
        var stKey = studentKeys[si];
        var st2 = studentMap[stKey];
        var id = classeName.replace(/\s/g, '') + String(si + 1).padStart(3, '0');
        var nomPrenom = st2.nom + (st2.prenom ? ' ' + st2.prenom : '');

        studentRows.push([
          id, st2.nom, st2.prenom, nomPrenom, st2.sexe, st2.lv2, st2.opt,
          st2.scoreCOM !== null ? String(st2.scoreCOM) : '',
          st2.scoreTRA !== null ? String(st2.scoreTRA) : '',
          st2.scorePART !== null ? String(st2.scorePART) : '',
          st2.scoreABS !== null ? String(st2.scoreABS) : '',
          '', '', '', classeName
        ]);
      }

      if (studentRows.length > 0) {
        sheet.getRange(2, 1, studentRows.length, sourceHeaders.length).setValues(studentRows);
        [8, 9, 10, 11].forEach(function(col) {
          sheet.getRange(2, col, studentRows.length, 1).setDataValidation(ruleCRIT);
        });
      }

      var widths = { 1:100, 2:120, 3:120, 4:180, 5:60, 6:55, 7:65, 8:50, 9:50, 10:50, 11:50, 12:85, 13:70, 14:70, 15:60 };
      for (var col in widths) sheet.setColumnWidth(parseInt(col), widths[col]);
    }

    // 8. LISTES DEROULANTES + FORMATAGE CONDITIONNEL
    try { ajouterListesDeroulantes(); } catch (e2) { Logger.log('ajouterListesDeroulantes: ' + e2.message); }

    // 9. NOM_PRENOM + IDs + CONSOLIDATION
    try { genererNomPrenomEtID(); } catch (e3) { Logger.log('genererNomPrenomEtID: ' + e3.message); }

    var consolResult = '';
    try { consolResult = consoliderDonnees(); } catch (e4) { consolResult = 'Non disponible'; }

    // 10. DISSO AUTO-SUGGESTION
    var dissoSuggestion = suggestDissoGroups_(studentMap, classeGroups);

    Logger.log('=== COMPILATION TERMINEE ===');
    var summary = {
      totalStudents: data.eleves.length,
      classesProcessed: Object.keys(classeGroups).length,
      classesList: Object.keys(classeGroups),
      sheetsCreated: sheetsCreated, sheetsUpdated: sheetsUpdated,
      scoresInjected: scoresCount,
      notesMatched: notesMatched, absMatched: absMatched,
      punMatched: punMatched, obsMatched: obsMatched,
      consolidation: consolResult,
      dissoSuggestion: dissoSuggestion
    };

    logAction('Import multi-paste: ' + data.eleves.length + ' eleves dans ' + Object.keys(classeGroups).length + ' classes');
    return { success: true, summary: summary };

  } catch (e) {
    Logger.log('Erreur v3_compileImport: ' + e.toString());
    Logger.log(e.stack);
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// FONCTIONS DE MATCHING
// =============================================================================

function findMatchingStudent_(studentMap, nom, prenom) {
  var cle = cleEleve_(nom, prenom);
  if (studentMap[cle]) return cle;

  var normNom = normaliserNom_(nom);
  var normPrenom = normaliserNom_(prenom);

  for (var key in studentMap) {
    var st = studentMap[key];
    if (normaliserNom_(st.nom) === normNom && normaliserNom_(st.prenom) === normPrenom) return key;
  }

  if (!prenom && nom) {
    var parts = nom.trim().split(/\s+/);
    if (parts.length >= 2) {
      var testNom = normaliserNom_(parts[0]);
      var testPrenom = normaliserNom_(parts.slice(1).join(' '));
      for (var key2 in studentMap) {
        var st2 = studentMap[key2];
        if (normaliserNom_(st2.nom) === testNom && normaliserNom_(st2.prenom) === testPrenom) return key2;
      }
    }
  }

  var nomMatches = [];
  for (var key3 in studentMap) {
    if (normaliserNom_(studentMap[key3].nom) === normNom) nomMatches.push(key3);
  }
  if (nomMatches.length === 1) return nomMatches[0];

  return null;
}

// =============================================================================
// CALCUL DES SCORES (VERSION IMPORT)
// =============================================================================

function calcScoreTRA_import_(moyennes) {
  if (!moyennes || Object.keys(moyennes).length === 0) return null;
  var coeffMap = { 'FRANC':4.5, 'MATH':3.5, 'HG':3.0, 'ANG':3.0, 'LV2':2.5, 'EPS':2.0, 'PHCH':1.5, 'SVT':1.5, 'TECH':1.5, 'APLA':1.0, 'MUS':1.0, 'LAT':1.0 };
  var totalPts = 0, totalCoeff = 0;
  for (var id in moyennes) {
    var note = moyennes[id];
    if (note === null || note === undefined) continue;
    var coeff = coeffMap[id] || 1.0;
    totalPts += note * coeff;
    totalCoeff += coeff;
  }
  if (totalCoeff === 0) return null;
  var moy = Math.round(totalPts / totalCoeff * 100) / 100;
  return attribuerScoreParSeuil_(moy, SCORES_CONFIG.SEUILS_TRA);
}

function calcScorePART_import_(oraux) {
  if (!oraux || Object.keys(oraux).length === 0) return null;
  var notes = [];
  for (var id in oraux) {
    if (oraux[id] !== null && oraux[id] !== undefined) notes.push(oraux[id]);
  }
  if (notes.length === 0) return null;
  var moy = notes.reduce(function(a, b) { return a + b; }, 0) / notes.length;
  moy = Math.round(moy * 100) / 100;
  return attribuerScoreParSeuil_(moy, SCORES_CONFIG.SEUILS_PART);
}

function calcScoreABS_import_(dj, nj) {
  if (dj === 0 && nj === 0) return 4;
  var seuils = SCORES_CONFIG.SEUILS_ABS;
  var scoreDJ = attribuerScoreParSeuil_(dj, seuils.DJ);
  var scoreNJ = attribuerScoreParSeuil_(nj, seuils.NJ);
  return Math.ceil(scoreDJ * seuils.poidsDJ + scoreNJ * seuils.poidsNJ);
}

function calcScoreCOM_import_(nbPunitions, nbIncidents, nbObservations) {
  var total = (nbPunitions || 0) + (nbIncidents || 0) + Math.ceil((nbObservations || 0) * 0.5);
  return attribuerScoreParSeuil_(total, SCORES_CONFIG.SEUILS_COM);
}

// =============================================================================
// DISSO AUTO-SUGGESTION
// =============================================================================

function suggestDissoGroups_(studentMap, classeGroups) {
  var ranked = [];
  for (var cle in studentMap) {
    var st = studentMap[cle];
    var penibilite = (st.nbPunitions || 0) * 2 + (st.nbIncidents || 0) * 3 + (st.nbObservations || 0);
    if (penibilite > 0) {
      ranked.push({
        nom: st.nom, prenom: st.prenom, classe: st.classe,
        penibilite: penibilite,
        details: { punitions: st.nbPunitions || 0, incidents: st.nbIncidents || 0, observations: st.nbObservations || 0 }
      });
    }
  }
  ranked.sort(function(a, b) { return b.penibilite - a.penibilite; });

  var nbDest = 5;
  try { var config = getConfig(); nbDest = parseInt(config.NB_DEST) || 5; } catch (e) {}

  var groups = [];
  var groupIndex = 0;
  for (var i = 0; i < ranked.length && groupIndex < 3; i += nbDest) {
    var group = { label: 'Lot ' + (groupIndex + 1), students: [] };
    for (var j = i; j < Math.min(i + nbDest, ranked.length); j++) {
      group.students.push(ranked[j]);
    }
    if (group.students.length > 1) { groups.push(group); groupIndex++; }
  }

  return { totalPenibles: ranked.length, topStudents: ranked.slice(0, 15), groups: groups };
}

// =============================================================================
// STATUT IMPORT
// =============================================================================

function v3_getImportStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = getConfig();
    var niveau = config.NIVEAU || '';
    var prefixe = '';
    try { prefixe = determinerPrefixeSource(niveau); } catch (e) { prefixe = ''; }

    var sheets = ss.getSheets();
    var sourcesRemplies = 0, sourcesVides = 0, totalEleves = 0;
    var sourcesDetail = [];

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (prefixe && name.startsWith(prefixe)) {
        var nbEleves = Math.max(0, sheets[i].getLastRow() - 1);
        if (nbEleves > 0) { sourcesRemplies++; totalEleves += nbEleves; }
        else sourcesVides++;
        sourcesDetail.push({ name: name, eleves: nbEleves });
      }
    }

    return { success: true, sourcesRemplies: sourcesRemplies, sourcesVides: sourcesVides, totalEleves: totalEleves, sources: sourcesDetail };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
