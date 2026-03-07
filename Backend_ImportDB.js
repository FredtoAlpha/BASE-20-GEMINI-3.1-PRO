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
 *   4. Comportement     -> COM (observations + punitions + incidents)
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
 * Normalise un nom : MAJUSCULES, trim, supprime accents, guillemets, espaces invisibles
 */
function normaliserNom_(s) {
  if (!s) return '';
  return String(s).trim()
    .replace(/[\u201C\u201D\u201E\u201F\u00AB\u00BB"'`\u2018\u2019\u201A\u2039\u203A]/g, '')  // guillemets typo + droits
    .replace(/[\u00A0\u2007\u202F\u200B\u200C\u200D\uFEFF]/g, ' ')  // espaces insecables/invisibles
    .replace(/\s+/g, ' ')  // compacter espaces multiples
    .trim()
    .toUpperCase()
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
  var s = String(val).trim().toUpperCase();
  if (s === '\u2640' || s === 'F' || s.indexOf('FEM') === 0 || s.indexOf('FÉM') === 0) return 'F';
  if (s === '\u2642' || s === 'M' || s.indexOf('MAS') === 0) return 'M';
  // Fallback: premiere lettre si F ou M, sinon vide (evite de passer une valeur invalide)
  var first = s.charAt(0);
  if (first === 'F') return 'F';
  if (first === 'M') return 'M';
  Logger.log('[WARN] parseSexe_: valeur non reconnue "' + String(val).trim() + '"');
  return '';
}

/**
 * Parse la colonne "Projet d'accompagnement" / DISPO Pronote.
 * Extrait le dispositif le plus fort si plusieurs presents.
 * Priorite : GEVASCO > PPS > PAP > PPRE > ULIS > SEGPA > UPE2A
 * Ex: "PAP, GEVASCO" -> "GEVASCO"
 *     "PAP" -> "PAP"
 *     "Notification MDPH, Convention ergothérapeute, GEVASCO" -> "GEVASCO"
 */
function parseDispo_(val) {
  if (!val) return '';
  var s = String(val).trim().toUpperCase();
  if (!s) return '';
  // Priorite decroissante (le premier trouve gagne)
  var priorite = ['GEVASCO', 'PPS', 'PAP', 'PPRE', 'ULIS', 'SEGPA', 'UPE2A'];
  for (var i = 0; i < priorite.length; i++) {
    if (s.indexOf(priorite[i]) >= 0) return priorite[i];
  }
  return '';
}

/**
 * Normalise le nom de classe Pronote en format X°Y
 * "4E 1" -> "4°1", "4e1" -> "4°1", "4EME1" -> "4°1", "4eme 1" -> "4°1"
 */
function normaliserClasse_(classe) {
  if (!classe) return '';
  var s = String(classe).trim();
  // "4E 1" / "4e1" / "4E1" / "4EME 1" / "4ème 1"
  s = s.replace(/(\d+)\s*[eèéÈÉ][mM]?[eèéÈÉ]?\s*(\d+)/i, '$1°$2');
  // Si toujours pas converti, essayer "4 E 1"
  s = s.replace(/(\d+)\s+[eèéÈÉ]\s+(\d+)/i, '$1°$2');
  return s;
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

  // Ordre = priorite : LATIN > GREC > CHAV > CHINOIS > LCALA > LLCA > EURO
  var optionsConnues = ['LATIN', 'GREC', 'CHAV', 'CHINOIS', 'LCALA', 'LLCA', 'EURO'];

  // Mapping des libelles Pronote vers le code option interne
  var optionsAliases = {
    'CHANT': 'CHAV',       // "CHANT CHORAL COLLECT" -> CHAV
    'CHORAL': 'CHAV',      // "CHORALE" -> CHAV
    'LCA': 'LATIN'         // "LCA LATIN" -> LATIN (deja couvert par LATIN mais securite)
  };

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

    // Si opt deja trouve, ne pas ecraser (premiere option = plus haute priorite)
    if (result.opt) continue;

    // Detecter option par nom exact d'abord
    var matched = false;
    for (var j = 0; j < optionsConnues.length; j++) {
      if (p.indexOf(optionsConnues[j]) >= 0) {
        result.opt = optionsConnues[j];
        matched = true;
        break;
      }
    }

    // Puis par aliases Pronote (ex: "CHANT CHORAL COLLECT" -> CHAV)
    if (!matched) {
      for (var alias in optionsAliases) {
        if (p.indexOf(alias) >= 0) {
          result.opt = optionsAliases[alias];
          break;
        }
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
    // Colonnes directes LV2 et DISPO (prioritaires sur parseOptions_ si presentes)
    var colLangue = findImportCol_(headers, ['LANGUE', '^LV2$']);
    var colDispo = findImportCol_(headers, ['DISPO', 'DISPOSITIF', 'PROJET.*ACC', 'ACCOMPAGNEMENT']);

    if (colNom === -1) return { success: false, error: 'Colonne NOM introuvable. Headers: ' + headers.join(', ') };
    if (colClasse === -1) return { success: false, error: 'Colonne CLASSE introuvable. Headers: ' + headers.join(', ') };

    Logger.log('Colonnes: NOM=' + colNom + ' PRENOM=' + colPrenom + ' SEXE=' + colSexe +
      ' CLASSE=' + colClasse + ' OPT=' + colOptions + ' LANGUE=' + colLangue + ' DISPO=' + colDispo);

    // Map de normalisation LV2 : nom complet Pronote -> code court Base Opti
    var languesLV2Map = {
      'ESPAGNOL': 'ESP', 'ALLEMAND': 'ALL', 'ITALIEN': 'ITA',
      'CHINOIS': 'CHI', 'PORTUGAIS': 'PT', 'ARABE': 'ARA',
      'RUSSE': 'RUS', 'JAPONAIS': 'JAP'
    };

    var eleves = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var row = rows[i];
      var nom = String(row[colNom] || '').trim();
      if (!nom) continue;

      var prenom = colPrenom >= 0 ? String(row[colPrenom] || '').trim() : '';
      var sexe = colSexe >= 0 ? parseSexe_(row[colSexe]) : '';
      var classe = normaliserClasse_(row[colClasse]);
      if (!classe) continue;

      // Fallback : parseOptions_ sur la colonne "Toutes les options"
      var opts = { lv2: '', opt: '' };
      if (colOptions >= 0) {
        opts = parseOptions_(row[colOptions]);
      }

      // PRIORITE : extraction directe colonne LANGUE/LV2 (court-circuite parseOptions_)
      if (colLangue >= 0) {
        var lv2Direct = String(row[colLangue] || '').trim().toUpperCase();
        if (lv2Direct) {
          opts.lv2 = languesLV2Map[lv2Direct] || lv2Direct;
        }
      }

      // Extraction DISPO avec priorite (GEVASCO > PPS > PAP > PPRE > ...)
      var dispoDirect = '';
      if (colDispo >= 0) {
        dispoDirect = parseDispo_(row[colDispo]);
      }

      eleves.push({
        nom: nom.toUpperCase(),
        prenom: prenom,
        sexe: sexe,
        classe: classe,
        lv2: opts.lv2,
        opt: opts.opt,
        dispo: dispoDirect
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

    // DETECTER LA SOUS-LIGNE "Abs" : Pronote exporte [Note | Abs] pour chaque matiere
    // La ligne juste apres le header contient "Abs", "1", "2", etc.
    var absColsSet = {};  // colonnes qui contiennent des absences (a exclure)
    var dataStartRow = headerRow + 1;
    if (headerRow + 1 < rows.length) {
      var subRow = rows[headerRow + 1];
      var absCount = 0;
      for (var sc = 0; sc < subRow.length; sc++) {
        var sv = String(subRow[sc] || '').trim().toUpperCase();
        if (sv === 'ABS' || sv === 'ABSENCES' || sv === 'NB ABS') {
          absColsSet[sc] = true;
          absCount++;
        }
      }
      if (absCount > 0) {
        dataStartRow = headerRow + 2;  // Sauter la sous-ligne "Abs"
        Logger.log('Notes - Sous-ligne Abs detectee: ' + absCount + ' colonnes Abs filtrees, donnees a partir ligne ' + dataStartRow);
      }
    }

    var colNom = findImportCol_(rawHeaders, ['^NOM$', '^[E\u00c9]L[E\u00c8]VE', 'NOM']);
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

      // FILTRER les colonnes "Abs" (absences par matiere, pas des notes !)
      var gradeCols = matchedCols.filter(function(col) { return !absColsSet[col]; });
      if (gradeCols.length === 0) gradeCols = matchedCols;  // fallback si pas de sous-ligne Abs

      // Apres filtrage des Abs :
      //   1 col = c'est la moyenne
      //   2+ cols = dernier = moyenne (ou oral pour langues), avant-dernier = ecrit/note1
      var colMoy, colOral;

      if (gradeCols.length >= 2) {
        colMoy = gradeCols[gradeCols.length - 1];  // Derniere col note = moyenne
        if (mat.hasOral) {
          // PRONOTE langues : col[0]=moyenne, col[1]=ecrit, col[2]=oral
          // 2 cols : col[0]=moyenne(TRA), col[1]=oral(PART)
          colMoy = gradeCols[0];
          colOral = gradeCols[gradeCols.length - 1];  // derniere = oral (PART)
        }
      } else {
        colMoy = gradeCols[0];
        colOral = null;
      }

      gradeMap[mat.id] = { col: colMoy, coeff: mat.coeff };
      detectedSubjects.push(mat.id + '(col' + colMoy + ',grade_cols=' + gradeCols.join('+') + ')');

      if (mat.hasOral && colOral !== null && colOral !== colMoy) {
        oralMap[mat.id] = colOral;
      }
    }

    Logger.log('Matieres detectees: ' + detectedSubjects.join(', '));
    Logger.log('Oraux detectes: ' + JSON.stringify(oralMap));

    // PARSER LES LIGNES ELEVES (demarre apres la sous-ligne Abs si detectee)
    var notes = [];
    var classeDetected = '';

    for (var i = dataStartRow; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.length < 3) continue;

      var nom = '', prenom = '', classe = '';

      if (colNom >= 0) nom = String(row[colNom] || '').trim();
      if (colPrenom >= 0) prenom = String(row[colPrenom] || '').trim();
      if (colClasse >= 0) classe = normaliserClasse_(row[colClasse]);

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

      // Nettoyer guillemets/invisibles dans nom et prenom (meme nettoyage que normaliserNom_)
      nom = nom.replace(/[\u201C\u201D\u201E\u201F\u00AB\u00BB"'`\u2018\u2019\u201A\u2039\u203A]/g, '')
               .replace(/[\u00A0\u2007\u202F\u200B\u200C\u200D\uFEFF]/g, ' ').replace(/\s+/g, ' ').trim();
      prenom = prenom.replace(/[\u201C\u201D\u201E\u201F\u00AB\u00BB"'`\u2018\u2019\u201A\u2039\u203A]/g, '')
                     .replace(/[\u00A0\u2007\u202F\u200B\u200C\u200D\uFEFF]/g, ' ').replace(/\s+/g, ' ').trim();

      // Si prenom vide et nom contient plusieurs mots, tenter split NOM PRENOM
      if (!prenom && nom.indexOf(' ') >= 0) {
        var nameParts = nom.split(/\s+/);
        nom = nameParts[0];
        prenom = nameParts.slice(1).join(' ');
      }

      // FILTRER les lignes parasites : en-tetes repetes et lignes de resume
      var nomUpper = nom.toUpperCase();
      if (nomUpper === 'NOM' || nomUpper === 'ELEVE' || nomUpper === 'ÉLÈVE' || nomUpper === 'ÉLÈVES'
          || nomUpper.indexOf('MOYENNE') >= 0 || nomUpper.indexOf('CLASSE') >= 0
          || nomUpper.indexOf('GROUPE') >= 0 || nomUpper.indexOf('GENERAL') >= 0
          || nomUpper.indexOf('GÉNÉRAL') >= 0) {
        continue;
      }

      if (classe) classeDetected = normaliserClasse_(classe);

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
    // [DIAG] Echantillon noms nettoyes
    if (notes.length > 0) {
      var sampleNames = notes.slice(0, 5).map(function(n) { return n.nom + '|' + n.prenom; });
      Logger.log('[DIAG] Notes echantillon noms nettoyes: ' + sampleNames.join(', '));
    }

    return {
      success: true,
      notes: notes,
      count: notes.length,
      classe: normaliserClasse_(classeDetected),
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

    var colNom = findImportCol_(headers, ['^NOM$', '^[E\u00c9]L[E\u00c8]VE', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[E\u00c9]NOM', 'PRENOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colDJ = findImportCol_(headers, ['^DJ$', 'DEMI.?JOURN', 'NB.*ABS', 'TOTAL']);
    var colJustif = findImportCol_(headers, ['JUSTIFI', 'JUST\\.?', 'STATUT']);
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
        absences.push({ nom: nom.toUpperCase(), prenom: prenom, classe: normaliserClasse_(classe), dj: dj, nj: nj });
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
          perStudent[cle] = { nom: nom2, prenom: prenom2, classe: normaliserClasse_(classe2), dj: 0, nj: 0 };
        }
        var djVal = colDJ >= 0 ? (parseNote_(row2[colDJ]) || 1) : 1;
        perStudent[cle].dj += djVal;
        var justifVal = colJustif >= 0 ? String(row2[colJustif] || '').trim().toUpperCase() : '';
        // Gere: "OUI", "Justifiée", "Régularisée" comme justifie ; "NON", "Non justifiée" comme non justifie
        var isJustifie = justifVal === 'OUI' || justifVal === 'O' || justifVal === 'X' || justifVal === '1'
          || (justifVal.indexOf('JUSTIFI') >= 0 && justifVal.indexOf('NON') === -1)
          || justifVal.indexOf('REGULARI') >= 0 || justifVal.indexOf('RÉGULARI') >= 0;
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

    var colNom = findImportCol_(headers, ['^NOM$', '^[E\u00c9]L[E\u00c8]VE', 'NOM']);
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

      punitions.push({ nom: nom, prenom: prenom, classe: normaliserClasse_(classe), nb: nb });
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
    Logger.log('Observations [v2] - Headers bruts: ' + headers.join(' | '));

    // "Élèves" dans Pronote, pas "Nom" : on ajoute le pattern ELEVE
    var colNom = findImportCol_(headers, ['^NOM$', '^[E\u00c9]L[E\u00c8]VE', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[E\u00c9]NOM', 'PRENOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    // Col C : Observations (negatif)
    var colObs = findImportCol_(headers, ['OBSERVATION']);
    // Col D : Encouragements/valorisations (attenuation)
    var colEncourage = findImportCol_(headers, ['ENCOURAGEMENT', 'ENCOUR', 'VALORIS']);
    // Col F : Lecon non apprise (negatif)
    var colLecon = findImportCol_(headers, ['LE[C\u00c7]ON.*NON', 'LECON']);
    // Col G : Oubli de materiel (negatif)
    var colOubli = findImportCol_(headers, ['OUBLI', 'MAT[E\u00c9]RIEL']);
    // Col H : Travail non fait (negatif)
    var colTravail = findImportCol_(headers, ['TRAVAIL.*NON', 'TRAVAIL']);

    Logger.log('Observations - Colonnes detectees: NOM=' + colNom + ' CLASSE=' + colClasse +
      ' OBS(C)=' + colObs + ' ENCOURAGE(D)=' + colEncourage +
      ' LECON(F)=' + colLecon + ' OUBLI(G)=' + colOubli + ' TRAVAIL(H)=' + colTravail);

    var observations = [];
    var currentClasse = '';
    var nbLignesIgnorees = 0;

    // Regex elargi pour detecter les lignes-resume de classe (tous les triangles/fleches Unicode)
    var classSummaryRe = /^[\u25A0-\u25FF\u2190-\u21FF\u2794\u27A1]/;
    // Regex pour detecter les lignes de synthese : "NN élèves" ou "NN eleves"
    var syntheseRe = /^\d+\s+[e\u00e9]l[e\u00e8]ve/i;

    for (var i = headerRow + 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.length < 2) continue;

      // Detection de la classe courante via col CLASSE
      var classeCell = colClasse >= 0 ? String(row[colClasse] || '').trim() : '';
      if (classeCell) currentClasse = classeCell;

      var firstCell = String(row[0] || '').trim();

      // Ligne resume de classe : detectee par symbole Unicode uniquement
      if (classSummaryRe.test(firstCell)) {
        currentClasse = firstCell.replace(/^[\u25A0-\u25FF\u2190-\u21FF\u2794\u27A1]\s*/, '').trim();
        currentClasse = currentClasse.replace(/[,;].*/, '').trim();
        nbLignesIgnorees++;
        continue;
      }

      var nom = colNom >= 0 ? String(row[colNom] || '').trim() : '';
      // Ligne de synthese : colonne NOM vide ou contient "NN élèves"
      if (!nom || syntheseRe.test(nom)) { nbLignesIgnorees++; continue; }
      if (classSummaryRe.test(nom) || nom.length < 2) { nbLignesIgnorees++; continue; }

      var prenom = colPrenom >= 0 ? String(row[colPrenom] || '').trim() : '';
      var classe = normaliserClasse_(classeCell || currentClasse);

      // Bloc negatif : C + F + G + H
      var nbObsNeg = 0;
      if (colObs >= 0)     nbObsNeg += (parseNote_(row[colObs]) || 0);
      if (colLecon >= 0)   nbObsNeg += (parseNote_(row[colLecon]) || 0);
      if (colOubli >= 0)   nbObsNeg += (parseNote_(row[colOubli]) || 0);
      if (colTravail >= 0) nbObsNeg += (parseNote_(row[colTravail]) || 0);

      // D = encouragements (attenuation, pas negatif)
      var nbEncourage = colEncourage >= 0 ? (parseNote_(row[colEncourage]) || 0) : 0;

      observations.push({
        nom: nom.toUpperCase(),
        prenom: prenom,
        classe: classe,
        nbObservationsNeg: nbObsNeg,
        nbEncourage: nbEncourage
      });
    }

    // Logs de controle : distribution
    var distNeg = {}, distEnc = {};
    for (var d = 0; d < observations.length; d++) {
      var kn = observations[d].nbObservationsNeg;
      var ke = observations[d].nbEncourage;
      distNeg[kn] = (distNeg[kn] || 0) + 1;
      distEnc[ke] = (distEnc[ke] || 0) + 1;
    }
    Logger.log('Observations: ' + observations.length + ' eleves parses, ' + nbLignesIgnorees + ' lignes ignorees (synthese)');
    Logger.log('Observations - Distribution nbObservationsNeg: ' + JSON.stringify(distNeg));
    Logger.log('Observations - Distribution nbEncourage: ' + JSON.stringify(distEnc));

    return { success: true, observations: observations, count: observations.length };

  } catch (e) {
    Logger.log('Erreur v3_parseObservations: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// 5. PARSING RETENUES (HEURES DE RETENUE)
// =============================================================================

/**
 * Parse les retenues (heures de retenue). Format libre : NOM, CLASSE, NB ou liste evenementielle.
 * @param {Array[]} rows
 * @returns {Object} {success, retenues, count}
 */
function v3_parseRetenues(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes pour les retenues.' };
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
    Logger.log('Retenues - Headers: ' + headers.join(' | '));

    var colNom = findImportCol_(headers, ['^NOM$', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[EÉ]NOM', 'PRENOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colNb = findImportCol_(headers, ['^NB', 'NOMBRE', 'TOTAL', 'RETENUE', 'HEURE']);
    var colDate = findImportCol_(headers, ['DATE', 'DU']);

    // Si pas de colonne NB, c'est une liste evenementielle (1 ligne = 1 retenue)
    var isEventList = colNb === -1 || colDate >= 0;

    var retenues;
    if (isEventList) {
      // Compter le nombre de lignes par eleve
      var perStudent = {};
      for (var i = headerRow + 1; i < rows.length; i++) {
        var row = rows[i];
        var nom = '', prenom = '';
        if (colNom >= 0) {
          var nomVal = String(row[colNom] || '').trim();
          if (!nomVal) continue;
          var parts = nomVal.split(/\s+/);
          if (parts.length >= 2 && colPrenom === -1) {
            nom = parts[0].toUpperCase();
            prenom = parts.slice(1).join(' ');
          } else {
            nom = nomVal.toUpperCase();
          }
        }
        if (colPrenom >= 0) prenom = String(row[colPrenom] || '').trim();
        if (!nom) continue;
        var classe = colClasse >= 0 ? normaliserClasse_(row[colClasse]) : '';
        var cle = cleEleve_(nom, prenom);
        if (!perStudent[cle]) perStudent[cle] = { nom: nom, prenom: prenom, classe: classe, nb: 0 };
        perStudent[cle].nb++;
        if (classe) perStudent[cle].classe = classe;
      }
      retenues = [];
      for (var key in perStudent) retenues.push(perStudent[key]);
    } else {
      retenues = [];
      for (var i2 = headerRow + 1; i2 < rows.length; i2++) {
        var row2 = rows[i2];
        var nom2 = '', prenom2 = '';
        if (colNom >= 0) {
          var nomVal2 = String(row2[colNom] || '').trim();
          if (!nomVal2) continue;
          var parts2 = nomVal2.split(/\s+/);
          if (parts2.length >= 2 && colPrenom === -1) {
            nom2 = parts2[0].toUpperCase();
            prenom2 = parts2.slice(1).join(' ');
          } else {
            nom2 = nomVal2.toUpperCase();
          }
        }
        if (colPrenom >= 0) prenom2 = String(row2[colPrenom] || '').trim();
        if (!nom2) continue;
        var classe2 = colClasse >= 0 ? normaliserClasse_(row2[colClasse]) : '';
        var nb = colNb >= 0 ? (parseNote_(row2[colNb]) || 0) : 1;
        retenues.push({ nom: nom2, prenom: prenom2, classe: classe2, nb: nb });
      }
    }

    Logger.log('Retenues: ' + retenues.length + ' eleves parses');
    return { success: true, retenues: retenues, count: retenues.length };

  } catch (e) {
    Logger.log('Erreur v3_parseRetenues: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// 6. PARSING INCIDENTS (FAITS GRAVES)
// =============================================================================

/**
 * Parse les incidents Pronote. Format : NOM, CLASSE, (NB ou evenementiel)
 * @param {Array[]} rows
 * @returns {Object} {success, incidents, count}
 */
function v3_parseIncidents(rows) {
  try {
    if (!rows || rows.length < 2) {
      return { success: false, error: 'Donnees insuffisantes pour les incidents.' };
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
    Logger.log('Incidents - Headers: ' + headers.join(' | '));

    var colNom = findImportCol_(headers, ['^NOM$', '^[E\u00c9]L[E\u00c8]VE', 'NOM']);
    var colPrenom = findImportCol_(headers, ['PR[EÉ]NOM', 'PRENOM']);
    var colClasse = findImportCol_(headers, ['CLASSE']);
    var colNb = findImportCol_(headers, ['^NB', 'NOMBRE', 'TOTAL', 'INCIDENT']);
    var colDate = findImportCol_(headers, ['DATE', 'DU']);

    var isEventList = colNb === -1 || colDate >= 0;

    var incidents;
    if (isEventList) {
      var perStudent = {};
      for (var i = headerRow + 1; i < rows.length; i++) {
        var row = rows[i];
        var nom = '', prenom = '';
        if (colNom >= 0) {
          var nomVal = String(row[colNom] || '').trim();
          if (!nomVal) continue;
          var parts = nomVal.split(/\s+/);
          if (parts.length >= 2 && colPrenom === -1) {
            nom = parts[0].toUpperCase();
            prenom = parts.slice(1).join(' ');
          } else {
            nom = nomVal.toUpperCase();
          }
        }
        if (colPrenom >= 0) prenom = String(row[colPrenom] || '').trim();
        if (!nom) continue;
        var classe = colClasse >= 0 ? normaliserClasse_(row[colClasse]) : '';
        var cle = cleEleve_(nom, prenom);
        if (!perStudent[cle]) perStudent[cle] = { nom: nom, prenom: prenom, classe: classe, nb: 0 };
        perStudent[cle].nb++;
        if (classe) perStudent[cle].classe = classe;
      }
      incidents = [];
      for (var key in perStudent) incidents.push(perStudent[key]);
    } else {
      incidents = [];
      for (var i2 = headerRow + 1; i2 < rows.length; i2++) {
        var row2 = rows[i2];
        var nom2 = '', prenom2 = '';
        if (colNom >= 0) {
          var nomVal2 = String(row2[colNom] || '').trim();
          if (!nomVal2) continue;
          var parts2 = nomVal2.split(/\s+/);
          if (parts2.length >= 2 && colPrenom === -1) {
            nom2 = parts2[0].toUpperCase();
            prenom2 = parts2.slice(1).join(' ');
          } else {
            nom2 = nomVal2.toUpperCase();
          }
        }
        if (colPrenom >= 0) prenom2 = String(row2[colPrenom] || '').trim();
        if (!nom2) continue;
        var classe2 = colClasse >= 0 ? normaliserClasse_(row2[colClasse]) : '';
        var nb = colNb >= 0 ? (parseNote_(row2[colNb]) || 0) : 1;
        incidents.push({ nom: nom2, prenom: prenom2, classe: classe2, nb: nb });
      }
    }

    Logger.log('Incidents: ' + incidents.length + ' eleves parses');
    return { success: true, incidents: incidents, count: incidents.length };

  } catch (e) {
    Logger.log('Erreur v3_parseIncidents: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// COMPILATION : FUSION + SCORES + ECRITURE ONGLETS SOURCES
// =============================================================================

/**
 * Compile toutes les donnees importees et cree les onglets sources peuples.
 *
 * IMPORTANT : Le client envoie des objets-resultat (ex: data.eleves = {success, eleves: [...], count})
 * Il faut donc UNWRAP pour acceder aux tableaux bruts.
 *
 * @param {Object} data - Toutes les donnees parsees (objets-resultat du client)
 * @returns {Object} {success, summary, dissoSuggestions}
 */
function v3_compileImport(data) {
  var runId = RunAudit_createId();
  var timer = RunAudit_startTimer();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    RunAudit_log(runId, 'INFO', '=== COMPILATION IMPORT MULTI-PASTE ===');

    // *** UNWRAP : extraire les tableaux bruts des objets-resultat ***
    var elevesList = (data.eleves && data.eleves.eleves) ? data.eleves.eleves : (Array.isArray(data.eleves) ? data.eleves : []);
    var notesResults = Array.isArray(data.notes) ? data.notes : [];
    var absencesList = (data.absences && data.absences.absences) ? data.absences.absences : (Array.isArray(data.absences) ? data.absences : []);
    var observationsList = (data.observations && data.observations.observations) ? data.observations.observations : (Array.isArray(data.observations) ? data.observations : []);
    var punitionsList = (data.punitions && data.punitions.punitions) ? data.punitions.punitions : (Array.isArray(data.punitions) ? data.punitions : []);
    var incidentsList = (data.incidents && data.incidents.incidents) ? data.incidents.incidents : (Array.isArray(data.incidents) ? data.incidents : []);

    Logger.log('UNWRAP: eleves=' + elevesList.length + ' notesResults=' + notesResults.length +
      ' absences=' + absencesList.length + ' observations=' + observationsList.length +
      ' punitions=' + punitionsList.length + ' incidents=' + incidentsList.length);

    if (elevesList.length === 0) {
      return { success: false, error: 'Aucun eleve dans la liste. Collez d\'abord la liste eleves (etape 1).' };
    }

    // 1. CONSTRUIRE LA MAP ELEVES (cle = NOM|PRENOM)
    var studentMap = {};
    var classeGroups = {};

    for (var i = 0; i < elevesList.length; i++) {
      var e = elevesList[i];
      var classe = normaliserClasse_(e.classe);
      var cle = cleEleve_(e.nom, e.prenom);
      studentMap[cle] = {
        nom: e.nom, prenom: e.prenom, sexe: e.sexe || '', classe: classe,
        lv2: e.lv2 || '', opt: e.opt || '', dispo: e.dispo || '',
        moyennes: {}, oraux: {},
        dj: 0, nj: 0,
        nbObservations: 0, nbPunitions: 0, nbIncidents: 0, nbEncourage: 0
      };
      if (!classeGroups[classe]) classeGroups[classe] = [];
      classeGroups[classe].push(cle);
    }

    Logger.log('Eleves mappes: ' + Object.keys(studentMap).length);
    Logger.log('Classes: ' + Object.keys(classeGroups).join(', '));

    // [DIAG] Echantillon cles studentMap
    var smKeys = Object.keys(studentMap);
    Logger.log('[DIAG] Echantillon cles studentMap (5 premiers): ' + smKeys.slice(0, 5).join(', '));

    // 2. FUSIONNER LES NOTES (notesResults = [ {success, notes: [{nom,prenom,moyennes,oraux}], classe} ])
    var notesMatched = 0;
    var notesTotal = 0;
    var notesUnmatched = [];
    for (var nr = 0; nr < notesResults.length; nr++) {
      var noteResult = notesResults[nr];
      var noteList = (noteResult && noteResult.notes) ? noteResult.notes : (Array.isArray(noteResult) ? noteResult : []);
      notesTotal += noteList.length;

      // [DIAG] Echantillon noms source notes (3 premiers de chaque classe)
      if (noteList.length > 0) {
        var sampleNotes = noteList.slice(0, 3).map(function(nd) {
          return '"' + nd.nom + '|' + (nd.prenom || '') + '" -> norm:"' + normaliserNom_(nd.nom) + '|' + normaliserNom_(nd.prenom) + '"';
        });
        Logger.log('[DIAG] Notes classe ' + nr + ' echantillon: ' + sampleNotes.join(', '));
      }

      for (var n = 0; n < noteList.length; n++) {
        var noteData = noteList[n];
        var cle2 = findMatchingStudent_(studentMap, noteData.nom, noteData.prenom);
        if (cle2) {
          studentMap[cle2].moyennes = noteData.moyennes || {};
          studentMap[cle2].oraux = noteData.oraux || {};
          notesMatched++;
        } else {
          notesUnmatched.push(noteData.nom + ' ' + (noteData.prenom || ''));
        }
      }
    }
    Logger.log('Notes fusionnees: ' + notesMatched + '/' + notesTotal);
    if (notesUnmatched.length > 0) {
      Logger.log('[WARN] Notes non matchees (' + notesUnmatched.length + '): ' + notesUnmatched.slice(0, 10).join(', '));
    }

    // 3. FUSIONNER LES ABSENCES
    var absMatched = 0;
    for (var a = 0; a < absencesList.length; a++) {
      var absData = absencesList[a];
      var cle3 = findMatchingStudent_(studentMap, absData.nom, absData.prenom);
      if (cle3) {
        studentMap[cle3].dj = absData.dj || 0;
        studentMap[cle3].nj = absData.nj || 0;
        absMatched++;
      }
    }
    Logger.log('Absences fusionnees: ' + absMatched + '/' + absencesList.length);

    // 4. FUSIONNER PUNITIONS
    var punMatched = 0;
    for (var p = 0; p < punitionsList.length; p++) {
      var punData = punitionsList[p];
      var cle4 = findMatchingStudent_(studentMap, punData.nom, punData.prenom);
      if (cle4) {
        studentMap[cle4].nbPunitions += (punData.nb || 0);
        punMatched++;
      }
    }
    Logger.log('Punitions fusionnees: ' + punMatched + '/' + punitionsList.length);

    // 5. FUSIONNER INCIDENTS
    var incMatched = 0;
    for (var ic = 0; ic < incidentsList.length; ic++) {
      var incData = incidentsList[ic];
      var cleInc = findMatchingStudent_(studentMap, incData.nom, incData.prenom);
      if (cleInc) {
        studentMap[cleInc].nbIncidents += (incData.nb || 0);
        incMatched++;
      }
    }
    Logger.log('Incidents fusionnes: ' + incMatched + '/' + incidentsList.length);

    // 6. FUSIONNER OBSERVATIONS (feuilles d'appel : C+F+G+H negatif, D=encourage)
    var obsMatched = 0;
    for (var o = 0; o < observationsList.length; o++) {
      var obsData = observationsList[o];
      var cleObs = findMatchingStudent_(studentMap, obsData.nom, obsData.prenom);
      if (cleObs) {
        studentMap[cleObs].nbObservations += (obsData.nbObservationsNeg || 0);
        studentMap[cleObs].nbEncourage += (obsData.nbEncourage || 0);
        obsMatched++;
      }
    }
    Logger.log('Observations fusionnees: ' + obsMatched + '/' + observationsList.length);

    // 7. CALCULER LES SCORES + DIAGNOSTIC
    var importCfg = getImportScoringCfg_();
    Logger.log('[SCORING] Config source: ' + importCfg._source +
      ' | seuils presents: TRA=' + !!importCfg.seuils.TRA +
      ' PART=' + !!importCfg.seuils.PART +
      ' COM=' + !!importCfg.seuils.COM +
      ' ABS=' + !!importCfg.seuils.ABS);

    var scoresCount = 0;
    var diagTRA = { total: 0, null: 0, s1: 0, s2: 0, s3: 0, s4: 0, samples: [] };
    var diagPART = { total: 0, null: 0, s1: 0, s2: 0, s3: 0, s4: 0 };
    for (var cleS in studentMap) {
      var st = studentMap[cleS];
      st.scoreTRA = calcScoreTRA_import_(st.moyennes);
      st.scorePART = calcScorePART_import_(st.oraux);
      st.scoreABS = calcScoreABS_import_(st.dj, st.nj);
      st.scoreCOM = calcScoreCOM_import_(st.nbPunitions, st.nbIncidents, st.nbObservations, st.nbEncourage);
      if (st.scoreTRA || st.scorePART || st.scoreABS || st.scoreCOM) scoresCount++;

      // Diagnostic TRA
      diagTRA.total++;
      if (st.scoreTRA === null) { diagTRA.null++; }
      else { diagTRA['s' + st.scoreTRA]++; }
      // Garder 5 exemples detailles pour le debug
      if (diagTRA.samples.length < 5 && Object.keys(st.moyennes).length > 0) {
        diagTRA.samples.push(st.nom + ' ' + st.prenom + ': moy=' + JSON.stringify(st.moyennes) + ' -> TRA=' + st.scoreTRA);
      }
      // Diagnostic PART
      diagPART.total++;
      if (st.scorePART === null) { diagPART.null++; }
      else { diagPART['s' + st.scorePART]++; }
    }

    Logger.log('Scores calcules: ' + scoresCount);
    Logger.log('[DIAG] TRA repartition: null=' + diagTRA.null + ' score1=' + diagTRA.s1 +
      ' score2=' + diagTRA.s2 + ' score3=' + diagTRA.s3 + ' score4=' + diagTRA.s4);
    Logger.log('[DIAG] TRA exemples: ' + diagTRA.samples.join(' | '));
    Logger.log('[DIAG] PART repartition: null=' + diagPART.null + ' score1=' + diagPART.s1 +
      ' score2=' + diagPART.s2 + ' score3=' + diagPART.s3 + ' score4=' + diagPART.s4);

    // 8. ECRIRE DANS LES ONGLETS SOURCES
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

        studentRows.push([
          '', st2.nom, st2.prenom, '', st2.sexe, st2.lv2, st2.opt,
          st2.scoreCOM !== null ? String(st2.scoreCOM) : '',
          st2.scoreTRA !== null ? String(st2.scoreTRA) : '',
          st2.scorePART !== null ? String(st2.scorePART) : '',
          st2.scoreABS !== null ? String(st2.scoreABS) : '',
          st2.dispo || '', '', '', classeName
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

    // 9. LISTES DEROULANTES + FORMATAGE CONDITIONNEL
    try { ajouterListesDeroulantes(); } catch (e2) { Logger.log('ajouterListesDeroulantes: ' + e2.message); }

    // 10. NOM_PRENOM + IDs + CONSOLIDATION
    // FLUSH OBLIGATOIRE : les ecritures GAS sont bufferisees, sans flush
    // genererNomPrenomEtID() lirait des feuilles vides (donnees non persistees)
    SpreadsheetApp.flush();
    try { genererNomPrenomEtID(); } catch (e3) { Logger.log('genererNomPrenomEtID: ' + e3.message); }

    var consolResult = '';
    try { consolResult = consoliderDonnees(); } catch (e4) { consolResult = 'Non disponible'; }

    // 11. DISSO AUTO-SUGGESTION
    var dissoSuggestion = suggestDissoGroups_(studentMap, classeGroups);

    // 12. DETECTION ANOMALIES LV2/OPT
    var anomLV2 = 0, anomOPT = 0;
    for (var cleA in studentMap) {
      var stA = studentMap[cleA];
      if (!stA.lv2) anomLV2++;
      if (!stA.opt) anomOPT++;
    }
    if (anomLV2 > 0) RunAudit_log(runId, 'WARN', 'LV2 vide pour ' + anomLV2 + ' eleve(s)');
    if (anomOPT > 0) RunAudit_log(runId, 'WARN', 'OPT vide pour ' + anomOPT + ' eleve(s)');

    RunAudit_log(runId, 'INFO', '=== COMPILATION TERMINEE ===');

    var durationMs = RunAudit_stopTimer(timer);
    var runReport = RunAudit_buildReport({
      runId: runId, operation: 'IMPORT', durationMs: durationMs, success: true,
      totalEleves: elevesList.length,
      notesMatched: notesMatched, notesTotal: notesTotal, notesUnmatched: notesUnmatched,
      absMatched: absMatched, absTotal: absencesList.length,
      obsMatched: obsMatched, obsTotal: observationsList.length,
      punMatched: punMatched, punTotal: punitionsList.length,
      incMatched: incMatched, incTotal: incidentsList.length,
      classesList: Object.keys(classeGroups),
      anomaliesLV2: anomLV2, anomaliesOPT: anomOPT
    });

    var summary = {
      totalStudents: elevesList.length,
      classesProcessed: Object.keys(classeGroups).length,
      classesList: Object.keys(classeGroups),
      sheetsCreated: sheetsCreated, sheetsUpdated: sheetsUpdated,
      scoresCalculated: scoresCount,
      notesMatched: notesMatched, absMatched: absMatched,
      obsMatched: obsMatched, punMatched: punMatched, incMatched: incMatched,
      consolidation: consolResult
    };

    logAction('Import multi-paste: ' + elevesList.length + ' eleves dans ' + Object.keys(classeGroups).length + ' classes');
    return { success: true, summary: summary, runReport: runReport, dissoSuggestions: dissoSuggestion.groups || [] };

  } catch (e) {
    var durationMs2 = RunAudit_stopTimer(timer);
    RunAudit_buildReport({
      runId: runId, operation: 'IMPORT', durationMs: durationMs2,
      success: false, error: e.toString()
    });
    Logger.log('Erreur v3_compileImport: ' + e.toString());
    Logger.log(e.stack);
    return { success: false, error: e.toString(), runId: runId };
  }
}

// =============================================================================
// FONCTIONS DE MATCHING
// =============================================================================

function findMatchingStudent_(studentMap, nom, prenom) {
  // 1. Match exact par cle NOM|PRENOM
  var cle = cleEleve_(nom, prenom);
  if (studentMap[cle]) return cle;

  var normNom = normaliserNom_(nom);
  var normPrenom = normaliserNom_(prenom);

  // 2. Match normalise (accents, casse, guillemets, invisibles)
  for (var key in studentMap) {
    var st = studentMap[key];
    if (normaliserNom_(st.nom) === normNom && normaliserNom_(st.prenom) === normPrenom) return key;
  }

  // 2b. Fullname normalise : NOM+PRENOM concatenes sans separateur
  var srcFull = (normNom + normPrenom).replace(/\s+/g, '');
  if (srcFull.length >= 3) {
    for (var keyF in studentMap) {
      var stF = studentMap[keyF];
      var refFull0 = (normaliserNom_(stF.nom) + normaliserNom_(stF.prenom)).replace(/\s+/g, '');
      if (srcFull === refFull0) return keyF;
    }
  }

  // 3. Fullname = "NOM PRENOM" dans une seule cellule : tester tous les splits possibles
  //    Ex: "THERY BUDIN Lola" -> essayer "THERY|BUDIN LOLA", "THERY BUDIN|LOLA"
  var fullName = (nom + (prenom ? ' ' + prenom : '')).trim();
  var fullParts = fullName.split(/\s+/);
  if (fullParts.length >= 2) {
    for (var sp = 1; sp < fullParts.length; sp++) {
      var testNom = normaliserNom_(fullParts.slice(0, sp).join(' '));
      var testPrenom = normaliserNom_(fullParts.slice(sp).join(' '));
      for (var key2 in studentMap) {
        var st2 = studentMap[key2];
        if (normaliserNom_(st2.nom) === testNom && normaliserNom_(st2.prenom) === testPrenom) return key2;
      }
    }
  }

  // 4. Match par nom seul si unique dans la base
  var nomMatches = [];
  for (var key3 in studentMap) {
    if (normaliserNom_(studentMap[key3].nom) === normNom) nomMatches.push(key3);
  }
  if (nomMatches.length === 1) return nomMatches[0];

  // 5. Match partiel : le nom fourni CONTIENT le nom de la base ou vice-versa
  var fullNorm = normaliserNom_(fullName);
  for (var key4 in studentMap) {
    var st4 = studentMap[key4];
    var refFull = normaliserNom_(st4.nom) + ' ' + normaliserNom_(st4.prenom);
    if (fullNorm === refFull) return key4;
  }

  return null;
}

// =============================================================================
// CALCUL DES SCORES (VERSION IMPORT)
// =============================================================================

/**
 * Helper : resout la config scoring depuis getScoringConfig() ou fallback statique.
 * Ne throw jamais — retourne toujours un objet utilisable.
 */
function getImportScoringCfg_() {
  try {
    if (typeof getScoringConfig === 'function') {
      var cfg = getScoringConfig();
      if (cfg && cfg.seuils && cfg.seuils.TRA && cfg.seuils.ABS) {
        cfg._source = 'dynamique (getScoringConfig)';
        return cfg;
      }
    }
  } catch (e) {
    Logger.log('[WARN] getScoringConfig() a echoue: ' + e.toString());
  }
  // Fallback statique (memes seuils que SCORING_DEFAULTS)
  Logger.log('[WARN] Scoring: fallback statique (getScoringConfig indisponible)');
  return {
    _source: 'fallback statique',
    seuils: {
      TRA: [
        { score: 4, min: 15, max: 20 },
        { score: 3, min: 12, max: 14.999 },
        { score: 2, min: 8, max: 11.999 },
        { score: 1, min: 0, max: 7.999 }
      ],
      PART: [
        { score: 4, min: 15, max: 20 },
        { score: 3, min: 12, max: 14.999 },
        { score: 2, min: 8, max: 11.999 },
        { score: 1, min: 0, max: 7.999 }
      ],
      COM: [
        { score: 4, min: 0, max: 0 },
        { score: 3, min: 1, max: 5 },
        { score: 2, min: 6, max: 20 },
        { score: 1, min: 21, max: 999 }
      ],
      ABS: {
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
      }
    }
  };
}

function calcScoreTRA_import_(moyennes) {
  if (!moyennes || Object.keys(moyennes).length === 0) return null;
  var cfg = getImportScoringCfg_();
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
  return attribuerScoreParSeuil_(moy, cfg.seuils.TRA);
}

function calcScorePART_import_(oraux) {
  if (!oraux || Object.keys(oraux).length === 0) return null;
  var cfg = getImportScoringCfg_();
  var notes = [];
  for (var id in oraux) {
    if (oraux[id] !== null && oraux[id] !== undefined) notes.push(oraux[id]);
  }
  if (notes.length === 0) return null;
  var moy = notes.reduce(function(a, b) { return a + b; }, 0) / notes.length;
  moy = Math.round(moy * 100) / 100;
  return attribuerScoreParSeuil_(moy, cfg.seuils.PART);
}

function calcScoreABS_import_(dj, nj) {
  if (dj === 0 && nj === 0) return 4;
  var cfg = getImportScoringCfg_();
  var seuils = cfg.seuils.ABS;
  var scoreDJ = attribuerScoreParSeuil_(dj, seuils.DJ);
  var scoreNJ = attribuerScoreParSeuil_(nj, seuils.NJ);
  return Math.ceil(scoreDJ * seuils.poidsDJ + scoreNJ * seuils.poidsNJ);
}

function calcScoreCOM_import_(nbPunitions, nbIncidents, nbObservations, nbEncourage) {
  var cfg = getImportScoringCfg_();
  // Attenuation : encouragements reduisent le bloc observations (max 30%), jamais les punitions/incidents
  var obsNet = nbObservations || 0;
  if ((nbEncourage || 0) > 0 && obsNet > 0) {
    var attenuation = Math.min(nbEncourage, Math.floor(obsNet * 0.3));
    obsNet = obsNet - attenuation;
  }
  // Ponderation : incidents x3, punitions (incl. retenues) x2, observations nettes x0.5
  var total = (nbPunitions || 0) * 2 + (nbIncidents || 0) * 3 + Math.ceil(obsNet * 0.5);
  return attribuerScoreParSeuil_(total, cfg.seuils.COM);
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
        penibilite: penibilite, score: penibilite,
        details: { observations: st.nbObservations || 0, punitions: st.nbPunitions || 0, incidents: st.nbIncidents || 0 }
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

// =============================================================================
// LOCK PATTERN — Detection import deja fait + Purge nucleaire
// =============================================================================

/**
 * Verifie si un import a deja ete compile (onglets sources peuples).
 * Utilise par le client pour afficher le panneau verrouille au lieu des textareas.
 * @returns {Object} { locked, totalEleves, classes: [{name, count}], timestamp }
 */
function v3_getImportLockStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();
    var sourceSheets = allSheets.filter(function(s) { return /.+°\d+$/.test(s.getName()); });

    var totalEleves = 0;
    var classes = [];

    for (var i = 0; i < sourceSheets.length; i++) {
      var sheet = sourceSheets[i];
      var nbRows = Math.max(0, sheet.getLastRow() - 1);
      if (nbRows > 0) {
        classes.push({ name: sheet.getName(), count: nbRows });
        totalEleves += nbRows;
      }
    }

    // Recuperer le timestamp de la derniere compilation depuis _CONFIG PROGRESS
    var timestamp = '';
    try {
      var configSheet = ss.getSheetByName('_CONFIG');
      if (configSheet) {
        var data = configSheet.getDataRange().getValues();
        for (var r = 0; r < data.length; r++) {
          if (data[r][0] === 'PROGRESS') {
            var prog = JSON.parse(data[r][1] || '{}');
            if (prog.phase === 'scores' || prog.phase === 'import') {
              timestamp = prog.timestamp || '';
            }
            break;
          }
        }
      }
    } catch (e2) { /* ignore */ }

    return {
      success: true,
      locked: totalEleves > 0,
      totalEleves: totalEleves,
      classes: classes,
      timestamp: timestamp
    };
  } catch (e) {
    return { success: false, locked: false, error: e.toString() };
  }
}

/**
 * Purge nucleaire : vide tous les onglets sources + CONSOLIDATION.
 * Retablit l'etat "avant import" pour permettre un nouvel import propre.
 * @returns {Object} { success, purged: [noms des onglets purges] }
 */
function v3_purgeImport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();
    var sourceSheets = allSheets.filter(function(s) { return /.+°\d+$/.test(s.getName()); });
    var purged = [];

    // 1. Supprimer les onglets sources (classes)
    for (var i = 0; i < sourceSheets.length; i++) {
      var name = sourceSheets[i].getName();
      ss.deleteSheet(sourceSheets[i]);
      purged.push(name);
    }

    // 2. Vider CONSOLIDATION
    var consolidation = ss.getSheetByName('CONSOLIDATION');
    if (consolidation) {
      consolidation.clear();
      purged.push('CONSOLIDATION');
    }

    // 3. Retrograder la progression a "config" (avant import)
    try {
      var configSheet = ss.getSheetByName('_CONFIG');
      if (configSheet) {
        var data = configSheet.getDataRange().getValues();
        for (var r = 0; r < data.length; r++) {
          if (data[r][0] === 'PROGRESS') {
            var progressData = {
              phase: 'config',
              timestamp: new Date().toISOString(),
              metadata: { purgedAt: new Date().toISOString() }
            };
            configSheet.getRange(r + 1, 2).setValue(JSON.stringify(progressData));
            break;
          }
        }
      }
    } catch (e2) { Logger.log('Erreur reset progress: ' + e2.message); }

    Logger.log('[INFO] Purge import: ' + purged.length + ' onglets purges: ' + purged.join(', '));

    return {
      success: true,
      purged: purged,
      message: purged.length + ' onglets purges. Vous pouvez recommencer l\'import.'
    };
  } catch (e) {
    Logger.log('[ERROR] v3_purgeImport: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
