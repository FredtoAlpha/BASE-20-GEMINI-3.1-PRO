/**
 * ===================================================================
 * BACKEND_SCORES.JS - MODULE DE CALCUL DES SCORES ÉLÈVES
 * ===================================================================
 * Calcule 4 scores (ABS, COM, TRA, PART) à partir des exports Pronote
 * et les injecte dans les colonnes des onglets sources élèves.
 *
 * DÉTECTION DYNAMIQUE DES COLONNES :
 * Les en-têtes Pronote varient selon l'établissement (AGL1 MOY, FRANC,
 * HI-GE, ESP2 MOY, etc.). Ce module détecte les colonnes par pattern
 * matching sur la ligne d'en-tête au lieu d'utiliser des indices fixes.
 *
 * ARCHITECTURE DES ONGLETS PRONOTE :
 * - DATA_ABS       → Export Pronote des absences
 * - DATA_INCIDENTS → Export Pronote des incidents/sanctions
 * - DATA_PUNITIONS → Export Pronote des punitions
 * - DATA_NOTES     → Export Pronote des notes/moyennes
 *
 * @version 2.0.0 — Détection dynamique des colonnes
 * ===================================================================
 */

// =============================================================================
// CONFIGURATION DU MODULE SCORES
// =============================================================================

var SCORES_CONFIG = {
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

  // ── Matières pour le score TRA ──
  // patterns = regex appliqués sur les en-têtes Pronote pour trouver la colonne
  // preferMoy = true → si "MOY" et "ECRIT" existent, prendre MOY
  MATIERES_TRA: [
    { nom: 'Français',      patterns: ['FRANC', 'FRAN[CÇ]'], coeff: 4.5 },
    { nom: 'Maths',         patterns: ['MATH'], coeff: 3.5 },
    { nom: 'Histoire-Géo',  patterns: ['HI.?GE', 'HIST.*G[EÉ]O', 'HG'], coeff: 3.0 },
    { nom: 'Anglais',       patterns: ['ANG.*MOY', 'AGL.*MOY', 'ANGLAIS', 'ANG(?!.*(?:ORAL|ECRI))'], coeff: 3.0 },
    { nom: 'LV2',           patterns: ['ESP.*MOY', 'ALL.*MOY', 'ITA.*MOY', 'ESP[^O]*$', 'ALL[^O]*$', 'ITA[^O]*$'], coeff: 2.5 },
    { nom: 'EPS',           patterns: ['^EPS'], coeff: 2.0 },
    { nom: 'Phys.-Chimie',  patterns: ['PH.?CH', 'PHYS', 'SC.?PH'], coeff: 1.5, multi: true },
    { nom: 'SVT',           patterns: ['^SVT'], coeff: 1.5, multi: true },
    { nom: 'Technologie',   patterns: ['TECHN'], coeff: 1.5, multi: true },
    { nom: 'Arts Pla.',     patterns: ['A.?PLA', 'ARTS'], coeff: 1.0 },
    { nom: 'Musique',       patterns: ['EDMUS', 'MUS'], coeff: 1.0 },
    { nom: 'Latin',         patterns: ['LAT', 'LCALA'], coeff: 1.0 }
  ],

  // ── Patterns pour les colonnes ORAL (score PART) ──
  PATTERNS_ORAL_ANG: ['ANG.*ORAL', 'AGL.*ORAL', 'ORAL.*ANG'],
  PATTERNS_ORAL_LV2: ['ESP.*ORAL', 'ALL.*ORAL', 'ITA.*ORAL', 'ORAL.*LV2'],

  // ── Patterns pour DATA_ABS ──
  PATTERNS_ABS: {
    nom:       ['NOM'],
    classe:    ['CLASSE'],
    dj:        ['DJ', 'DEMI.?JOURN', 'DJ.*BULL'],
    justifiee: ['JUSTIFI']
  },

  // ── Patterns pour DATA_INCIDENTS ──
  PATTERNS_INC: {
    nom:     ['NOM'],
    classe:  ['CLASSE'],
    gravite: ['GRAVIT', 'GRAV']
  },

  // ── Patterns pour DATA_PUNITIONS ──
  PATTERNS_PUN: {
    nom:    ['NOM'],
    classe: ['CLASSE'],
    nb:     ['^NB', 'NOMBRE', 'QT', 'QUANT']
  }
};

// =============================================================================
// DÉTECTION DYNAMIQUE DES COLONNES
// =============================================================================

/**
 * Cherche l'indice (0-based) de la première colonne dont l'en-tête
 * matche l'un des patterns fournis.
 * @param {string[]} headers — ligne d'en-têtes normalisée (uppercase, trimmed)
 * @param {string[]} patterns — liste de regex patterns à tester
 * @returns {number} indice 0-based, ou -1 si non trouvé
 */
function findCol_(headers, patterns) {
  for (var p = 0; p < patterns.length; p++) {
    var re = new RegExp(patterns[p], 'i');
    for (var c = 0; c < headers.length; c++) {
      if (re.test(headers[c])) return c;
    }
  }
  return -1;
}

/**
 * Cherche TOUS les indices de colonnes matchant les patterns.
 * Utile pour les matières à groupes (Techno G1/G2, SVT G1/G2, etc.)
 * @param {string[]} headers
 * @param {string[]} patterns
 * @returns {number[]} tableau d'indices 0-based
 */
function findAllCols_(headers, patterns) {
  var found = [];
  for (var p = 0; p < patterns.length; p++) {
    var re = new RegExp(patterns[p], 'i');
    for (var c = 0; c < headers.length; c++) {
      if (re.test(headers[c]) && found.indexOf(c) === -1) {
        found.push(c);
      }
    }
  }
  return found;
}

/**
 * Normalise une ligne d'en-têtes : uppercase + trim.
 * Scanne les 2 premières lignes de données pour trouver celle
 * qui ressemble le plus à des en-têtes (texte, pas des nombres).
 * @param {Array[]} data — toutes les données de la feuille
 * @returns {{ headers: string[], dataStartRow: number }}
 */
function detectHeaders_(data) {
  if (!data || data.length === 0) return { headers: [], dataStartRow: 0 };

  // Heuristique : la ligne d'en-tête est celle avec le plus de cellules texte
  var bestRow = 0;
  var bestTextCount = 0;

  var maxScan = Math.min(data.length, 3);
  for (var r = 0; r < maxScan; r++) {
    var textCount = 0;
    for (var c = 0; c < data[r].length; c++) {
      var val = String(data[r][c]).trim();
      if (val && isNaN(val) && val !== 'Abs' && val !== 'Disp') textCount++;
    }
    if (textCount > bestTextCount) {
      bestTextCount = textCount;
      bestRow = r;
    }
  }

  var headers = [];
  for (var c = 0; c < data[bestRow].length; c++) {
    headers.push(String(data[bestRow][c]).trim().toUpperCase());
  }

  return {
    headers: headers,
    dataStartRow: bestRow + 1 // données commencent après la ligne d'en-tête
  };
}

// =============================================================================
// FONCTIONS SERVEUR V3 — Adaptateurs pour Console Pilotage V3
// =============================================================================

/**
 * Initialise les onglets DATA_* pour recevoir les exports Pronote.
 * NE pré-écrit PAS de colonnes : l'utilisateur colle l'export tel quel.
 * Le moteur détecte dynamiquement les colonnes par leurs en-têtes.
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

    var instructions = {
      'DATA_ABS':
        '📋 ABSENCES — Collez ici l\'export Pronote complet (avec ses en-têtes). ' +
        'Colonnes attendues : Nom, Classe, Demi-journées (DJ), Justifiée.',
      'DATA_INCIDENTS':
        '📋 INCIDENTS — Collez ici l\'export Pronote complet (avec ses en-têtes). ' +
        'Colonnes attendues : Nom, Classe, Gravité.',
      'DATA_PUNITIONS':
        '📋 PUNITIONS — Collez ici l\'export Pronote complet (avec ses en-têtes). ' +
        'Colonnes attendues : Nb, Nom, Classe.',
      'DATA_NOTES':
        '📋 NOTES — Collez ici l\'export Pronote complet (avec ses en-têtes). ' +
        'Le moteur détecte automatiquement les matières par leurs intitulés (FRANC, MATH, AGL1 MOY, etc.).'
    };

    var instrBg = '#e8eaf6';
    var instrColor = '#283593';

    for (var nom in instructions) {
      var ws = ss.getSheetByName(nom);
      if (!ws) continue;
      if (ws.getLastRow() > 1) continue; // ne pas écraser si données déjà présentes

      ws.getRange('A1').setValue(instructions[nom]);
      ws.getRange('A1')
        .setFontStyle('italic').setFontColor(instrColor)
        .setBackground(instrBg).setFontSize(11)
        .setWrap(true);
      ws.setColumnWidth(1, 800);
    }

    return {
      success: true,
      message: created.length > 0
        ? 'Onglets créés : ' + created.join(', ') +
          '\nCollez les exports Pronote tels quels — le moteur détecte les colonnes automatiquement.'
        : 'Tous les onglets DATA existent déjà.',
      tabs: onglets.map(function(nom) {
        var ws = ss.getSheetByName(nom);
        return {
          name: nom,
          rows: ws ? Math.max(0, ws.getLastRow() - 1) : 0,
          hasData: ws ? ws.getLastRow() > 1 : false
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

    var absResults = calculerScoreABS_(ss);
    var comResults = calculerScoreCOM_(ss);
    var traResults = calculerScoreTRA_(ss);
    var partResults = calculerScorePART_(ss);

    var fusion = fusionnerScores_(absResults, comResults, traResults, partResults);

    // Lister les onglets sources pour diagnostic
    var allSheetNames = ss.getSheets().map(function(s) { return s.getName(); });
    var sourceSheetNames = allSheetNames.filter(function(n) { return /.+°\d+$/.test(n); });
    Logger.log('Onglets sources trouvés pour injection: ' + sourceSheetNames.join(', '));

    var injected = injecterScoresDansOngletsSources_(ss, fusion);

    Logger.log('=== RÉSULTAT INJECTION: ' + injected.updated + ' mis à jour, ' + injected.notFound + ' non trouvés ===');

    // Construire le tableau détaillé pour affichage immédiat côté client
    // Scores bruts 1-4 (identiques à ce qui est injecté dans les onglets sources)
    var preview = [];
    for (var nom in fusion) {
      var e = fusion[nom];
      preview.push({
        nom: nom,
        classe: e.classe,
        abs: e.scoreABS,
        com: e.scoreCOM,
        tra: e.scoreTRA,
        part: e.scorePART
      });
    }
    preview.sort(function(a, b) {
      return (a.classe || '').localeCompare(b.classe || '') || a.nom.localeCompare(b.nom);
    });

    return {
      success: true,
      results: {
        abs: { count: absResults.length, ok: true },
        com: { count: comResults.length, ok: true },
        tra: { count: traResults.length, ok: true },
        part: { count: partResults.length, ok: true }
      },
      injected: injected,
      totalEleves: Object.keys(fusion).length,
      preview: preview,
      debug: {
        sourceSheets: sourceSheetNames,
        fusionKeys: Object.keys(fusion).slice(0, 5),
        fusionCount: Object.keys(fusion).length
      }
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
      results: results.slice(0, 20)
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

    // Scores bruts 1-4 (identiques à ce qui est dans les onglets sources)
    var preview = [];
    for (var nom in fusion) {
      var e = fusion[nom];
      preview.push({
        nom: nom,
        classe: e.classe,
        abs: e.scoreABS,
        com: e.scoreCOM,
        tra: e.scoreTRA,
        part: e.scorePART
      });
    }

    preview.sort(function(a, b) {
      return (a.classe || '').localeCompare(b.classe || '') || a.nom.localeCompare(b.nom);
    });

    return {
      success: true,
      totalEleves: preview.length,
      preview: preview
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// MODULE ABS — Score d'assiduité (détection dynamique)
// =============================================================================

function calculerScoreABS_(ss) {
  var wsData = ss.getSheetByName('DATA_ABS');
  if (!wsData || wsData.getLastRow() < 2) return [];

  var data = wsData.getDataRange().getValues();
  var det = detectHeaders_(data);
  var h = det.headers;
  if (h.length === 0) return [];

  var pats = SCORES_CONFIG.PATTERNS_ABS;
  var colNom     = findCol_(h, pats.nom);
  var colClasse  = findCol_(h, pats.classe);
  var colDJ      = findCol_(h, pats.dj);
  var colJust    = findCol_(h, pats.justifiee);

  if (colNom === -1) {
    Logger.log('DATA_ABS: colonne NOM introuvable dans: ' + h.join(' | '));
    return [];
  }

  var seuils = SCORES_CONFIG.SEUILS_ABS;
  var eleves = {};

  for (var i = det.dataStartRow; i < data.length; i++) {
    var nom = String(data[i][colNom]).trim();
    if (!nom) continue;
    var classe = colClasse >= 0 ? String(data[i][colClasse]).trim() : '';
    var djVal  = colDJ >= 0 ? parseNotePronote_(data[i][colDJ]) : null;
    var justif = colJust >= 0 ? String(data[i][colJust]).trim() : '';

    if (!eleves[nom]) {
      eleves[nom] = { classe: classe, djTotal: 0, nonJustifiees: 0 };
    }
    if (djVal !== null) eleves[nom].djTotal += djVal;
    if (justif.toUpperCase() === 'NON') eleves[nom].nonJustifiees++;
  }

  var resultats = [];
  for (var nomKey in eleves) {
    var e = eleves[nomKey];
    var scoreDJ = attribuerScoreParSeuil_(e.djTotal, seuils.DJ);
    var scoreNJ = attribuerScoreParSeuil_(e.nonJustifiees, seuils.NJ);
    var scoreABS = Math.ceil(scoreDJ * seuils.poidsDJ + scoreNJ * seuils.poidsNJ);

    resultats.push({
      nom: nomKey, classe: e.classe,
      dj: Math.round(e.djTotal * 10) / 10,
      nj: e.nonJustifiees,
      scoreABS: scoreABS
    });
  }

  return resultats;
}

// =============================================================================
// MODULE COM — Score de comportement (détection dynamique)
// =============================================================================

function calculerScoreCOM_(ss) {
  var wsInc = ss.getSheetByName('DATA_INCIDENTS');
  var wsPun = ss.getSheetByName('DATA_PUNITIONS');

  if ((!wsInc || wsInc.getLastRow() < 2) && (!wsPun || wsPun.getLastRow() < 2)) return [];

  var seuils = SCORES_CONFIG.SEUILS_COM;

  // ── Punitions ──
  var punitions = {};
  if (wsPun && wsPun.getLastRow() >= 2) {
    var dataPun = wsPun.getDataRange().getValues();
    var detPun = detectHeaders_(dataPun);
    var hPun = detPun.headers;
    var pPun = SCORES_CONFIG.PATTERNS_PUN;
    var colPNom    = findCol_(hPun, pPun.nom);
    var colPClasse = findCol_(hPun, pPun.classe);
    var colPNb     = findCol_(hPun, pPun.nb);

    if (colPNom >= 0) {
      for (var i = detPun.dataStartRow; i < dataPun.length; i++) {
        var nom = String(dataPun[i][colPNom]).trim();
        if (!nom) continue;
        var nb = colPNb >= 0 ? (parseInt(dataPun[i][colPNb]) || 0) : 1;
        var classe = colPClasse >= 0 ? String(dataPun[i][colPClasse]).trim() : '';
        if (!punitions[nom]) punitions[nom] = { nb: 0, classe: '' };
        punitions[nom].nb += nb;
        if (classe) punitions[nom].classe = classe;
      }
    }
  }

  // ── Incidents ──
  var incidents = {};
  if (wsInc && wsInc.getLastRow() >= 2) {
    var dataInc = wsInc.getDataRange().getValues();
    var detInc = detectHeaders_(dataInc);
    var hInc = detInc.headers;
    var pInc = SCORES_CONFIG.PATTERNS_INC;
    var colINom    = findCol_(hInc, pInc.nom);
    var colIClasse = findCol_(hInc, pInc.classe);
    var colIGrav   = findCol_(hInc, pInc.gravite);

    if (colINom >= 0) {
      for (var i = detInc.dataStartRow; i < dataInc.length; i++) {
        var nom = String(dataInc[i][colINom]).trim();
        if (!nom) continue;
        var classe = colIClasse >= 0 ? String(dataInc[i][colIClasse]).trim() : '';
        var grav = 1;
        if (colIGrav >= 0) {
          var gravStr = String(dataInc[i][colIGrav]).trim();
          if (gravStr && gravStr.indexOf('/') > -1) {
            grav = parseInt(gravStr.split('/')[0]) || 1;
          } else {
            grav = parseInt(gravStr) || 1;
          }
        }

        if (!incidents[nom]) {
          incidents[nom] = { classe: '', nbInc: 0, ptsGrav: 0 };
        }
        if (classe) incidents[nom].classe = classe;
        incidents[nom].nbInc++;
        incidents[nom].ptsGrav += grav;
      }
    }
  }

  // Fusionner punitions + incidents
  var tousNoms = {};
  for (var nom in punitions) tousNoms[nom] = true;
  for (var nom in incidents) tousNoms[nom] = true;

  var resultats = [];
  for (var nomKey in tousNoms) {
    var ptsPun = punitions[nomKey] ? punitions[nomKey].nb : 0;
    var ptsInc = incidents[nomKey] ? incidents[nomKey].ptsGrav * 3 : 0;
    var total = ptsPun + ptsInc;
    var classe = (punitions[nomKey] ? punitions[nomKey].classe : '') ||
                 (incidents[nomKey] ? incidents[nomKey].classe : '');
    var scoreCOM = attribuerScoreParSeuil_(total, seuils);

    resultats.push({ nom: nomKey, classe: classe, total: total, scoreCOM: scoreCOM });
  }

  return resultats;
}

// =============================================================================
// MODULE TRA — Score de travail (détection dynamique des matières)
// =============================================================================

function calculerScoreTRA_(ss) {
  var wsData = ss.getSheetByName('DATA_NOTES');
  if (!wsData || wsData.getLastRow() < 2) return [];

  var data = wsData.getDataRange().getValues();
  var det = detectHeaders_(data);
  var h = det.headers;
  if (h.length === 0) return [];

  // Trouver Nom et Classe
  var colNom    = findCol_(h, ['NOM']);
  var colClasse = findCol_(h, ['CLASSE']);
  if (colNom === -1) {
    Logger.log('DATA_NOTES: colonne NOM introuvable dans: ' + h.join(' | '));
    return [];
  }

  // Résoudre dynamiquement les colonnes de chaque matière
  var matieresConf = SCORES_CONFIG.MATIERES_TRA;
  var matieresResolues = [];
  var matieresManquantes = [];

  for (var m = 0; m < matieresConf.length; m++) {
    var conf = matieresConf[m];
    var cols;
    if (conf.multi) {
      cols = findAllCols_(h, conf.patterns);
    } else {
      var idx = findCol_(h, conf.patterns);
      cols = idx >= 0 ? [idx] : [];
    }
    if (cols.length > 0) {
      matieresResolues.push({ nom: conf.nom, cols: cols, coeff: conf.coeff });
    } else {
      matieresManquantes.push(conf.nom);
    }
  }

  if (matieresManquantes.length > 0) {
    Logger.log('DATA_NOTES: matières non trouvées: ' + matieresManquantes.join(', ') +
               ' | En-têtes: ' + h.join(' | '));
  }

  if (matieresResolues.length === 0) {
    Logger.log('DATA_NOTES: aucune matière détectée — abandon calcul TRA');
    return [];
  }

  var seuils = SCORES_CONFIG.SEUILS_TRA;
  var resultats = [];

  for (var i = det.dataStartRow; i < data.length; i++) {
    var nom = String(data[i][colNom]).trim();
    if (!nom) continue;
    var classe = colClasse >= 0 ? String(data[i][colClasse]).trim() : '';

    var totalPts = 0;
    var totalCoeff = 0;

    for (var mi = 0; mi < matieresResolues.length; mi++) {
      var mat = matieresResolues[mi];
      var note = null;
      for (var c = 0; c < mat.cols.length; c++) {
        var colIdx = mat.cols[c];
        if (colIdx < data[i].length) {
          var n = parseNotePronote_(data[i][colIdx]);
          if (n !== null) { note = n; break; }
        }
      }
      if (note !== null) {
        totalPts += note * mat.coeff;
        totalCoeff += mat.coeff;
      }
    }

    var moyPond = totalCoeff > 0 ? Math.round(totalPts / totalCoeff * 100) / 100 : null;
    var scoreTRA = moyPond !== null ? attribuerScoreParSeuil_(moyPond, seuils) : null;

    resultats.push({ nom: nom, classe: classe, moyPond: moyPond, scoreTRA: scoreTRA });
  }

  return resultats;
}

// =============================================================================
// MODULE PART — Score de participation orale (détection dynamique)
// =============================================================================

function calculerScorePART_(ss) {
  var wsData = ss.getSheetByName('DATA_NOTES');
  if (!wsData || wsData.getLastRow() < 2) return [];

  var data = wsData.getDataRange().getValues();
  var det = detectHeaders_(data);
  var h = det.headers;
  if (h.length === 0) return [];

  var colNom    = findCol_(h, ['NOM']);
  var colClasse = findCol_(h, ['CLASSE']);
  if (colNom === -1) return [];

  // Trouver les colonnes ORAL
  var colOralAng = findCol_(h, SCORES_CONFIG.PATTERNS_ORAL_ANG);
  var colsOralLV2 = findAllCols_(h, SCORES_CONFIG.PATTERNS_ORAL_LV2);

  if (colOralAng === -1 && colsOralLV2.length === 0) {
    Logger.log('DATA_NOTES: aucune colonne ORAL trouvée — abandon calcul PART');
    return [];
  }

  var seuils = SCORES_CONFIG.SEUILS_PART;
  var resultats = [];

  for (var i = det.dataStartRow; i < data.length; i++) {
    var nom = String(data[i][colNom]).trim();
    if (!nom) continue;
    var classe = colClasse >= 0 ? String(data[i][colClasse]).trim() : '';

    var notes = [];
    if (colOralAng >= 0) {
      var oAng = parseNotePronote_(data[i][colOralAng]);
      if (oAng !== null) notes.push(oAng);
    }
    for (var lv = 0; lv < colsOralLV2.length; lv++) {
      var oLV2 = parseNotePronote_(data[i][colsOralLV2[lv]]);
      if (oLV2 !== null) { notes.push(oLV2); break; } // première LV2 trouvée
    }

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
  var allSheets = ss.getSheets();
  var sheets = allSheets.filter(function(s) {
    return /.+°\d+$/.test(s.getName());
  });

  Logger.log('=== INJECTION SCORES ===');
  Logger.log('Tous les onglets: ' + allSheets.map(function(s) { return s.getName(); }).join(', '));
  Logger.log('Onglets sources détectés (pattern °digit): ' + sheets.map(function(s) { return s.getName(); }).join(', '));
  Logger.log('Nombre élèves dans fusion: ' + Object.keys(fusion).length);
  if (Object.keys(fusion).length > 0) {
    var premierNom = Object.keys(fusion)[0];
    Logger.log('Exemple fusion: "' + premierNom + '" => ' + JSON.stringify(fusion[premierNom]));
  }

  var totalUpdated = 0;
  var totalNotFound = 0;

  sheets.forEach(function(sheet) {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var headers = data[0];
    // Normaliser les en-têtes : trim + majuscules pour éviter les espaces invisibles
    var headersNorm = headers.map(function(h) { return String(h).trim().toUpperCase(); });

    var idxNom = headersNorm.indexOf('NOM');
    var idxPrenom = headersNorm.indexOf('PRENOM');
    var idxNomPrenom = headersNorm.indexOf('NOM_PRENOM');
    var idxCOM = headersNorm.indexOf('COM');
    var idxTRA = headersNorm.indexOf('TRA');
    var idxPART = headersNorm.indexOf('PART');
    var idxABS = headersNorm.indexOf('ABSENCE');
    if (idxABS === -1) idxABS = headersNorm.indexOf('ABS');

    Logger.log('Onglet ' + sheet.getName() + ' — headers bruts: [' + headers.join(' | ') + ']');
    Logger.log('  idx: NOM=' + idxNom + ' PRENOM=' + idxPrenom + ' NOM_PRENOM=' + idxNomPrenom + ' COM=' + idxCOM + ' TRA=' + idxTRA + ' PART=' + idxPART + ' ABS=' + idxABS);
    Logger.log('  Lignes de données: ' + (data.length - 1));

    if (idxCOM === -1 && idxTRA === -1 && idxPART === -1 && idxABS === -1) {
      Logger.log('  ⚠️ SKIP: aucune colonne score trouvée !');
      return;
    }

    var sheetUpdated = 0;
    var sheetNotFound = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && (idxNom === -1 || !row[idxNom])) continue;

      var nomPrenom = idxNomPrenom >= 0 ? String(row[idxNomPrenom]).trim() : '';
      var nom = idxNom >= 0 ? String(row[idxNom]).trim() : '';
      var prenom = idxPrenom >= 0 ? String(row[idxPrenom]).trim() : '';

      // Essayer plusieurs stratégies de matching
      var match = null;
      var matchKey = '';

      // 1) Match exact NOM_PRENOM
      if (nomPrenom && fusion[nomPrenom]) {
        match = fusion[nomPrenom];
        matchKey = 'NOM_PRENOM exact';
      }
      // 2) Match exact NOM seul
      if (!match && nom && fusion[nom]) {
        match = fusion[nom];
        matchKey = 'NOM exact';
      }
      // 3) Match "NOM PRENOM" concaténé
      if (!match) {
        var fullName = (nom + ' ' + prenom).trim();
        if (fullName && fusion[fullName]) {
          match = fusion[fullName];
          matchKey = 'NOM+PRENOM';
        }
      }
      // 4) Match "PRENOM NOM" (format inverse)
      if (!match) {
        var reverseName = (prenom + ' ' + nom).trim();
        if (reverseName && fusion[reverseName]) {
          match = fusion[reverseName];
          matchKey = 'PRENOM+NOM inverse';
        }
      }
      // 5) Recherche insensible à la casse dans les clés de fusion
      if (!match && nomPrenom) {
        var npUpper = nomPrenom.toUpperCase();
        for (var key in fusion) {
          if (key.toUpperCase() === npUpper) {
            match = fusion[key];
            matchKey = 'NOM_PRENOM case-insensitive';
            break;
          }
        }
      }

      if (match) {
        var rowNum = i + 1;
        var wrote = [];
        if (idxCOM >= 0 && match.scoreCOM !== undefined && match.scoreCOM !== null) {
          sheet.getRange(rowNum, idxCOM + 1).setValue(String(match.scoreCOM));
          wrote.push('COM=' + match.scoreCOM);
        }
        if (idxTRA >= 0 && match.scoreTRA !== undefined && match.scoreTRA !== null) {
          sheet.getRange(rowNum, idxTRA + 1).setValue(String(match.scoreTRA));
          wrote.push('TRA=' + match.scoreTRA);
        }
        if (idxPART >= 0 && match.scorePART !== undefined && match.scorePART !== null) {
          sheet.getRange(rowNum, idxPART + 1).setValue(String(match.scorePART));
          wrote.push('PART=' + match.scorePART);
        }
        if (idxABS >= 0 && match.scoreABS !== undefined && match.scoreABS !== null) {
          sheet.getRange(rowNum, idxABS + 1).setValue(String(match.scoreABS));
          wrote.push('ABS=' + match.scoreABS);
        }
        if (sheetUpdated < 3) {
          Logger.log('  ✅ Ligne ' + rowNum + ' match (' + matchKey + '): ' + (nomPrenom || nom) + ' → ' + wrote.join(', '));
        }
        sheetUpdated++;
        totalUpdated++;
      } else {
        if (sheetNotFound < 5) {
          Logger.log('  ❌ Ligne ' + (i+1) + ' PAS TROUVÉ: nom="' + nom + '" prenom="' + prenom + '" nom_prenom="' + nomPrenom + '"');
        }
        sheetNotFound++;
        totalNotFound++;
      }
    }

    Logger.log('  → ' + sheet.getName() + ': ' + sheetUpdated + ' mis à jour, ' + sheetNotFound + ' non trouvés');
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
  if (score14 === null || score14 === undefined) return 2.5;
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
