/**
 * ===================================================================
 * SCORING_MATIERES.JS — Matières et coefficients par niveau
 * ===================================================================
 *
 * Définit les matières, leurs patterns Pronote et coefficients
 * pour chaque niveau scolaire (6e, 5e, 4e, 3e).
 *
 * En 6e : pas de LV2, pas de Phys-Chimie
 * En 5e+ : ajout LV2, Phys-Chimie, SVT
 * En 4e/3e : coefficients pondérés pour le brevet
 *
 * @version 1.0.0
 * ===================================================================
 */

// =============================================================================
// MATIÈRES PAR NIVEAU — Score TRA (travail)
// =============================================================================

var MATIERES_PAR_NIVEAU = {
  '6e': [
    { nom: 'Francais',     patterns: ['FRANC', 'FRAN[CÇ]'], coeff: 5.0 },
    { nom: 'Maths',        patterns: ['MATH'], coeff: 4.5 },
    { nom: 'Histoire-Geo', patterns: ['HI.?GE', 'HIST.*G[EÉ]O', 'HG'], coeff: 3.0 },
    { nom: 'Anglais',      patterns: ['ANG.*MOY', 'AGL.*MOY', 'ANGLAIS', 'ANG(?!.*(?:ORAL|ECRI))'], coeff: 3.0 },
    { nom: 'EPS',          patterns: ['^EPS'], coeff: 2.0 },
    { nom: 'Technologie',  patterns: ['TECHN'], coeff: 1.5, multi: true },
    { nom: 'SVT',          patterns: ['^SVT'], coeff: 1.5, multi: true },
    { nom: 'Arts Pla.',    patterns: ['A.?PLA', 'ARTS'], coeff: 1.0 },
    { nom: 'Musique',      patterns: ['EDMUS', 'MUS'], coeff: 1.0 }
    // Pas de LV2, pas de Phys-Chimie en 6e
  ],

  '5e': [
    { nom: 'Francais',      patterns: ['FRANC', 'FRAN[CÇ]'], coeff: 4.5 },
    { nom: 'Maths',         patterns: ['MATH'], coeff: 3.5 },
    { nom: 'Histoire-Geo',  patterns: ['HI.?GE', 'HIST.*G[EÉ]O', 'HG'], coeff: 3.0 },
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

  '4e': [
    { nom: 'Francais',      patterns: ['FRANC', 'FRAN[CÇ]'], coeff: 4.5 },
    { nom: 'Maths',         patterns: ['MATH'], coeff: 4.0 },
    { nom: 'Histoire-Geo',  patterns: ['HI.?GE', 'HIST.*G[EÉ]O', 'HG'], coeff: 3.0 },
    { nom: 'Anglais',       patterns: ['ANG.*MOY', 'AGL.*MOY', 'ANGLAIS', 'ANG(?!.*(?:ORAL|ECRI))'], coeff: 3.0 },
    { nom: 'LV2',           patterns: ['ESP.*MOY', 'ALL.*MOY', 'ITA.*MOY', 'ESP[^O]*$', 'ALL[^O]*$', 'ITA[^O]*$'], coeff: 2.5 },
    { nom: 'EPS',           patterns: ['^EPS'], coeff: 2.0 },
    { nom: 'Phys.-Chimie',  patterns: ['PH.?CH', 'PHYS', 'SC.?PH'], coeff: 2.0, multi: true },
    { nom: 'SVT',           patterns: ['^SVT'], coeff: 1.5, multi: true },
    { nom: 'Technologie',   patterns: ['TECHN'], coeff: 1.5, multi: true },
    { nom: 'Arts Pla.',     patterns: ['A.?PLA', 'ARTS'], coeff: 1.0 },
    { nom: 'Musique',       patterns: ['EDMUS', 'MUS'], coeff: 1.0 },
    { nom: 'Latin',         patterns: ['LAT', 'LCALA'], coeff: 1.5 },
    { nom: 'Grec',          patterns: ['GREC'], coeff: 1.0 }
  ],

  '3e': [
    { nom: 'Francais',      patterns: ['FRANC', 'FRAN[CÇ]'], coeff: 5.0 },
    { nom: 'Maths',         patterns: ['MATH'], coeff: 5.0 },
    { nom: 'Histoire-Geo',  patterns: ['HI.?GE', 'HIST.*G[EÉ]O', 'HG'], coeff: 3.5 },
    { nom: 'Anglais',       patterns: ['ANG.*MOY', 'AGL.*MOY', 'ANGLAIS', 'ANG(?!.*(?:ORAL|ECRI))'], coeff: 3.0 },
    { nom: 'LV2',           patterns: ['ESP.*MOY', 'ALL.*MOY', 'ITA.*MOY', 'ESP[^O]*$', 'ALL[^O]*$', 'ITA[^O]*$'], coeff: 2.5 },
    { nom: 'EPS',           patterns: ['^EPS'], coeff: 2.0 },
    { nom: 'Phys.-Chimie',  patterns: ['PH.?CH', 'PHYS', 'SC.?PH'], coeff: 2.5, multi: true },
    { nom: 'SVT',           patterns: ['^SVT'], coeff: 2.0, multi: true },
    { nom: 'Technologie',   patterns: ['TECHN'], coeff: 1.5, multi: true },
    { nom: 'Arts Pla.',     patterns: ['A.?PLA', 'ARTS'], coeff: 1.0 },
    { nom: 'Musique',       patterns: ['EDMUS', 'MUS'], coeff: 1.0 },
    { nom: 'Latin',         patterns: ['LAT', 'LCALA'], coeff: 2.0 },
    { nom: 'Grec',          patterns: ['GREC'], coeff: 1.0 }
  ]
};

// =============================================================================
// PATTERNS ORAL PAR NIVEAU — Score PART (participation)
// =============================================================================

var PATTERNS_ORAL_PAR_NIVEAU = {
  '6e': {
    ang: ['ANG.*ORAL', 'AGL.*ORAL', 'ORAL.*ANG'],
    lv2: [] // pas de LV2 en 6e
  },
  '5e': {
    ang: ['ANG.*ORAL', 'AGL.*ORAL', 'ORAL.*ANG'],
    lv2: ['ESP.*ORAL', 'ALL.*ORAL', 'ITA.*ORAL', 'ORAL.*LV2']
  },
  '4e': {
    ang: ['ANG.*ORAL', 'AGL.*ORAL', 'ORAL.*ANG'],
    lv2: ['ESP.*ORAL', 'ALL.*ORAL', 'ITA.*ORAL', 'ORAL.*LV2']
  },
  '3e': {
    ang: ['ANG.*ORAL', 'AGL.*ORAL', 'ORAL.*ANG'],
    lv2: ['ESP.*ORAL', 'ALL.*ORAL', 'ITA.*ORAL', 'ORAL.*LV2']
  }
};

// =============================================================================
// API PUBLIQUE
// =============================================================================

/**
 * Retourne les matières et coefficients pour un niveau donné.
 * Fallback: retourne le niveau 5e si le niveau est inconnu.
 *
 * @param {string} niveau - '6e', '5e', '4e' ou '3e'
 * @returns {Array} Tableau de { nom, patterns, coeff, multi? }
 */
function getMatieresForLevel(niveau) {
  var key = normalizeNiveau_(niveau);
  return MATIERES_PAR_NIVEAU[key] || MATIERES_PAR_NIVEAU['5e'];
}

/**
 * Retourne les patterns ORAL pour un niveau donné.
 *
 * @param {string} niveau - '6e', '5e', '4e' ou '3e'
 * @returns {Object} { ang: string[], lv2: string[] }
 */
function getOralPatternsForLevel(niveau) {
  var key = normalizeNiveau_(niveau);
  return PATTERNS_ORAL_PAR_NIVEAU[key] || PATTERNS_ORAL_PAR_NIVEAU['5e'];
}

/**
 * Détecte le niveau depuis un nom d'onglet source.
 * Exemples : "5°2" → "5e", "6°1" → "6e", "4°3CACHE" → "4e"
 *
 * @param {string} sheetName - Nom de l'onglet
 * @returns {string} '6e', '5e', '4e', '3e' ou '5e' par défaut
 */
function detectNiveauFromSheetName(sheetName) {
  if (!sheetName) return '5e';
  var match = String(sheetName).match(/^(\d)/);
  if (match) {
    var num = match[1];
    if (num === '6' || num === '5' || num === '4' || num === '3') {
      return num + 'e';
    }
  }
  return '5e';
}

/**
 * Détecte le niveau depuis la configuration globale ou les onglets sources.
 * Stratégie :
 * 1. Config._CONFIG NIVEAU
 * 2. Premier onglet source trouvé (pattern °digit)
 * 3. Fallback '5e'
 *
 * @returns {string} '6e', '5e', '4e' ou '3e'
 */
function detectNiveauAuto() {
  // Stratégie 1 : lire depuis _CONFIG
  try {
    var ss = SpreadsheetApp.getActive();
    var configSheet = ss.getSheetByName('_CONFIG');
    if (configSheet) {
      var data = configSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        var param = String(data[i][0] || '').trim().toUpperCase();
        if (param === 'NIVEAU') {
          var val = String(data[i][1] || '').trim().toLowerCase();
          // "5e", "5°", "5eme" → "5e"
          var numMatch = val.match(/(\d)/);
          if (numMatch) return numMatch[1] + 'e';
        }
      }
    }
  } catch (e) { /* ignore */ }

  // Stratégie 2 : premier onglet source
  try {
    var sheets = SpreadsheetApp.getActive().getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (/.+°\d+$/.test(name)) {
        return detectNiveauFromSheetName(name);
      }
    }
  } catch (e) { /* ignore */ }

  return '5e';
}

/**
 * Liste tous les niveaux disponibles.
 * @returns {string[]} ['6e', '5e', '4e', '3e']
 */
function getAvailableNiveaux() {
  return Object.keys(MATIERES_PAR_NIVEAU);
}

// =============================================================================
// UTILITAIRES INTERNES
// =============================================================================

/**
 * Normalise un identifiant de niveau vers le format standard 'Xe'.
 * "5°", "5e", "5eme", "5ème", "5" → "5e"
 */
function normalizeNiveau_(niveau) {
  if (!niveau) return '5e';
  var s = String(niveau).trim().toLowerCase();
  var m = s.match(/(\d)/);
  if (m) {
    var num = m[1];
    if (num === '6' || num === '5' || num === '4' || num === '3') {
      return num + 'e';
    }
  }
  return '5e';
}
