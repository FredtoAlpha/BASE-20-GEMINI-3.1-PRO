/**
 * ===================================================================
 * SCORING_CONFIG.JS — Configuration dynamique du scoring
 * ===================================================================
 *
 * Gère la configuration scoring par niveau (6e/5e/4e/3e) :
 * - Mode seuils fixes (comportement actuel) ou percentile
 * - Seuils personnalisables par critère (COM/TRA/PART/ABS)
 * - Distribution percentile configurable
 * - Persistance dans l'onglet caché _SCORING_CONFIG (pattern KV)
 *
 * @version 1.0.0
 * ===================================================================
 */

// =============================================================================
// DEFAULTS — Valeurs par défaut (identiques à l'ancien SCORES_CONFIG)
// =============================================================================

var SCORING_DEFAULTS = {
  mode: 'seuils', // 'seuils' ou 'percentile'

  seuils: {
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
    },
    COM: [
      { score: 4, min: 0, max: 0 },
      { score: 3, min: 1, max: 5 },
      { score: 2, min: 6, max: 20 },
      { score: 1, min: 21, max: 999 }
    ],
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
    ]
  },

  // Distribution percentile (si mode='percentile')
  percentile: {
    distribution: { 1: 0.10, 2: 0.25, 3: 0.40, 4: 0.25 }
  },

  // Poids des critères pour le score composite
  poidsCriteres: { COM: 0.25, TRA: 0.40, PART: 0.10, ABS: 0.25 },

  // Patterns de détection colonnes Pronote (indépendants du niveau)
  patterns: {
    ABS: {
      nom: ['NOM'],
      classe: ['CLASSE'],
      dj: ['DJ', 'DEMI.?JOURN', 'DJ.*BULL'],
      justifiee: ['JUSTIFI']
    },
    INC: {
      nom: ['NOM'],
      classe: ['CLASSE'],
      gravite: ['GRAVIT', 'GRAV']
    },
    PUN: {
      nom: ['NOM'],
      classe: ['CLASSE'],
      nb: ['^NB', 'NOMBRE', 'QT', 'QUANT']
    }
  }
};

// =============================================================================
// KV STORE — Persistance dans _SCORING_CONFIG
// =============================================================================

var SCORING_CONFIG_SHEET_NAME = '_SCORING_CONFIG';

/**
 * S'assure que la feuille _SCORING_CONFIG existe (cachée).
 */
function ensureScoringConfigSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SCORING_CONFIG_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SCORING_CONFIG_SHEET_NAME);
    sh.hideSheet();
    var headers = ['KEY', 'VALUE', 'SCOPE', 'UPDATED_AT'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#6366f1')
      .setFontColor('#FFFFFF');
  }

  return sh;
}

/**
 * Lit une clé depuis _SCORING_CONFIG.
 */
function scoringKvGet_(key, scope, defaultValue) {
  scope = scope || 'GLOBAL';
  defaultValue = defaultValue !== undefined ? defaultValue : null;

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SCORING_CONFIG_SHEET_NAME);

  if (!sh || sh.getLastRow() <= 1) return defaultValue;

  var last = sh.getLastRow();
  var data = sh.getRange(2, 1, last - 1, 3).getValues();

  for (var i = 0; i < data.length; i++) {
    var cellKey = String(data[i][0]).trim();
    var cellScope = String(data[i][2]).trim();
    if (cellKey === key && cellScope === scope) {
      var val = data[i][1];
      // Trim les chaînes pour éviter les espaces parasites de Google Sheets
      if (typeof val === 'string') val = val.trim();
      return val;
    }
  }

  return defaultValue;
}

/**
 * Écrit une clé dans _SCORING_CONFIG.
 */
function scoringKvSet_(key, value, scope) {
  scope = scope || 'GLOBAL';

  var sh = ensureScoringConfigSheet_();
  var last = sh.getLastRow();
  var now = new Date();

  var valueStr = (typeof value === 'object') ? JSON.stringify(value) : String(value);

  var row = -1;
  if (last > 1) {
    var data = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var cellKey = String(data[i][0]).trim();
      var cellScope = String(data[i][2]).trim();
      if (cellKey === key && cellScope === scope) {
        row = i + 2;
        break;
      }
    }
  }

  if (row === -1) row = last + 1;

  sh.getRange(row, 1, 1, 4).setValues([[key, valueStr, scope, now]]);
}

// =============================================================================
// API PUBLIQUE
// =============================================================================

/**
 * Retourne la configuration scoring pour un niveau donné.
 * Merge: SCORING_DEFAULTS + overrides depuis _SCORING_CONFIG.
 *
 * @param {string} [niveau] - '6e', '5e', '4e', '3e' (optionnel)
 * @returns {Object} Config scoring complète
 */
function getScoringConfig(niveau) {
  // Commencer avec les defaults
  var config = JSON.parse(JSON.stringify(SCORING_DEFAULTS));

  // Lire le mode depuis KV (valider que c'est une valeur connue)
  var modeKv = scoringKvGet_('scoring.mode', 'GLOBAL', null);
  if (modeKv) {
    var modeStr = String(modeKv).trim().toLowerCase();
    if (modeStr === 'seuils' || modeStr === 'percentile') {
      config.mode = modeStr;
    } else {
      Logger.log('Scoring_Config: mode KV invalide "' + modeKv + '", fallback sur defaults');
    }
  }

  // Lire les seuils custom depuis KV
  var seuilsJson = scoringKvGet_('scoring.seuils', 'GLOBAL', null);
  if (seuilsJson) {
    try {
      var customSeuils = JSON.parse(seuilsJson);
      // Merge par critère (ne remplace que ce qui est fourni)
      for (var crit in customSeuils) {
        if (customSeuils.hasOwnProperty(crit)) {
          config.seuils[crit] = customSeuils[crit];
        }
      }
    } catch (e) {
      Logger.log('Scoring_Config: erreur parsing seuils custom: ' + e.message);
    }
  }

  // Lire la distribution percentile depuis KV
  var distJson = scoringKvGet_('scoring.percentile.distribution', 'GLOBAL', null);
  if (distJson) {
    try {
      config.percentile.distribution = JSON.parse(distJson);
    } catch (e) {
      Logger.log('Scoring_Config: erreur parsing distribution: ' + e.message);
    }
  }

  // Override par niveau si disponible
  if (niveau) {
    var niveauSeuilsJson = scoringKvGet_('scoring.seuils', niveau, null);
    if (niveauSeuilsJson) {
      try {
        var niveauSeuils = JSON.parse(niveauSeuilsJson);
        for (var crit in niveauSeuils) {
          if (niveauSeuils.hasOwnProperty(crit)) {
            config.seuils[crit] = niveauSeuils[crit];
          }
        }
      } catch (e) {
        Logger.log('Scoring_Config: erreur parsing seuils niveau ' + niveau + ': ' + e.message);
      }
    }
  }

  // Lire les poids critères
  var poidsJson = scoringKvGet_('scoring.poidsCriteres', 'GLOBAL', null);
  if (poidsJson) {
    try {
      config.poidsCriteres = JSON.parse(poidsJson);
    } catch (e) {
      Logger.log('Scoring_Config: erreur parsing poids: ' + e.message);
    }
  }

  return config;
}

/**
 * Sauvegarde la configuration scoring.
 *
 * @param {Object} config - { mode, seuils, percentile, poidsCriteres }
 * @param {string} [niveau] - Scope niveau (optionnel, 'GLOBAL' par défaut)
 */
function saveScoringConfig(config, niveau) {
  var scope = niveau || 'GLOBAL';

  if (config.mode) {
    scoringKvSet_('scoring.mode', config.mode, 'GLOBAL');
  }

  if (config.seuils) {
    scoringKvSet_('scoring.seuils', config.seuils, scope);
  }

  if (config.percentile && config.percentile.distribution) {
    var d = config.percentile.distribution;
    var sum = (d[1] || 0) + (d[2] || 0) + (d[3] || 0) + (d[4] || 0);
    if (sum > 0 && Math.abs(sum - 1.0) > 0.05) {
      Logger.log('⚠️ saveScoringConfig: normalisation distribution percentile (somme=' + sum.toFixed(3) + ')');
      d = { 1: (d[1] || 0) / sum, 2: (d[2] || 0) / sum, 3: (d[3] || 0) / sum, 4: (d[4] || 0) / sum };
    }
    scoringKvSet_('scoring.percentile.distribution', d, 'GLOBAL');
  }

  if (config.poidsCriteres) {
    scoringKvSet_('scoring.poidsCriteres', config.poidsCriteres, 'GLOBAL');
  }

  SpreadsheetApp.flush();
}

/**
 * Retourne le mode scoring actif.
 * @returns {string} 'seuils' ou 'percentile'
 */
function getScoringMode() {
  return scoringKvGet_('scoring.mode', 'GLOBAL', 'seuils');
}

/**
 * Retourne la distribution percentile configurée.
 * @returns {Object} { 1: 0.10, 2: 0.25, 3: 0.40, 4: 0.25 }
 */
function getPercentileDistribution() {
  var distJson = scoringKvGet_('scoring.percentile.distribution', 'GLOBAL', null);
  if (distJson) {
    try {
      return JSON.parse(distJson);
    } catch (e) { /* fallback */ }
  }
  return SCORING_DEFAULTS.percentile.distribution;
}

/**
 * Retourne les patterns de détection de colonnes Pronote.
 * @returns {Object} { ABS: {...}, INC: {...}, PUN: {...} }
 */
function getScoringPatterns() {
  return SCORING_DEFAULTS.patterns;
}
