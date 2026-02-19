/**
 * ===================================================================
 * HARMONY CONSTANTS - Source unique des valeurs partagées
 * ===================================================================
 *
 * Centralise les listes LV2/OPT, critères académiques, PRNG seedable,
 * et ClassState incrémental pour les deux pipelines (NAUTILUS + LEGACY).
 *
 * RÈGLE : Toute modification de langue ou option se fait ICI.
 *
 * Branche : claude/multi-restart-algo-j0ITC
 * Date : 2026-02-10
 * ===================================================================
 */

// Liste des LV2 connues (langues vivantes 2)
const HARMONY_LV2_LIST = ['ITA', 'ESP', 'ALL', 'PT'];

// Liste des options connues
const HARMONY_OPT_LIST = ['CHAV', 'LATIN', 'GREC'];

// Critères académiques utilisés pour le scoring
const HARMONY_CRITERIA = ['COM', 'TRA', 'PART', 'ABS'];

// Scores possibles (1 à 4)
const HARMONY_SCORE_VALUES = [1, 2, 3, 4];

/**
 * Vérifie si une valeur est une LV2 connue
 * @param {string} val - Valeur à tester (ex: 'ITA')
 * @returns {boolean}
 */
function isKnownLV2(val) {
  return HARMONY_LV2_LIST.indexOf(val) >= 0;
}

/**
 * Vérifie si une valeur est une option connue
 * @param {string} val - Valeur à tester (ex: 'CHAV')
 * @returns {boolean}
 */
function isKnownOPT(val) {
  return HARMONY_OPT_LIST.indexOf(val) >= 0;
}

/**
 * Calcule le profil académique moyen d'un élève (COM+TRA+PART+ABS)/4
 * @param {Object} row - Ligne de données
 * @param {Object} idx - Map des index de colonnes {COM: n, TRA: n, ...}
 * @returns {number} Score moyen entre 1 et 4
 */
function calculateStudentProfile(row, idx) {
  var sum = 0;
  var count = 0;
  HARMONY_CRITERIA.forEach(function(crit) {
    var val = Number(row[idx[crit]]);
    if (val >= 1 && val <= 4) {
      sum += val;
      count++;
    }
  });
  return count > 0 ? sum / count : 2.5;
}

/**
 * Détecte si un élève est une "tête de classe" (profil fort)
 * @param {number} com - Score COM
 * @param {number} tra - Score TRA
 * @param {number} part - Score PART (optionnel)
 * @returns {boolean}
 */
function isHeadStudent(com, tra, part) {
  var scoreMoy = part !== undefined ? (com + tra + part) / 3 : (com + tra) / 2;
  return (com >= 4 || tra >= 4) || scoreMoy >= 3.5;
}

/**
 * Détecte si un élève est en difficulté (niveau 1)
 * @param {number} com - Score COM
 * @param {number} tra - Score TRA
 * @returns {boolean}
 */
function isNiv1Student(com, tra) {
  return (com <= 1 || tra <= 1);
}

// ===================================================================
// N3 REPLICANT : PRNG SEEDABLE (Mulberry32)
// ===================================================================

/**
 * Générateur pseudo-aléatoire seedable (Mulberry32).
 * GAS n'offre pas de Math.random seedable, cette implémentation
 * permet de rejouer un run identique en fournissant le même seed.
 *
 * Usage:
 *   var rng = createRNG(42);       // seed fixe pour reproductibilité
 *   var rng = createRNG();          // seed aléatoire (Date.now)
 *   var val = rng.next();           // [0, 1)
 *   var idx = rng.randInt(0, 10);   // entier dans [0, 10)
 *   var item = rng.pick(array);     // élément aléatoire
 *   var shuffled = rng.shuffle(arr); // mélange Fisher-Yates
 *   Logger.log(rng.seed);           // pour rejouer plus tard
 *
 * @param {number} [seed] - Seed entier (défaut: Date.now())
 * @returns {Object} { next, randInt, pick, shuffle, seed }
 */
function createRNG(seed) {
  if (seed === undefined || seed === null) {
    seed = Date.now();
  }
  seed = seed | 0; // Forcer entier 32 bits
  var _state = seed;

  function _mulberry32() {
    _state = (_state + 0x6D2B79F5) | 0;
    var t = Math.imul(_state ^ (_state >>> 15), 1 | _state);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  }

  return {
    seed: seed,
    /** @returns {number} float dans [0, 1) */
    next: function() { return _mulberry32(); },
    /** @returns {number} entier dans [min, max) */
    randInt: function(min, max) {
      return min + Math.floor(_mulberry32() * (max - min));
    },
    /** @returns {*} élément aléatoire du tableau */
    pick: function(arr) {
      if (!arr || arr.length === 0) return undefined;
      return arr[Math.floor(_mulberry32() * arr.length)];
    },
    /** Mélange Fisher-Yates in-place, retourne le tableau */
    shuffle: function(arr) {
      for (var i = arr.length - 1; i > 0; i--) {
        var j = Math.floor(_mulberry32() * (i + 1));
        var tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
      }
      return arr;
    }
  };
}

// ===================================================================
// N1 REPLICANT : ClassState incrémental
// ===================================================================

/**
 * Maintient les compteurs d'une classe pour calcul incrémental O(1).
 * Au lieu de recalculer parité + histogrammes sur tous les élèves à chaque
 * simulation de swap, on met à jour uniquement les deltas.
 *
 * @param {string} className
 * @param {Array} studentIndices - indices dans data[]
 * @param {Array} data - tableau complet des élèves
 * @param {Object} headerIdx - map { COM: n, TRA: n, PART: n, ABS: n, SEXE: n }
 */
function ClassState(className, studentIndices, data, headerIdx, targetSize) {
  this.name = className;
  this.size = studentIndices.length;
  this.targetSize = targetSize || 0; // MULTI-RESTART: effectif cible pour pénalité
  this.nbF = 0;
  this.nbM = 0;
  this.nbTetes = 0;
  this.nbNiv1 = 0;
  this.sumCOM = 0;
  this.sumTRA = 0;
  this.sumPART = 0;
  this.sumABS = 0;

  // Histogrammes [score 1..4] pour chaque critère
  this.histCOM = [0, 0, 0, 0];
  this.histTRA = [0, 0, 0, 0];
  this.histPART = [0, 0, 0, 0];
  this.histABS = [0, 0, 0, 0];

  // Initialiser depuis les données
  for (var i = 0; i < studentIndices.length; i++) {
    this._addStudent(studentIndices[i], data, headerIdx);
  }
}

/** Ajoute un élève aux compteurs (O(1)) */
ClassState.prototype._addStudent = function(idx, data, hIdx) {
  var row = data[idx];
  var sexe = String(row[hIdx.SEXE] || 'M').toUpperCase().trim().charAt(0);
  if (sexe === 'F') this.nbF++; else this.nbM++;

  var com = Number(row[hIdx.COM] || 2.5);
  var tra = Number(row[hIdx.TRA] || 2.5);
  var part = hIdx.PART >= 0 ? Number(row[hIdx.PART] || 2.5) : 2.5;
  var abs = hIdx.ABS >= 0 ? Number(row[hIdx.ABS] || 2.5) : 2.5;

  this.sumCOM += com;
  this.sumTRA += tra;
  this.sumPART += part;
  this.sumABS += abs;

  // Histogrammes (clamp 1-4)
  var ci = Math.max(0, Math.min(3, Math.round(com) - 1));
  var ti = Math.max(0, Math.min(3, Math.round(tra) - 1));
  var pi = Math.max(0, Math.min(3, Math.round(part) - 1));
  var ai = Math.max(0, Math.min(3, Math.round(abs) - 1));
  this.histCOM[ci]++;
  this.histTRA[ti]++;
  this.histPART[pi]++;
  this.histABS[ai]++;

  if (isHeadStudent(com, tra, part)) this.nbTetes++;
  if (isNiv1Student(com, tra)) this.nbNiv1++;
};

/** Retire un élève des compteurs (O(1)) */
ClassState.prototype._removeStudent = function(idx, data, hIdx) {
  var row = data[idx];
  var sexe = String(row[hIdx.SEXE] || 'M').toUpperCase().trim().charAt(0);
  if (sexe === 'F') this.nbF--; else this.nbM--;

  var com = Number(row[hIdx.COM] || 2.5);
  var tra = Number(row[hIdx.TRA] || 2.5);
  var part = hIdx.PART >= 0 ? Number(row[hIdx.PART] || 2.5) : 2.5;
  var abs = hIdx.ABS >= 0 ? Number(row[hIdx.ABS] || 2.5) : 2.5;

  this.sumCOM -= com;
  this.sumTRA -= tra;
  this.sumPART -= part;
  this.sumABS -= abs;

  var ci = Math.max(0, Math.min(3, Math.round(com) - 1));
  var ti = Math.max(0, Math.min(3, Math.round(tra) - 1));
  var pi = Math.max(0, Math.min(3, Math.round(part) - 1));
  var ai = Math.max(0, Math.min(3, Math.round(abs) - 1));
  this.histCOM[ci]--;
  this.histTRA[ti]--;
  this.histPART[pi]--;
  this.histABS[ai]--;

  if (isHeadStudent(com, tra, part)) this.nbTetes--;
  if (isNiv1Student(com, tra)) this.nbNiv1--;

  this.size--;
};

/**
 * Calcule l'erreur de cette classe seule (O(1) car on utilise les compteurs).
 * @param {Object} globalStats - { ratioF, avgCOM, avgTRA, avgPART, totalStudents }
 * @param {Object} targetDistribution - distribution cible par score [1..4]
 * @param {Object} weights - { parity, com, tra, part, abs, profiles }
 * @returns {number} erreur de la classe
 */
ClassState.prototype.computeError = function(globalStats, targetDistribution, weights) {
  if (this.size === 0) return 10000;

  var error = 0;

  // MULTI-RESTART: Pénalité d'effectif (manquait dans V14I, crucial pour l'équilibre)
  if (this.targetSize > 0) {
    var sizeDiff = this.size - this.targetSize;
    error += sizeDiff * sizeDiff * (weights.effectif || 2.0);
  }

  // Parité
  var ratioF = this.nbF / this.size;
  error += Math.abs(ratioF - globalStats.ratioF) * (weights.parity || 1.0);

  // Harmonie COM/TRA/PART/ABS (distribution des scores)
  var hists = { COM: this.histCOM, TRA: this.histTRA, PART: this.histPART, ABS: this.histABS };
  var weightMap = { COM: weights.com || 1.0, TRA: weights.tra || 0.5, PART: weights.part || 0.3, ABS: weights.abs || 0.2 };

  for (var crit in hists) {
    var hist = hists[crit];
    var w = weightMap[crit];
    if (!targetDistribution) {
      // Fallback : écart de la moyenne à la cible
      var sums = { COM: this.sumCOM, TRA: this.sumTRA, PART: this.sumPART, ABS: this.sumABS };
      var avg = sums[crit] / this.size;
      var target = globalStats['avg' + crit] || 2.5;
      error += Math.abs(avg - target) * w;
    } else {
      // Distribution : comparer histogramme réel vs proportionnel
      for (var v = 0; v < 4; v++) {
        var expected = (targetDistribution[crit] && targetDistribution[crit][v + 1] !== undefined)
          ? targetDistribution[crit][v + 1] * this.size
          : this.size / 4;
        error += Math.abs(hist[v] - expected) * w;
      }
    }
  }

  // Profils (têtes de classe, niv1)
  var pw = weights.profiles || 2.0;
  var targetHeadMin = 2;
  var targetHeadMax = Math.max(5, Math.ceil(this.size * 0.2));
  var targetNiv1Max = Math.max(4, Math.ceil(this.size * 0.15));

  if (this.nbTetes < targetHeadMin) {
    error += Math.pow(targetHeadMin - this.nbTetes, 2) * pw;
  }
  if (this.nbTetes > targetHeadMax) {
    error += (this.nbTetes - targetHeadMax) * pw * 0.5;
  }
  if (this.nbNiv1 > targetNiv1Max) {
    error += Math.pow(this.nbNiv1 - targetNiv1Max, 2) * pw;
  }

  return error;
};

/**
 * Simule l'erreur si on retire studentOut et ajoute studentIn (O(1)).
 * Ne modifie PAS l'état interne — lecture seule.
 */
ClassState.prototype.simulateSwap = function(outIdx, inIdx, data, hIdx, globalStats, targetDistribution, weights) {
  // Sauvegarder, modifier, calculer, restaurer
  this._removeStudent(outIdx, data, hIdx);
  this._addStudent(inIdx, data, hIdx);
  this.size++; // _addStudent ne l'incrémente pas, _removeStudent l'a décrémenté
  var err = this.computeError(globalStats, targetDistribution, weights);
  // Restaurer
  this._removeStudent(inIdx, data, hIdx);
  this._addStudent(outIdx, data, hIdx);
  this.size++;
  return err;
};

/**
 * Applique réellement le swap (retire outIdx, ajoute inIdx).
 */
ClassState.prototype.applySwap = function(outIdx, inIdx, data, hIdx) {
  this._removeStudent(outIdx, data, hIdx);
  this._addStudent(inIdx, data, hIdx);
  this.size++; // _addStudent doesn't increment, _removeStudent decremented
};

// ===================================================================
// MULTI-RESTART CONFIGURATION
// ===================================================================

/**
 * Configuration par défaut du multi-restart.
 * Utilisée par les deux pipelines (NAUTILUS + LEGACY).
 * GAS autorise 6 min d'exécution ; un run Phase 4 prend ~10-30s,
 * donc 5 restarts tiennent largement dans le budget.
 */
var MULTI_RESTART_CONFIG = {
  maxRestarts: 5,           // Nombre de seeds à tester
  seedSpacing: 7919,        // Espacement entre seeds (nombre premier)
  crossPhaseLoops: 2,       // Boucles Phase3→Phase4 supplémentaires
  reshuffleWorstRatio: 0.5, // Fraction du worst-class à réinjecter dans le pool
  minImprovementPct: 0.005  // Seuil d'amélioration minimal pour continuer les boucles (0.5%)
};

/**
 * Clone profond d'un tableau 2D (snapshot _BASEOPTI / data).
 * Plus rapide que JSON.parse(JSON.stringify()) pour les tableaux de primitives.
 */
function snapshotData_(data) {
  var copy = new Array(data.length);
  for (var i = 0; i < data.length; i++) {
    copy[i] = data[i].slice();
  }
  return copy;
}

/**
 * Clone profond d'un objet byClass { className: [indices] }.
 */
function snapshotByClass_(byClass) {
  var copy = {};
  for (var cls in byClass) {
    copy[cls] = byClass[cls].slice();
  }
  return copy;
}

/**
 * Calcule le score global d'erreur pour l'ensemble des classes (somme de ClassState.computeError).
 * @param {Object} classStates - Map className → ClassState
 * @param {Object} globalStats
 * @param {Object} targetDistribution
 * @param {Object} weights
 * @returns {number}
 */
function computeTotalError_(classStates, globalStats, targetDistribution, weights) {
  var err = 0;
  for (var cls in classStates) {
    err += classStates[cls].computeError(globalStats, targetDistribution, weights);
  }
  return err;
}
