/**
 * ===================================================================
 * SCORING_PERCENTILE.JS — Moteur de scoring par percentile
 * ===================================================================
 *
 * Calcule les scores 1-4 en se basant sur le rang de chaque élève
 * dans la cohorte, plutôt que sur des seuils fixes.
 *
 * Distribution configurable : { 1: 0.10, 2: 0.25, 3: 0.40, 4: 0.25 }
 * → score 1 = bottom 10%, score 2 = next 25%, score 3 = next 40%, score 4 = top 25%
 *
 * @version 1.0.0
 * ===================================================================
 */

/**
 * Calcule les scores 1-4 par percentile pour un tableau de valeurs.
 *
 * @param {Array} entries - Tableau de { index: number, valeur: number|null }
 *   - index = identifiant unique de l'élève (position dans le tableau original)
 *   - valeur = note brute (null = pas de note → exclu du calcul)
 * @param {Object} distribution - { 1: fraction, 2: fraction, 3: fraction, 4: fraction }
 *   Les fractions doivent sommer à 1.0 (±0.01)
 * @returns {Array} Tableau de { index, valeur, score } (mêmes entries + score ajouté)
 */
function computePercentileScores(entries, distribution) {
  if (!entries || entries.length === 0) return [];
  if (!distribution) distribution = { 1: 0.10, 2: 0.25, 3: 0.40, 4: 0.25 };

  // Valider que la distribution somme à ~1.0, normaliser sinon
  var distSum = (distribution[1] || 0) + (distribution[2] || 0) + (distribution[3] || 0) + (distribution[4] || 0);
  if (distSum > 0 && Math.abs(distSum - 1.0) > 0.05) {
    Logger.log('⚠️ Distribution percentile invalide (somme=' + distSum.toFixed(3) + '), normalisation...');
    distribution = {
      1: (distribution[1] || 0) / distSum,
      2: (distribution[2] || 0) / distSum,
      3: (distribution[3] || 0) / distSum,
      4: (distribution[4] || 0) / distSum
    };
  }

  // Séparer les entrées avec/sans valeur
  var withValue = [];
  var withoutValue = [];

  for (var i = 0; i < entries.length; i++) {
    if (entries[i].valeur !== null && entries[i].valeur !== undefined && !isNaN(entries[i].valeur)) {
      withValue.push({ index: entries[i].index, valeur: entries[i].valeur });
    } else {
      withoutValue.push({ index: entries[i].index, valeur: null, score: null });
    }
  }

  if (withValue.length === 0) {
    return withoutValue;
  }

  // Trier par valeur croissante
  withValue.sort(function(a, b) { return a.valeur - b.valeur; });

  var N = withValue.length;

  // Calculer les seuils de coupure cumulés
  var cumul1 = distribution[1] || 0.10;
  var cumul2 = cumul1 + (distribution[2] || 0.25);
  var cumul3 = cumul2 + (distribution[3] || 0.40);
  // Le reste va au score 4

  // Indices de coupure
  var cut1 = Math.floor(N * cumul1);
  var cut2 = Math.floor(N * cumul2);
  var cut3 = Math.floor(N * cumul3);

  // Assigner les scores
  for (var i = 0; i < withValue.length; i++) {
    if (i < cut1) {
      withValue[i].score = 1;
    } else if (i < cut2) {
      withValue[i].score = 2;
    } else if (i < cut3) {
      withValue[i].score = 3;
    } else {
      withValue[i].score = 4;
    }
  }

  // Fusionner et retourner
  return withValue.concat(withoutValue);
}

/**
 * Calcule les seuils de coupure percentile pour une série de valeurs.
 * Utile pour l'affichage dans l'UI (montrer les notes de coupure).
 *
 * @param {number[]} valeurs - Notes brutes (exclure les null avant d'appeler)
 * @param {Object} distribution - { 1: fraction, 2: fraction, 3: fraction, 4: fraction }
 * @returns {Object} {
 *   cutoffs: [c1, c2, c3], // notes de coupure entre score 1-2, 2-3, 3-4
 *   counts: [n1, n2, n3, n4], // nombre d'élèves par score
 *   total: number
 * }
 */
function computePercentileSeuils(valeurs, distribution) {
  if (!valeurs || valeurs.length === 0) {
    return { cutoffs: [0, 0, 0], counts: [0, 0, 0, 0], total: 0 };
  }
  if (!distribution) distribution = { 1: 0.10, 2: 0.25, 3: 0.40, 4: 0.25 };

  // Filtrer les valeurs valides et trier
  var sorted = valeurs.filter(function(v) {
    return v !== null && v !== undefined && !isNaN(v);
  }).sort(function(a, b) { return a - b; });

  var N = sorted.length;
  if (N === 0) {
    return { cutoffs: [0, 0, 0], counts: [0, 0, 0, 0], total: 0 };
  }

  var cumul1 = distribution[1] || 0.10;
  var cumul2 = cumul1 + (distribution[2] || 0.25);
  var cumul3 = cumul2 + (distribution[3] || 0.40);

  var idx1 = Math.max(0, Math.floor(N * cumul1) - 1);
  var idx2 = Math.max(0, Math.floor(N * cumul2) - 1);
  var idx3 = Math.max(0, Math.floor(N * cumul3) - 1);

  var cut1 = Math.floor(N * cumul1);
  var cut2 = Math.floor(N * cumul2);
  var cut3 = Math.floor(N * cumul3);

  return {
    cutoffs: [
      Math.round(sorted[idx1] * 100) / 100,
      Math.round(sorted[idx2] * 100) / 100,
      Math.round(sorted[idx3] * 100) / 100
    ],
    counts: [
      cut1,
      cut2 - cut1,
      cut3 - cut2,
      N - cut3
    ],
    total: N
  };
}

/**
 * Applique le scoring percentile sur un critère complet.
 * Wrapper de haut niveau pour le pipeline Backend_Scores.
 *
 * @param {Array} resultats - Tableau de résultats bruts
 *   Chaque élément doit avoir: { nom, classe, valeurBrute }
 * @param {string} scoreField - Nom du champ score à remplir (ex: 'scoreTRA')
 * @param {Object} [distribution] - Distribution personnalisée
 * @returns {Array} Mêmes résultats avec le champ score rempli
 */
function applyPercentileToResults(resultats, scoreField, distribution) {
  if (!resultats || resultats.length === 0) return resultats;

  // Construire les entries pour le moteur
  var entries = [];
  for (var i = 0; i < resultats.length; i++) {
    entries.push({
      index: i,
      valeur: resultats[i].valeurBrute !== undefined ? resultats[i].valeurBrute : null
    });
  }

  // Calculer les scores
  var scored = computePercentileScores(entries, distribution);

  // Réinjecter dans les résultats
  for (var j = 0; j < scored.length; j++) {
    var idx = scored[j].index;
    resultats[idx][scoreField] = scored[j].score;
  }

  return resultats;
}
