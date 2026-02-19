/**
 * ===================================================================
 * 🔄 PRIME LEGACY - SYSTÈME DE MOBILITÉ
 * ===================================================================
 *
 * Calcul des flags de mobilité (FIXE/PERMUT/LIBRE) pour le pipeline LEGACY
 * Basé sur Mobility_System.gs adapté pour le contexte LEGACY
 *
 * LOGIQUE :
 * - FIXE : Une seule classe autorisée (LV2/OPT unique)
 * - PERMUT : Deux classes autorisées (peut permuter entre elles)
 * - LIBRE : Plus de deux classes autorisées
 * - GROUPE_FIXE : Groupe ASSO avec une seule classe commune
 * - GROUPE_PERMUT : Groupe ASSO avec plusieurs classes communes
 *
 * Date : 2025-11-14
 * Branche : claude/legacy-pipeline-renewal-01FK5TFnxx6JjwZ9bMkF5hqw
 *
 * ===================================================================
 */

/**
 * Calcule les flags de mobilité (FIXE/PERMUT/LIBRE) après Phase 1
 * Analyse les LV2/OPT et détermine quelles classes sont autorisées pour chaque élève
 *
 * @param {Object} ctx - Contexte LEGACY
 */
function computeMobilityFlags_LEGACY(ctx) {
  logLine('INFO', '🔄 Calcul mobilité (FIXE/PERMUT/LIBRE)...');

  const ss = ctx.ss || SpreadsheetApp.getActive();

  // ========== ÉTAPE 1 : CONSTRUIRE TABLE DES OFFRES ==========
  // Quelles classes proposent quelles LV2/OPT ?
  const classOffers = buildClassOffersFromQuotas_LEGACY(ctx);

  logLine('INFO', '  📊 Offres par classe :');
  for (const classe in classOffers) {
    const lv2List = Array.from(classOffers[classe].LV2).join(', ') || 'aucune';
    const optList = Array.from(classOffers[classe].OPT).join(', ') || 'aucune';
    logLine('INFO', '    • ' + classe + ' : LV2={' + lv2List + '}, OPT={' + optList + '}');
  }

  // ========== ÉTAPE 2 : PARCOURIR CHAQUE ONGLET TEST ==========
  let totalProcessed = 0;
  let stats = { FIXE: 0, PERMUT: 0, LIBRE: 0, GROUPE_FIXE: 0, GROUPE_PERMUT: 0 };

  (ctx.cacheSheets || []).forEach(function(testName) {
    const testSheet = ss.getSheetByName(testName);
    if (!testSheet || testSheet.getLastRow() <= 1) {
      logLine('WARN', '  ⚠️ ' + testName + ' vide, skip');
      return;
    }

    const data = testSheet.getDataRange().getValues();
    const headers = data[0];

    // Trouver index des colonnes
    const idxLV2 = headers.indexOf('LV2');
    const idxOPT = headers.indexOf('OPT');
    const idxMobilite = headers.indexOf('MOBILITE');
    const idxFixe = headers.indexOf('FIXE');
    const idxAsso = headers.indexOf('ASSO');
    const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');

    if (idxMobilite === -1 || idxFixe === -1) {
      logLine('WARN', '  ⚠️ ' + testName + ' : colonnes MOBILITE ou FIXE manquantes, skip');
      return;
    }

    // ========== ÉTAPE 3 : CALCULER MOBILITÉ POUR CHAQUE ÉLÈVE ==========
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const lv2 = String(row[idxLV2] || '').trim().toUpperCase();
      const opt = String(row[idxOPT] || '').trim().toUpperCase();
      const asso = String(row[idxAsso] || '').trim().toUpperCase();
      const assigned = String(row[idxAssigned] || '').trim();

      // Calculer les classes autorisées pour cet élève
      const allowedClasses = computeAllowedClasses_LEGACY(lv2, opt, classOffers);

      // Déterminer le statut de mobilité
      let mobilite = 'LIBRE';
      let fixe = '';

      if (allowedClasses.length === 0) {
        mobilite = 'CONFLIT';
        fixe = '❌';
      } else if (allowedClasses.length === 1) {
        mobilite = 'FIXE';
        fixe = allowedClasses[0];
        stats.FIXE++;
      } else if (allowedClasses.length === 2) {
        mobilite = 'PERMUT';
        fixe = allowedClasses.join(' ↔ ');
        stats.PERMUT++;
      } else {
        mobilite = 'LIBRE';
        fixe = '';
        stats.LIBRE++;
      }

      // Si l'élève a un code ASSO, préfixer avec GROUPE_
      if (asso) {
        if (mobilite === 'FIXE') {
          mobilite = 'GROUPE_FIXE';
          stats.GROUPE_FIXE++;
          stats.FIXE--;
        } else if (mobilite === 'PERMUT') {
          mobilite = 'GROUPE_PERMUT';
          stats.GROUPE_PERMUT++;
          stats.PERMUT--;
        }
      }

      // Écrire dans les colonnes
      data[i][idxMobilite] = mobilite;
      data[i][idxFixe] = fixe;

      totalProcessed++;
    }

    // ========== ÉTAPE 4 : RÉÉCRIRE LES DONNÉES ==========
    testSheet.getRange(1, 1, data.length, headers.length).setValues(data);
    logLine('INFO', '  ✅ ' + testName + ' : ' + (data.length - 1) + ' élèves traités');
  });

  SpreadsheetApp.flush();

  // ========== ÉTAPE 5 : AFFICHER STATISTIQUES ==========
  logLine('INFO', '✅ Mobilité calculée pour ' + totalProcessed + ' élèves');
  logLine('INFO', '  📊 Statistiques :');
  logLine('INFO', '    • FIXE : ' + stats.FIXE + ' élèves');
  logLine('INFO', '    • PERMUT : ' + stats.PERMUT + ' élèves');
  logLine('INFO', '    • LIBRE : ' + stats.LIBRE + ' élèves');
  logLine('INFO', '    • GROUPE_FIXE : ' + stats.GROUPE_FIXE + ' groupes');
  logLine('INFO', '    • GROUPE_PERMUT : ' + stats.GROUPE_PERMUT + ' groupes');
}

// ===================================================================
// FONCTIONS AUXILIAIRES
// ===================================================================

/**
 * Construit la table des offres LV2/OPT par classe depuis ctx.quotas
 * @param {Object} ctx - Contexte LEGACY
 * @returns {Object} { "6°1": { LV2: Set(['ITA']), OPT: Set(['CHAV']) }, ... }
 */
function buildClassOffersFromQuotas_LEGACY(ctx) {
  const offers = {};

  // Initialiser toutes les classes
  (ctx.niveaux || []).forEach(function(classe) {
    offers[classe] = {
      LV2: new Set(),
      OPT: new Set()
    };
  });

  // Parcourir les quotas
  for (const classe in (ctx.quotas || {})) {
    if (!offers[classe]) {
      offers[classe] = { LV2: new Set(), OPT: new Set() };
    }

    const quotas = ctx.quotas[classe];

    for (const optName in quotas) {
      if (quotas[optName] > 0) {
        const optNameUpper = optName.toUpperCase();

        // Heuristique : LV2 connues (ITA, ESP, ALL, PT, CHI, LAT, etc.)
        if (isLV2Option_LEGACY(optNameUpper)) {
          offers[classe].LV2.add(optNameUpper);
        } else {
          // Sinon, c'est une OPT
          offers[classe].OPT.add(optNameUpper);
        }
      }
    }
  }

  return offers;
}

/**
 * Détermine si une option est une LV2
 * @param {string} optName - Nom de l'option (en majuscules)
 * @returns {boolean}
 */
function isLV2Option_LEGACY(optName) {
  const lv2Patterns = [
    'ITA', 'ITALIEN', 'ITAL',
    'ESP', 'ESPAGNOL',
    'ALL', 'ALLEMAND', 'ALLEM',
    'PT', 'PORTUGAIS', 'PORT',
    'CHI', 'CHINOIS',
    'LAT', 'LATIN',
    'GRE', 'GREC',
    'ARA', 'ARABE',
    'RUS', 'RUSSE',
    'JAP', 'JAPONAIS'
  ];

  return lv2Patterns.some(function(pattern) {
    return optName.indexOf(pattern) >= 0;
  });
}

/**
 * Calcule les classes autorisées pour un élève donné ses LV2/OPT
 * @param {string} lv2 - LV2 de l'élève
 * @param {string} opt - OPT de l'élève
 * @param {Object} classOffers - Table des offres par classe
 * @returns {Array<string>} Liste des classes autorisées
 */
function computeAllowedClasses_LEGACY(lv2, opt, classOffers) {
  const allClasses = Object.keys(classOffers);
  let allowed = allClasses.slice();

  // Filtrer par LV2 (si non vide et différent de ESP/ANG qui sont par défaut)
  if (lv2 && lv2 !== 'ESP' && lv2 !== 'ANG') {
    allowed = allowed.filter(function(classe) {
      return classOffers[classe].LV2.has(lv2);
    });
  }

  // Filtrer par OPT (si non vide)
  if (opt) {
    allowed = allowed.filter(function(classe) {
      return classOffers[classe].OPT.has(opt);
    });
  }

  return allowed;
}

/**
 * Affiche un rapport détaillé de mobilité dans les logs
 * @param {Object} ctx - Contexte LEGACY
 */
function reportMobilityStatus_LEGACY(ctx) {
  logLine('INFO', '');
  logLine('INFO', '📊 RAPPORT DE MOBILITÉ LEGACY');
  logLine('INFO', '─'.repeat(80));

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const stats = {
    total: 0,
    FIXE: 0,
    PERMUT: 0,
    LIBRE: 0,
    GROUPE_FIXE: 0,
    GROUPE_PERMUT: 0,
    CONFLIT: 0
  };

  (ctx.cacheSheets || []).forEach(function(testName) {
    const testSheet = ss.getSheetByName(testName);
    if (!testSheet || testSheet.getLastRow() <= 1) return;

    const data = testSheet.getDataRange().getValues();
    const headers = data[0];
    const idxMobilite = headers.indexOf('MOBILITE');

    if (idxMobilite === -1) return;

    for (let i = 1; i < data.length; i++) {
      const mobilite = String(data[i][idxMobilite] || '').trim().toUpperCase();
      stats.total++;

      if (mobilite.indexOf('FIXE') >= 0) {
        stats.FIXE++;
      } else if (mobilite.indexOf('PERMUT') >= 0) {
        stats.PERMUT++;
      } else if (mobilite.indexOf('LIBRE') >= 0) {
        stats.LIBRE++;
      } else if (mobilite.indexOf('CONFLIT') >= 0) {
        stats.CONFLIT++;
      }

      if (mobilite.indexOf('GROUPE_FIXE') >= 0) stats.GROUPE_FIXE++;
      if (mobilite.indexOf('GROUPE_PERMUT') >= 0) stats.GROUPE_PERMUT++;
    }
  });

  logLine('INFO', '  • Total élèves : ' + stats.total);
  logLine('INFO', '  • FIXE : ' + stats.FIXE + ' (' + ((stats.FIXE / stats.total) * 100).toFixed(1) + '%)');
  logLine('INFO', '  • PERMUT : ' + stats.PERMUT + ' (' + ((stats.PERMUT / stats.total) * 100).toFixed(1) + '%)');
  logLine('INFO', '  • LIBRE : ' + stats.LIBRE + ' (' + ((stats.LIBRE / stats.total) * 100).toFixed(1) + '%)');
  logLine('INFO', '  • GROUPE_FIXE : ' + stats.GROUPE_FIXE);
  logLine('INFO', '  • GROUPE_PERMUT : ' + stats.GROUPE_PERMUT);

  if (stats.CONFLIT > 0) {
    logLine('WARN', '  ⚠️ CONFLITS : ' + stats.CONFLIT + ' élèves sans classe autorisée !');
  }

  logLine('INFO', '─'.repeat(80));
  logLine('INFO', '');
}
