/**
 * ===================================================================
 * APP.CONTEXT.JS - CONSTRUCTION DU CONTEXTE
 * ===================================================================
 *
 * Module contenant les fonctions de construction du contexte d'exécution.
 * Responsabilités: lecture paramètres UI, construction mapping, offres LV2/OPT.
 *
 * ARCHITECTURE PHASE 5 - Refactoring progressif
 * Extraction depuis Orchestration_V14I.js
 *
 * Date: 26 novembre 2025
 * Version: 1.1.0 — SAFE: suppression des stubs incomplets en collision avec Orchestration_V14I.js
 * ===================================================================
 */

// ===================================================================
// CONSTRUCTION DU CONTEXTE DEPUIS L'INTERFACE
// ===================================================================
// NOTE: makeCtxFromSourceSheets_() est définie dans Orchestration_V14I.js (version complète)
// La version ici était un stub incomplet (retour partiel à 5 champs seulement).

/**
 * Construit le contexte depuis l'interface utilisateur
 *
 * @param {Object} options - Options { sourceFamily, targetFamily, maxSwaps, ... }
 * @returns {Object} Contexte complet
 *
 * @example
 * const ctx = makeCtxFromUI_({ sourceFamily: 'TEST', targetFamily: 'CACHE' });
 */
function makeCtxFromUI_(options) {
  const ss = getActiveSS_();

  // Lire le mode source depuis options ou UI (TEST/FIN/CACHE/...)
  const modeSrc = (options && options.sourceFamily) ? String(options.sourceFamily).trim() : (readModeFromUI_() || 'TEST');

  // Le target est CACHE pour l'optimisation
  const writeTarget = (options && options.targetFamily) ? String(options.targetFamily).trim() : 'CACHE';

  // Lire les niveaux à traiter (dynamiquement depuis _STRUCTURE ou _CONFIG)
  const niveaux = (typeof genererNiveauxDynamiques === 'function')
    ? genererNiveauxDynamiques()
    : (readNiveauxFromUI_() || ['6°1', '6°2', '6°3', '6°4', '6°5']);

  // ✅ Construire les noms de feuilles avec les helpers (suffixe uniquement)
  const srcSheets = makeSheetsList_(niveaux, modeSrc);     // ['6°1TEST', '6°2TEST', ...]
  const cacheSheets = makeSheetsList_(niveaux, writeTarget); // ['6°1CACHE', '6°2CACHE', ...]

  logLine('INFO', '📋 Contexte: Source=' + modeSrc + ', Target=' + writeTarget);
  logLine('INFO', '📋 Onglets source: ' + srcSheets.join(', '));
  logLine('INFO', '📋 Onglets cible: ' + cacheSheets.join(', '));

  // Lire les quotas par classe depuis l'interface
  const quotas = readQuotasFromUI_();

  // Lire les cibles d'effectifs par classe
  const targets = readTargetsFromUI_();

  // Lire la tolérance de parité
  const tolParite = readParityToleranceFromUI_() || 2;

  // Lire le nombre max de swaps depuis options ou UI
  const maxSwaps = (options && options.maxSwaps) ? parseInt(options.maxSwaps) : (readMaxSwapsFromUI_() || 500);

  // Lire les autorisations de classes pour options/LV2
  const autorisations = readClassAuthorizationsFromUI_();

  return {
    ss,
    modeSrc,
    writeTarget,
    niveaux,
    srcSheets,
    cacheSheets,
    quotas,
    targets,
    tolParite,
    maxSwaps,
    autorisations
  };
}

// ===================================================================
// LECTURE DES PARAMÈTRES DEPUIS L'INTERFACE
// ===================================================================

/**
 * Lit le mode source depuis l'interface (TEST/CACHE/FIN)
 *
 * @returns {string} Mode source
 */
function readModeFromUI_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uiSheet = ss.getSheetByName('_INTERFACE_V2') || ss.getSheetByName('UI_Config');
  if (!uiSheet) return 'TEST';

  try {
    const value = uiSheet.getRange('B2').getValue();
    return String(value).trim() || 'TEST';
  } catch (e) {
    return 'TEST';
  }
}

/**
 * @deprecated Utiliser genererNiveauxDynamiques() à la place.
 * Conservé pour compat : délègue maintenant à la version dynamique
 * si disponible, au lieu de renvoyer des niveaux hardcodés.
 */
function readNiveauxFromUI_() {
  if (typeof genererNiveauxDynamiques === 'function') {
    try {
      const dyn = genererNiveauxDynamiques();
      if (Array.isArray(dyn) && dyn.length > 0) return dyn;
    } catch (e) {
      if (typeof logLine === 'function') {
        logLine('WARN', `readNiveauxFromUI_ fallback sur hardcoded: ${e && e.message}`);
      }
    }
  }
  return ['6°1', '6°2', '6°3', '6°4', '6°5'];
}

/**
 * Lit les quotas par classe depuis l'interface ou _STRUCTURE
 *
 * @returns {Object} Quotas {classe: {option: quota}}
 */
function readQuotasFromUI_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const structSheet = ss.getSheetByName('_STRUCTURE');
  if (structSheet) {
    return readQuotasFromStructure_(structSheet);
  }

  return {
    "6°1": { ITA: 6 },
    "6°2": {},
    "6°3": { CHAV: 10 },
    "6°4": {},
    "6°5": {}
  };
}

/**
 * Lit les effectifs cibles depuis l'interface ou _STRUCTURE
 *
 * @returns {Object} Effectifs {classe: effectif}
 */
function readTargetsFromUI_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const structSheet = ss.getSheetByName('_STRUCTURE');
  if (structSheet) {
    return readTargetsFromStructure_(structSheet);
  }

  return {
    "6°1": 25,
    "6°2": 25,
    "6°3": 25,
    "6°4": 25,
    "6°5": 25
  };
}

/**
 * @deprecated Valeur codée en dur, lire depuis _OPTI_CONFIG
 */
function readParityToleranceFromUI_() {
  return 2;
}

/**
 * @deprecated Valeur codée en dur
 */
function readMaxSwapsFromUI_() {
  return 500;
}

// NOTE: readClassAuthorizationsFromUI_() est définie dans Orchestration_V14I.js
// La version ici retournait {} (stub vide). La version d'Orchestration retourne les données réelles.

// ===================================================================
// LECTURE DU MAPPING DEPUIS _STRUCTURE
// ===================================================================

/**
 * Lit le mapping CLASSE_ORIGINE → CLASSE_DEST depuis _STRUCTURE
 *
 * @returns {Object} Mapping {origine: destination}
 */
function readSourceToDestMapping_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structSheet = ss.getSheetByName('_STRUCTURE');
  const mapping = {};

  if (!structSheet) {
    return mapping;
  }

  try {
    const data = structSheet.getDataRange().getValues();

    let headerRow = -1;
    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] || '').trim().toUpperCase();
        if (cell === 'CLASSE_DEST' || cell === 'CLASSE_ORIGINE') {
          headerRow = i;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1) {
      logLine('WARN', '⚠️ En-têtes non trouvés dans _STRUCTURE');
      return mapping;
    }

    const headers = data[headerRow];
    const origineCol = headers.indexOf('CLASSE_ORIGINE');
    const destCol = headers.indexOf('CLASSE_DEST');

    if (origineCol === -1 || destCol === -1) {
      logLine('WARN', '⚠️ Colonnes CLASSE_ORIGINE ou CLASSE_DEST introuvables');
      return mapping;
    }

    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      const origine = String(row[origineCol] || '').trim();
      const dest = String(row[destCol] || '').trim();

      if (origine && dest) {
        mapping[origine] = dest;
        logLine('INFO', '  📌 Mapping: ' + origine + ' → ' + dest);
      }
    }

  } catch (e) {
    logLine('WARN', 'Erreur lecture mapping depuis _STRUCTURE : ' + e.message);
  }

  return mapping;
}

// ===================================================================
// CONSTRUCTION DES OFFRES LV2/OPT
// ===================================================================

/**
 * Construit la table des classes offrant LV2/OPT depuis ctx.quotas
 *
 * @param {Object} ctx - Contexte (avec quotas, cacheSheets)
 * @returns {Object} Offres {classe: {LV2: Set, OPT: Set}}
 */
function buildClassOffers_(ctx) {
  const offers = {};

  (ctx.cacheSheets || []).forEach(function(cl) {
    const base = cl.replace(/CACHE$/, '');
    offers[base] = { LV2: new Set(), OPT: new Set() };
  });

  Object.keys(ctx.quotas || {}).forEach(function(classe) {
    const base = classe;
    if (!offers[base]) offers[base] = { LV2: new Set(), OPT: new Set() };
    const q = ctx.quotas[classe] || {};

    Object.keys(q).forEach(function(label) {
      const L = _u_(label);
      if (/(ITA|ALL|ESP|PT|CHI|ANG|GER|LAT2?|ALLEMAND|ESPAGNOL|ITALIEN|CHINOIS|PORTUGAIS)/.test(L)) {
        offers[base].LV2.add(L);
      } else {
        offers[base].OPT.add(L);
      }
    });
  });

  return offers;
}

/**
 * Construit l'offre avec quotas détaillés depuis ctx.quotas
 *
 * @param {Object} ctx - Contexte (avec quotas, cacheSheets)
 * @returns {Object} Offre {cls: {LV2: [], OPT: [], quotas: {option: quota}}}
 */
function buildOfferWithQuotas_(ctx) {
  const res = {};

  (ctx.cacheSheets || []).forEach(function(name) {
    const cls = name.replace(/CACHE$/, '').trim();
    res[cls] = { LV2: [], OPT: [], quotas: {} };
  });

  Object.keys(ctx.quotas || {}).forEach(function(cls) {
    res[cls] = res[cls] || { LV2: [], OPT: [], quotas: {} };
    Object.keys(ctx.quotas[cls]).forEach(function(k) {
      const K = k.toUpperCase();
      const q = Number(ctx.quotas[cls][k]) || 0;
      res[cls].quotas[K] = q;

      if (K === 'CHAV' || K === 'LAT' || K === 'GRE' || K === 'OPT' || K === 'ITA_OPT') {
        res[cls].OPT.push(K === 'ITA_OPT' ? 'ITA' : K);
      } else {
        res[cls].LV2.push(K);
      }
    });
  });

  return res;
}

/**
 * Retourne les classes autorisées pour un élève selon ses LV2/OPT
 *
 * @param {Object} eleve - L'élève
 * @param {Object} classOffers - Offres {classe: {LV2: Set, OPT: Set}}
 * @returns {Array<string>} Liste des classes autorisées
 */
function computeAllow_(eleve, classOffers) {
  const lv2 = _u_(eleve.LV2 || eleve.lv2);
  const opt = _u_(eleve.OPT || eleve.opt);
  const allClasses = Object.keys(classOffers);
  let allowed = allClasses.slice();

  if (lv2) {
    allowed = allowed.filter(function(cl) { return classOffers[cl].LV2.has(lv2); });
  }
  if (opt) {
    allowed = allowed.filter(function(cl) { return classOffers[cl].OPT.has(opt); });
  }

  return allowed;
}
