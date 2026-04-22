/**
 * ===================================================================
 * APP.CONFIG — Facade unifiée de configuration
 * ===================================================================
 *
 * Problème : la configuration du projet est éclatée entre 4 sources :
 *   1. Config.js (constante CONFIG + getConfig())
 *   2. Onglet _CONFIG (lu via lireTousLesParametresConfig)
 *   3. Scoring_Config.js (getScoringConfig, getScoringMode)
 *   4. OptiConfig_System.js (getOptiConfigForUI)
 *
 * Risque : incohérences entre sources, difficulté à tracer la valeur
 * effective d'un paramètre.
 *
 * Solution : getAppConfig(opts) est le POINT D'ACCÈS UNIQUE qui
 * fusionne ces sources de manière contrôlée. Les anciennes fonctions
 * restent disponibles (non-breaking), mais le nouveau code doit
 * utiliser getAppConfig().
 *
 * Ordre de précédence (la source la plus spécifique gagne) :
 *   Config.js (base)
 *     ← onglet _CONFIG (surcharge admin)
 *     ← section scoring (si opts.scoring === true)
 *     ← section opti (si opts.opti === true)
 * ===================================================================
 */

/**
 * Retourne la configuration unifiée de l'application.
 *
 * @param {Object} [opts] - Options de construction
 * @param {boolean} [opts.scoring=false] - Inclure la config scoring
 * @param {string}  [opts.scoringNiveau] - Niveau ciblé ('6°', etc.). Défaut : cfg.NIVEAU.
 * @param {boolean} [opts.opti=false] - Inclure la config optimisation
 * @param {boolean} [opts.skipSheet=false] - Ne pas lire l'onglet _CONFIG
 *                                           (utile pour les tests hors-GAS)
 * @returns {Object} configuration fusionnée
 *
 * @example
 *   // Lecture simple
 *   var cfg = getAppConfig();
 *   Logger.log(cfg.NB_DEST);
 *
 * @example
 *   // Avec scoring pour un niveau donné
 *   var cfg = getAppConfig({ scoring: true, scoringNiveau: '6°' });
 *   Logger.log(cfg.SCORING.mode);
 */
function getAppConfig(opts) {
  opts = opts || {};
  var cfg = {};

  // 1. Base : Config.js
  if (typeof getConfig === 'function') {
    try {
      var base = getConfig();
      if (base && typeof base === 'object') cfg = base;
    } catch (e) {
      if (typeof logLine === 'function') logLine('WARN', 'getAppConfig : getConfig() failed — ' + (e && e.message));
    }
  }

  // 2. Surcouche : onglet _CONFIG (admin-editable)
  if (!opts.skipSheet && typeof lireTousLesParametresConfig === 'function') {
    try {
      var sheetParams = lireTousLesParametresConfig();
      if (sheetParams && typeof sheetParams === 'object') {
        for (var k in sheetParams) {
          if (Object.prototype.hasOwnProperty.call(sheetParams, k)) {
            cfg[k] = sheetParams[k];
          }
        }
      }
    } catch (e) {
      if (typeof logLine === 'function') logLine('WARN', 'getAppConfig : lecture _CONFIG failed — ' + (e && e.message));
    }
  }

  // 3. Scoring (opt-in)
  if (opts.scoring === true && typeof getScoringConfig === 'function') {
    try {
      var niveau = opts.scoringNiveau || cfg.NIVEAU;
      cfg.SCORING = getScoringConfig(niveau);
    } catch (e) {
      if (typeof logLine === 'function') logLine('WARN', 'getAppConfig : scoring failed — ' + (e && e.message));
    }
  }

  // 4. Opti (opt-in)
  if (opts.opti === true && typeof getOptiConfigForUI === 'function') {
    try {
      var optiResp = getOptiConfigForUI();
      if (optiResp && optiResp.success && optiResp.config) {
        cfg.OPTI = optiResp.config;
      }
    } catch (e) {
      if (typeof logLine === 'function') logLine('WARN', 'getAppConfig : opti failed — ' + (e && e.message));
    }
  }

  return cfg;
}
