// =============================================================================
// RunAudit_Utils.js — Utilitaires d'audit de run (runId, timer, rapport)
// =============================================================================
// Fournit :
//   - RunAudit_createId()       : génère un identifiant unique de run
//   - RunAudit_startTimer()     : lance un chrono
//   - RunAudit_stopTimer(timer) : arrête le chrono et retourne durationMs
//   - RunAudit_buildReport(ctx) : construit un rapport post-run standard
//   - RunAudit_log(runId, msg)  : log avec préfixe runId
// =============================================================================

/**
 * Génère un identifiant de run unique.
 * Format : RUN-YYYYMMDD-HHmmss-XXXX  (XXXX = 4 hex aléatoires)
 * @returns {string}
 */
function RunAudit_createId() {
  var now = new Date();
  var pad2 = function(n) { return n < 10 ? '0' + n : '' + n; };
  var datePart = '' + now.getFullYear() + pad2(now.getMonth() + 1) + pad2(now.getDate());
  var timePart = pad2(now.getHours()) + pad2(now.getMinutes()) + pad2(now.getSeconds());
  var rand = Math.floor(Math.random() * 65536).toString(16);
  while (rand.length < 4) rand = '0' + rand;
  return 'RUN-' + datePart + '-' + timePart + '-' + rand.toUpperCase();
}

/**
 * Démarre un timer de run.
 * @returns {Object} objet timer à passer à RunAudit_stopTimer()
 */
function RunAudit_startTimer() {
  return { startMs: Date.now() };
}

/**
 * Arrête le timer et retourne la durée en ms.
 * @param {Object} timer - objet retourné par RunAudit_startTimer()
 * @returns {number} durée en millisecondes
 */
function RunAudit_stopTimer(timer) {
  return Date.now() - (timer && timer.startMs ? timer.startMs : Date.now());
}

/**
 * Formate une durée en ms en chaîne lisible.
 * @param {number} ms
 * @returns {string} ex: "2.3s" ou "145ms"
 */
function RunAudit_formatDuration(ms) {
  if (ms >= 1000) return (ms / 1000).toFixed(1) + 's';
  return ms + 'ms';
}

/**
 * Log un message avec le préfixe runId.
 * @param {string} runId
 * @param {string} level - INFO, WARN, ERROR
 * @param {string} msg
 */
function RunAudit_log(runId, level, msg) {
  logLine(level, '[' + runId + '] ' + msg);
}

/**
 * Construit un rapport post-run standard.
 *
 * @param {Object} ctx - Contexte du run
 * @param {string} ctx.runId       - identifiant du run
 * @param {string} ctx.operation   - nom de l'opération (IMPORT, SCORING, PIPELINE)
 * @param {number} ctx.durationMs  - durée du run en ms
 * @param {boolean} ctx.success    - succès ou échec
 * @param {string} [ctx.error]     - message d'erreur si échec
 *
 * Import-specific (optionnels) :
 * @param {number} [ctx.totalEleves]    - nombre total d'élèves
 * @param {number} [ctx.notesMatched]   - notes matchées
 * @param {number} [ctx.notesTotal]     - notes totales
 * @param {string[]} [ctx.notesUnmatched] - noms non matchés
 * @param {number} [ctx.absMatched]     - absences matchées
 * @param {number} [ctx.absTotal]       - absences totales
 * @param {number} [ctx.obsMatched]     - observations matchées
 * @param {number} [ctx.obsTotal]       - observations totales
 * @param {number} [ctx.punMatched]     - punitions matchées
 * @param {number} [ctx.punTotal]       - punitions totales
 * @param {number} [ctx.incMatched]     - incidents matchés
 * @param {number} [ctx.incTotal]       - incidents totaux
 * @param {string[]} [ctx.classesList]  - liste des classes
 * @param {number} [ctx.anomaliesLV2]   - élèves sans LV2
 * @param {number} [ctx.anomaliesOPT]   - élèves sans OPT (quand attendu)
 *
 * Scoring-specific (optionnels) :
 * @param {number} [ctx.scoresTotal]        - nombre d'élèves scorés
 * @param {Object} [ctx.distribution]       - {1:n, 2:n, 3:n, 4:n}
 *
 * Pipeline-specific (optionnels) :
 * @param {number} [ctx.swapsApplied]   - nombre de swaps
 *
 * @returns {Object} rapport structuré
 */
function RunAudit_buildReport(ctx) {
  var report = {
    runId: ctx.runId || 'UNKNOWN',
    operation: ctx.operation || 'UNKNOWN',
    timestamp: new Date().toISOString(),
    durationMs: ctx.durationMs || 0,
    durationHuman: RunAudit_formatDuration(ctx.durationMs || 0),
    success: !!ctx.success
  };

  if (ctx.error) report.error = ctx.error;

  // --- Matching (import) ---
  if (ctx.totalEleves !== undefined) {
    report.matching = {
      eleves: ctx.totalEleves,
      notes: { matched: ctx.notesMatched || 0, total: ctx.notesTotal || 0, unmatched: ctx.notesUnmatched || [] },
      absences: { matched: ctx.absMatched || 0, total: ctx.absTotal || 0 },
      observations: { matched: ctx.obsMatched || 0, total: ctx.obsTotal || 0 },
      punitions: { matched: ctx.punMatched || 0, total: ctx.punTotal || 0 },
      incidents: { matched: ctx.incMatched || 0, total: ctx.incTotal || 0 }
    };
  }

  // --- Anomalies LV2/OPT ---
  if (ctx.anomaliesLV2 !== undefined || ctx.anomaliesOPT !== undefined) {
    report.anomalies = {
      lv2Vide: ctx.anomaliesLV2 || 0,
      optVide: ctx.anomaliesOPT || 0
    };
  }

  // --- Classes ---
  if (ctx.classesList) {
    report.classes = ctx.classesList;
  }

  // --- Scoring ---
  if (ctx.distribution) {
    report.scoring = {
      total: ctx.scoresTotal || 0,
      distribution: ctx.distribution
    };
  }

  // --- Pipeline ---
  if (ctx.swapsApplied !== undefined) {
    report.pipeline = { swapsApplied: ctx.swapsApplied };
  }

  // Log le rapport complet
  RunAudit_log(report.runId, 'INFO', report.operation + ' ' +
    (report.success ? 'OK' : 'FAIL') + ' in ' + report.durationHuman);

  return report;
}
