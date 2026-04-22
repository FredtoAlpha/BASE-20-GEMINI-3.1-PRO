/**
 * ===================================================================
 * TESTS UNITAIRES PHASE 4 & HELPERS
 * ===================================================================
 *
 * Harness minimaliste pour Google Apps Script (pas de framework externe).
 * Exécuter runAllPhase4Tests_() depuis l'éditeur GAS pour lancer toute la
 * suite. Chaque test est isolé : une exception n'arrête pas le reste.
 *
 * Résultat : log dans Stackdriver (Logger.log) + compteur [OK/FAIL/TOTAL].
 *
 * ⚠️ Ces tests couvrent les FONCTIONS PURES (sans I/O Sheet). Les fonctions
 * à effet de bord (Phase4_balanceScoresSwaps_BASEOPTI_V3, etc.) ne sont
 * pas testées unitairement ici — les tester nécessiterait un mock
 * complet de SpreadsheetApp, hors périmètre de cette passe.
 * ===================================================================
 */

// ===================================================================
// HARNESS D'ASSERTIONS
// ===================================================================

var _TEST_STATE = { total: 0, passed: 0, failed: 0, errors: [] };

function _assert(cond, msg) {
  if (!cond) {
    throw new Error('ASSERT FAILED: ' + (msg || '(no message)'));
  }
}

function _assertEq(actual, expected, msg) {
  if (actual !== expected) {
    throw new Error('ASSERT_EQ FAILED: ' + (msg || '') +
      ' | actual=' + JSON.stringify(actual) +
      ' expected=' + JSON.stringify(expected));
  }
}

function _assertNear(actual, expected, epsilon, msg) {
  var diff = Math.abs(actual - expected);
  if (isNaN(actual) || diff > epsilon) {
    throw new Error('ASSERT_NEAR FAILED: ' + (msg || '') +
      ' | actual=' + actual + ' expected=' + expected + ' ε=' + epsilon);
  }
}

function _assertThrows(fn, msg) {
  var threw = false;
  try { fn(); } catch (e) { threw = true; }
  if (!threw) throw new Error('ASSERT_THROWS FAILED: ' + (msg || 'expected exception'));
}

function _runTest(name, fn) {
  _TEST_STATE.total++;
  try {
    fn();
    _TEST_STATE.passed++;
    Logger.log('  ✅ ' + name);
  } catch (e) {
    _TEST_STATE.failed++;
    _TEST_STATE.errors.push({ name: name, message: e && e.message });
    Logger.log('  ❌ ' + name + ' — ' + (e && e.message));
  }
}

// ===================================================================
// TESTS : calculateGlobalStats_Ultimate (Phase4_Ultimate.js)
// ===================================================================

function _testCalcStats_CohorteVide() {
  var r = calculateGlobalStats_Ultimate([]);
  _assertEq(r.ratioF, 0.5, 'ratioF défaut');
  _assertEq(r.avgCOM, 2.5, 'avgCOM défaut');
  _assertEq(r.avgABS, 2.5, 'avgABS défaut');
}

function _testCalcStats_ValeursNormales() {
  var data = [
    { sexe: 'F', COM: 3, TRA: 4, PART: 2, ABS: 1 },
    { sexe: 'M', COM: 1, TRA: 2, PART: 3, ABS: 4 }
  ];
  var r = calculateGlobalStats_Ultimate(data);
  _assertEq(r.ratioF, 0.5, 'ratioF 1/2');
  _assertNear(r.avgCOM, 2, 1e-9, 'avgCOM=(3+1)/2');
  _assertNear(r.avgTRA, 3, 1e-9, 'avgTRA=(4+2)/2');
  _assertNear(r.avgPART, 2.5, 1e-9, 'avgPART=(2+3)/2');
  _assertNear(r.avgABS, 2.5, 1e-9, 'avgABS=(1+4)/2');
}

function _testCalcStats_AbsZeroReel() {
  // Correction du piège 0-as-falsy : un élève avec ABS=0 doit compter pour 0,
  // pas être remplacé par le défaut 2.
  var data = [
    { sexe: 'F', COM: 4, TRA: 4, PART: 4, ABS: 0 },
    { sexe: 'F', COM: 4, TRA: 4, PART: 4, ABS: 0 }
  ];
  var r = calculateGlobalStats_Ultimate(data);
  _assertNear(r.avgABS, 0, 1e-9, 'ABS=0 réel préservé (pas de || 2)');
}

function _testCalcStats_NaNResilience() {
  var data = [
    { sexe: 'F', COM: NaN, TRA: 3, PART: undefined, ABS: null }
  ];
  var r = calculateGlobalStats_Ultimate(data);
  _assert(!isNaN(r.avgCOM), 'avgCOM non-NaN malgré NaN entrant');
  _assert(!isNaN(r.avgTRA), 'avgTRA non-NaN');
  _assert(!isNaN(r.avgPART), 'avgPART non-NaN malgré undefined');
  _assert(!isNaN(r.avgABS), 'avgABS non-NaN malgré null');
}

function _testCalcStats_ChampsAbsents() {
  // Aucun champ défini → fallback 2 partout
  var data = [{ sexe: 'M' }, { sexe: 'F' }];
  var r = calculateGlobalStats_Ultimate(data);
  _assertNear(r.avgCOM, 2, 1e-9);
  _assertNear(r.avgTRA, 2, 1e-9);
  _assertNear(r.avgPART, 2, 1e-9);
  _assertNear(r.avgABS, 2, 1e-9);
}

function _testCalcStats_RatioF_100pc() {
  var data = [{ sexe: 'F', COM: 2 }, { sexe: 'F', COM: 2 }, { sexe: 'F', COM: 2 }];
  var r = calculateGlobalStats_Ultimate(data);
  _assertEq(r.ratioF, 1, '100% filles');
}

function _testCalcStats_RatioF_0pc() {
  var data = [{ sexe: 'M', COM: 2 }, { sexe: 'M', COM: 2 }];
  var r = calculateGlobalStats_Ultimate(data);
  _assertEq(r.ratioF, 0, '0% filles');
}

// ===================================================================
// TESTS : isKnownOPT & isLV2OPTCompatible (App.HarmonyConstants.js)
// ===================================================================

function _testIsKnownOPT_Connues() {
  _assertEq(isKnownOPT('LATIN'), true, 'LATIN est OPT');
  _assertEq(isKnownOPT('CHAV'), true, 'CHAV est OPT');
  _assertEq(isKnownOPT('GREC'), true, 'GREC est OPT');
}

function _testIsKnownOPT_Inconnues() {
  _assertEq(isKnownOPT('INCONNUE'), false, 'INCONNUE n\'est pas OPT');
  _assertEq(isKnownOPT(''), false, 'chaîne vide');
  _assertEq(isKnownOPT('ita'), false, 'casse respectée (ita minuscule)');
}

function _testIsLV2OPT_ITAxCHAV_Interdit() {
  _assertEq(isLV2OPTCompatible('ITA', 'CHAV'), false, 'ITA+CHAV interdit');
}

function _testIsLV2OPT_ITAxLATIN_OK() {
  _assertEq(isLV2OPTCompatible('ITA', 'LATIN'), true, 'ITA+LATIN OK');
  _assertEq(isLV2OPTCompatible('ITA', 'GREC'), true, 'ITA+GREC OK');
}

function _testIsLV2OPT_VideOK() {
  _assertEq(isLV2OPTCompatible('', 'CHAV'), true, 'LV2 vide = OK');
  _assertEq(isLV2OPTCompatible('ITA', ''), true, 'OPT vide = OK');
  _assertEq(isLV2OPTCompatible('', ''), true, 'les deux vides = OK');
}

// ===================================================================
// TESTS : createRNG (App.HarmonyConstants.js) — reproductibilité
// ===================================================================

function _testRNG_SameSeedSameSequence() {
  var r1 = createRNG(42);
  var r2 = createRNG(42);
  for (var i = 0; i < 10; i++) {
    _assertEq(r1.next(), r2.next(), 'seed=42 itération ' + i);
  }
}

function _testRNG_DifferentSeedDifferentSequence() {
  var r1 = createRNG(42);
  var r2 = createRNG(43);
  var collisions = 0;
  for (var i = 0; i < 20; i++) {
    if (r1.next() === r2.next()) collisions++;
  }
  _assert(collisions < 5, 'seeds différentes → séquences distinctes (collisions=' + collisions + ')');
}

function _testRNG_Range() {
  var r = createRNG(123);
  for (var i = 0; i < 100; i++) {
    var v = r.next();
    _assert(v >= 0 && v < 1, 'RNG.next() dans [0,1[ : ' + v);
  }
}

// ===================================================================
// TESTS : budget adaptatif (règle de calcul)
// ===================================================================

function _computeAdaptiveBudget(N, defaultMaxRestarts) {
  return {
    maxSwaps: Math.max(500, N * 5),
    maxRestarts: N >= 150 ? Math.max(10, defaultMaxRestarts) : defaultMaxRestarts
  };
}

function _testBudget_PetitCohort() {
  var b = _computeAdaptiveBudget(50, 5);
  _assertEq(b.maxSwaps, 500, 'petit cohort : plancher 500 swaps');
  _assertEq(b.maxRestarts, 5, 'petit cohort : pas de bump restarts');
}

function _testBudget_MoyenCohort() {
  var b = _computeAdaptiveBudget(120, 5);
  _assertEq(b.maxSwaps, 600, '120×5=600 swaps');
  _assertEq(b.maxRestarts, 5, '120 < 150 : pas de bump');
}

function _testBudget_GrosCohort() {
  var b = _computeAdaptiveBudget(200, 5);
  _assertEq(b.maxSwaps, 1000, '200×5=1000 swaps');
  _assertEq(b.maxRestarts, 10, '200 >= 150 : bump à 10');
}

function _testBudget_TresGrosCohort() {
  var b = _computeAdaptiveBudget(400, 5);
  _assertEq(b.maxSwaps, 2000, '400×5=2000 swaps');
  _assertEq(b.maxRestarts, 10, 'bump à 10');
}

// ===================================================================
// TESTS : qualité absolue (fonction _computeQualityScore)
// ===================================================================

function _computeQualityScore(errInit, errFinal) {
  if (!isFinite(errInit) || errInit <= 0) return null;
  if (errFinal <= 0) return 100;
  if (errFinal >= errInit) return 0;
  return Math.max(0, Math.min(100, 100 * (1 - errFinal / errInit)));
}

function _testQuality_AmeliorationForte() {
  _assertNear(_computeQualityScore(100, 20), 80, 1e-9, 'erreur divisée par 5 = 80%');
}

function _testQuality_AmeliorationFaible() {
  _assertNear(_computeQualityScore(100, 90), 10, 1e-9, 'amélioration 10%');
}

function _testQuality_OptimumParfait() {
  _assertEq(_computeQualityScore(100, 0), 100, 'erreur 0 = qualité 100');
}

function _testQuality_Degradation() {
  _assertEq(_computeQualityScore(100, 150), 0, 'dégradation → 0');
  _assertEq(_computeQualityScore(100, 100), 0, 'égalité → 0');
}

function _testQuality_BaselineInvalide() {
  _assertEq(_computeQualityScore(0, 10), null, 'baseline 0 → null');
  _assertEq(_computeQualityScore(-5, 10), null, 'baseline <0 → null');
  _assertEq(_computeQualityScore(Infinity, 10), null, 'baseline Infinity → null');
}

// ===================================================================
// TESTS : logique de rollback Phase 4 (simulation algébrique)
// ===================================================================
// On ne peut pas exécuter Phase4_balanceScoresSwaps_BASEOPTI_V3 sans
// Spreadsheet. On teste la RÈGLE : bestError >= initialError → rollback.

function _testRollback_Amelioration() {
  var initialError = 100;
  var bestError = 80;
  var shouldRollback = bestError >= initialError;
  _assertEq(shouldRollback, false, 'amélioration : pas de rollback');
}

function _testRollback_Degradation() {
  var initialError = 100;
  var bestError = 120;
  var shouldRollback = bestError >= initialError;
  _assertEq(shouldRollback, true, 'dégradation : rollback');
}

function _testRollback_Egalite() {
  var initialError = 100;
  var bestError = 100;
  var shouldRollback = bestError >= initialError;
  _assertEq(shouldRollback, true, 'égalité : rollback (on ne bouge pas pour rien)');
}

function _testRollback_InitialInfinity() {
  // Si baseline échoue, initialError=Infinity → jamais de rollback
  var initialError = Infinity;
  var bestError = 50;
  var shouldRollback = isFinite(initialError) && bestError >= initialError;
  _assertEq(shouldRollback, false, 'baseline inconnue : on accepte le résultat');
}

// ===================================================================
// POINT D'ENTRÉE : runAllPhase4Tests_
// ===================================================================

function runAllPhase4Tests_() {
  _TEST_STATE = { total: 0, passed: 0, failed: 0, errors: [] };

  Logger.log('═══════════════════════════════════════════════════════════');
  Logger.log('  SUITE DE TESTS — PHASE 4 & HELPERS PURS');
  Logger.log('═══════════════════════════════════════════════════════════');

  Logger.log('\n📊 calculateGlobalStats_Ultimate :');
  _runTest('cohorte vide', _testCalcStats_CohorteVide);
  _runTest('valeurs normales', _testCalcStats_ValeursNormales);
  _runTest('ABS=0 réel non confondu avec absence', _testCalcStats_AbsZeroReel);
  _runTest('résilience NaN/undefined/null', _testCalcStats_NaNResilience);
  _runTest('tous champs absents', _testCalcStats_ChampsAbsents);
  _runTest('ratio F = 100%', _testCalcStats_RatioF_100pc);
  _runTest('ratio F = 0%', _testCalcStats_RatioF_0pc);

  Logger.log('\n🏷️ isKnownOPT / isLV2OPTCompatible :');
  _runTest('OPT connues', _testIsKnownOPT_Connues);
  _runTest('OPT inconnues', _testIsKnownOPT_Inconnues);
  _runTest('ITA+CHAV interdit', _testIsLV2OPT_ITAxCHAV_Interdit);
  _runTest('ITA+LATIN/GREC OK', _testIsLV2OPT_ITAxLATIN_OK);
  _runTest('LV2/OPT vide OK', _testIsLV2OPT_VideOK);

  Logger.log('\n🎲 createRNG :');
  _runTest('même seed = même séquence', _testRNG_SameSeedSameSequence);
  _runTest('seeds différentes distinctes', _testRNG_DifferentSeedDifferentSequence);
  _runTest('sortie dans [0,1[', _testRNG_Range);

  Logger.log('\n⏪ Rollback Phase 4 :');
  _runTest('amélioration : pas de rollback', _testRollback_Amelioration);
  _runTest('dégradation : rollback', _testRollback_Degradation);
  _runTest('égalité : rollback', _testRollback_Egalite);
  _runTest('baseline inconnue : pas de rollback', _testRollback_InitialInfinity);

  Logger.log('\n⚙️ Budget adaptatif :');
  _runTest('petit cohort (N=50)', _testBudget_PetitCohort);
  _runTest('moyen cohort (N=120)', _testBudget_MoyenCohort);
  _runTest('gros cohort (N=200)', _testBudget_GrosCohort);
  _runTest('très gros cohort (N=400)', _testBudget_TresGrosCohort);

  Logger.log('\n⭐ Score qualité absolu :');
  _runTest('amélioration forte', _testQuality_AmeliorationForte);
  _runTest('amélioration faible', _testQuality_AmeliorationFaible);
  _runTest('optimum parfait', _testQuality_OptimumParfait);
  _runTest('dégradation / égalité', _testQuality_Degradation);
  _runTest('baseline invalide', _testQuality_BaselineInvalide);

  Logger.log('\n═══════════════════════════════════════════════════════════');
  Logger.log('  RÉSULTAT : ' + _TEST_STATE.passed + '/' + _TEST_STATE.total +
    ' (échecs : ' + _TEST_STATE.failed + ')');
  Logger.log('═══════════════════════════════════════════════════════════');

  if (_TEST_STATE.failed > 0) {
    Logger.log('\n❌ ÉCHECS :');
    _TEST_STATE.errors.forEach(function(e) {
      Logger.log('  • ' + e.name + ' : ' + e.message);
    });
  }

  return {
    total: _TEST_STATE.total,
    passed: _TEST_STATE.passed,
    failed: _TEST_STATE.failed,
    errors: _TEST_STATE.errors
  };
}
