/**
 * ===================================================================
 * 📋 PRIME LEGACY - MENU GOOGLE SHEETS
 * ===================================================================
 *
 * @deprecated Ce menu n'est PAS appelé depuis Code.js > onOpen().
 * L'entrée officielle est Console V3 (menu "PILOTAGE CLASSE").
 * Conservé uniquement pour les fonctions utilitaires
 * legacy_viewSourceClasses_PRIME et legacy_viewTestResults_PRIME
 * qui sont référencées depuis d'autres fichiers.
 *
 * Date : 2025-11-13 (nettoyé 2026-03-06)
 * ===================================================================
 */

/**
 * @deprecated Menu LEGACY non utilisé. L'entrée officielle est Console V3.
 * Orphan callbacks retirés : showPilotageConsole, legacy_showPipelineStatus,
 * legacy_runPhase[1-4]_PRIME, legacy_runJulesCodex_Menu.
 */
function createLegacyMenu_PRIME() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('⚙️ PRIME LEGACY (deprecated)')
    .addItem('Lancer Interface Moderne', 'showLegacyInterface')
    .addItem('Diagnostic Pré-Lancement', 'legacy_runDiagnostic_Menu')
    .addSeparator()
    .addItem('Pipeline Complet', 'legacy_runFullPipeline_PRIME')
    .addSeparator()
    .addItem('Voir Classes Sources', 'legacy_viewSourceClasses_PRIME')
    .addItem('Voir Résultats TEST', 'legacy_viewTestResults_PRIME')
    .addSeparator()
    .addSubMenu(ui.createMenu('Logs')
      .addItem('Ouvrir Logs', 'openLegacyLogsSheet')
      .addItem('Afficher Derniers Logs', 'showRecentLegacyLogs')
      .addItem('Exporter Logs', 'exportLegacyLogsToFile')
      .addItem('Effacer Logs', 'clearLegacyLogs'))
    .addToUi();

  logLine('INFO', '✅ Menu PRIME LEGACY créé (deprecated - utiliser Console V3)');
}

/**
 * Affiche les classes sources détectées
 */
function legacy_viewSourceClasses_PRIME() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();

    const sourceSheets = allSheets.filter(function(s) {
      // Support: 6°1, GAMARRA°4, etc. (toujours avec °)
      return /^[A-Za-z0-9_-]+°\d+$/.test(s.getName());
    });

    sourceSheets.sort(function(a, b) {
      return a.getName().localeCompare(b.getName());
    });

    // Compter les élèves par source
    let details = '';
    let totalEleves = 0;

    sourceSheets.forEach(function(s) {
      const numEleves = Math.max(0, s.getLastRow() - 1);
      totalEleves += numEleves;
      details += '• ' + s.getName() + ' : ' + numEleves + ' élèves\n';
    });

    ui.alert(
      '📋 Classes Sources Détectées',
      'ONGLETS SOURCES (' + sourceSheets.length + ') :\n\n' +
      details +
      '\nTOTAL : ' + totalEleves + ' élèves\n\n' +
      (sourceSheets.length > 0
        ? '✅ Prêt à lancer le pipeline LEGACY'
        : '⚠️ Aucun onglet source trouvé'),
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Erreur', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Affiche les résultats dans les onglets TEST
 */
function legacy_viewTestResults_PRIME() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();

    const testSheets = allSheets.filter(function(s) {
      return s.getName().endsWith('TEST');
    });

    testSheets.sort(function(a, b) {
      return a.getName().localeCompare(b.getName());
    });

    if (testSheets.length === 0) {
      ui.alert(
        '⚠️ Aucun Résultat TEST',
        'Aucun onglet TEST trouvé.\n\n' +
        'Lancez d\'abord le pipeline LEGACY pour créer les onglets TEST.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Compter les élèves par TEST
    let details = '';
    let totalEleves = 0;
    let totalAssigned = 0;

    testSheets.forEach(function(s) {
      const numEleves = Math.max(0, s.getLastRow() - 1);
      totalEleves += numEleves;

      // Compter élèves assignés
      if (numEleves > 0) {
        const data = s.getDataRange().getValues();
        const headers = data[0];
        const idxAssigned = headers.indexOf('_CLASS_ASSIGNED');

        if (idxAssigned >= 0) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxAssigned] || '').trim()) {
              totalAssigned++;
            }
          }
        }
      }

      details += '• ' + s.getName() + ' : ' + numEleves + ' élèves\n';
    });

    const pctAssigned = totalEleves > 0
      ? ((totalAssigned / totalEleves) * 100).toFixed(1)
      : 0;

    ui.alert(
      '📊 Résultats TEST',
      'ONGLETS TEST (' + testSheets.length + ') :\n\n' +
      details +
      '\nTOTAL : ' + totalEleves + ' élèves\n' +
      'ASSIGNÉS : ' + totalAssigned + ' (' + pctAssigned + '%)\n\n' +
      '✅ Pipeline exécuté avec succès',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Erreur', e.toString(), ui.ButtonSet.OK);
  }
}
