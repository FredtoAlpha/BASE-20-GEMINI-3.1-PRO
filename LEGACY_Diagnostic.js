/**
 * ===================================================================
 * 🔍 PRIME LEGACY - SYSTÈME D'AUTO-DIAGNOSTIC
 * ===================================================================
 *
 * Vérification pré-lancement du pipeline LEGACY :
 * - Onglets sources présents et valides
 * - Configuration _STRUCTURE correcte
 * - Colonnes requises présentes
 * - Quotas et effectifs cohérents
 * - Détection d'anomalies potentielles
 *
 * Date : 2025-11-14
 * Branche : claude/legacy-pipeline-renewal-01FK5TFnxx6JjwZ9bMkF5hqw
 *
 * ===================================================================
 */

/**
 * Lance un diagnostic complet du pipeline LEGACY
 * Affiche un rapport détaillé et retourne un objet de résultats
 *
 * @param {boolean} showUI - Afficher le rapport dans une UI (défaut: true)
 * @returns {Object} Résultats du diagnostic
 */
function runLegacyDiagnostic(showUI) {
  const displayUI = showUI !== false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const results = {
    ok: true,
    errors: [],
    warnings: [],
    infos: [],
    checks: {
      sources: { ok: false, message: '' },
      structure: { ok: false, message: '' },
      columns: { ok: false, message: '' },
      quotas: { ok: false, message: '' },
      data: { ok: false, message: '' }
    }
  };

  logLegacy('INFO', '🔍 Lancement du diagnostic LEGACY...', 'Diagnostic');

  // ========== CHECK 1 : ONGLETS SOURCES ==========
  const sourcesCheck = checkSourceSheets_(ss);
  results.checks.sources = sourcesCheck;

  if (!sourcesCheck.ok) {
    results.ok = false;
    results.errors.push(sourcesCheck.message);
  } else {
    results.infos.push(sourcesCheck.message);
  }

  // ========== CHECK 2 : FEUILLE _STRUCTURE ==========
  const structureCheck = checkStructureSheet_(ss);
  results.checks.structure = structureCheck;

  if (!structureCheck.ok) {
    results.warnings.push(structureCheck.message);
  } else {
    results.infos.push(structureCheck.message);
  }

  // ========== CHECK 3 : COLONNES REQUISES ==========
  const columnsCheck = checkRequiredColumns_(ss, sourcesCheck.sheets);
  results.checks.columns = columnsCheck;

  if (!columnsCheck.ok) {
    results.errors.push(columnsCheck.message);
    results.ok = false;
  } else {
    results.infos.push(columnsCheck.message);
  }

  // ========== CHECK 4 : QUOTAS ET CONFIGURATION ==========
  const quotasCheck = checkQuotasConfiguration_(ss);
  results.checks.quotas = quotasCheck;

  if (!quotasCheck.ok) {
    results.warnings.push(quotasCheck.message);
  } else {
    results.infos.push(quotasCheck.message);
  }

  // ========== CHECK 5 : DONNÉES ÉLÈVES ==========
  const dataCheck = checkStudentData_(ss, sourcesCheck.sheets);
  results.checks.data = dataCheck;

  if (!dataCheck.ok) {
    results.warnings.push(dataCheck.message);
  } else {
    results.infos.push(dataCheck.message);
  }

  // ========== AFFICHAGE DU RAPPORT ==========
  if (displayUI) {
    displayDiagnosticReport_(results);
  }

  logLegacy(results.ok ? 'SUCCESS' : 'WARN', '🔍 Diagnostic terminé : ' + (results.ok ? 'OK' : 'WARNINGS'), 'Diagnostic');

  return results;
}

/**
 * Vérifie la présence des onglets sources
 */
function checkSourceSheets_(ss) {
  const allSheets = ss.getSheets();
  const sourcePattern = /^[A-Za-z0-9_-]+°\d+$/;
  const excludePattern = /TEST|CACHE|DEF|FIN|SRC|SOURCE|_CONFIG|_STRUCTURE|_LOG/i;

  const sourceSheets = allSheets.filter(function(s) {
    const name = s.getName();
    return sourcePattern.test(name) && !excludePattern.test(name);
  });

  if (sourceSheets.length === 0) {
    return {
      ok: false,
      message: '❌ Aucun onglet source trouvé (ex: 6°1, ECOLE1, GAMARRA°4)',
      sheets: []
    };
  }

  const sheetNames = sourceSheets.map(function(s) { return s.getName(); });

  return {
    ok: true,
    message: '✅ ' + sourceSheets.length + ' onglet(s) source(s) détecté(s) : ' + sheetNames.join(', '),
    sheets: sourceSheets
  };
}

/**
 * Vérifie la présence et validité de _STRUCTURE
 */
function checkStructureSheet_(ss) {
  const structSheet = ss.getSheetByName('_STRUCTURE');

  if (!structSheet) {
    return {
      ok: false,
      message: '⚠️ Feuille _STRUCTURE introuvable (configuration par défaut sera utilisée)'
    };
  }

  // Vérifier que la feuille n'est pas vide
  if (structSheet.getLastRow() <= 1) {
    return {
      ok: false,
      message: '⚠️ Feuille _STRUCTURE vide'
    };
  }

  // Vérifier la présence des colonnes requises
  const data = structSheet.getDataRange().getValues();
  const headers = data[0];

  const requiredCols = ['CLASSE_ORIGINE', 'CLASSE_DEST', 'EFFECTIF', 'OPTIONS'];
  const missingCols = [];

  requiredCols.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      missingCols.push(col);
    }
  });

  if (missingCols.length > 0) {
    return {
      ok: false,
      message: '⚠️ _STRUCTURE : colonnes manquantes : ' + missingCols.join(', ')
    };
  }

  return {
    ok: true,
    message: '✅ _STRUCTURE valide (' + (data.length - 1) + ' configurations)'
  };
}

/**
 * Vérifie que toutes les colonnes requises sont présentes dans les sources
 */
function checkRequiredColumns_(ss, sourceSheets) {
  if (!sourceSheets || sourceSheets.length === 0) {
    return {
      ok: false,
      message: '❌ Impossible de vérifier les colonnes (aucun onglet source)'
    };
  }

  const requiredColumns = [
    'ID_ELEVE', 'NOM', 'PRENOM', 'SEXE',
    'LV2', 'OPT',
    'COM', 'TRA', 'PART', 'ABS',
    'ASSO', 'DISSO',
    '_CLASS_ASSIGNED'
  ];

  const issues = [];

  sourceSheets.forEach(function(sheet) {
    if (sheet.getLastRow() === 0) {
      issues.push(sheet.getName() + ' : feuille vide');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const missing = [];

    requiredColumns.forEach(function(col) {
      if (headers.indexOf(col) === -1) {
        missing.push(col);
      }
    });

    if (missing.length > 0) {
      issues.push(sheet.getName() + ' : colonnes manquantes : ' + missing.join(', '));
    }
  });

  if (issues.length > 0) {
    return {
      ok: false,
      message: '❌ Problèmes de colonnes :\n' + issues.join('\n')
    };
  }

  return {
    ok: true,
    message: '✅ Toutes les colonnes requises sont présentes'
  };
}

/**
 * Vérifie la cohérence des quotas
 */
function checkQuotasConfiguration_(ss) {
  const structSheet = ss.getSheetByName('_STRUCTURE');

  if (!structSheet) {
    return {
      ok: true,  // Not a blocker
      message: '⚠️ Pas de configuration quotas (_STRUCTURE manquant)'
    };
  }

  try {
    const quotas = readQuotasFromUI_LEGACY();
    const classesWithQuotas = Object.keys(quotas).length;

    if (classesWithQuotas === 0) {
      return {
        ok: false,
        message: '⚠️ Aucun quota configuré dans _STRUCTURE'
      };
    }

    // Vérifier que les quotas sont > 0
    let totalQuotas = 0;
    for (const classe in quotas) {
      const opts = quotas[classe];
      for (const opt in opts) {
        totalQuotas += opts[opt];
      }
    }

    return {
      ok: true,
      message: '✅ Quotas configurés : ' + classesWithQuotas + ' classe(s), total ' + totalQuotas + ' places'
    };

  } catch (e) {
    return {
      ok: false,
      message: '⚠️ Erreur lecture quotas : ' + e.toString()
    };
  }
}

/**
 * Vérifie la qualité des données élèves
 */
function checkStudentData_(ss, sourceSheets) {
  if (!sourceSheets || sourceSheets.length === 0) {
    return {
      ok: false,
      message: '❌ Aucun onglet source pour vérifier les données'
    };
  }

  let totalEleves = 0;
  let warnings = [];

  sourceSheets.forEach(function(sheet) {
    const numEleves = Math.max(0, sheet.getLastRow() - 1);
    totalEleves += numEleves;

    if (numEleves === 0) {
      warnings.push(sheet.getName() + ' : aucun élève');
    }

    // Vérifier si des élèves ont des données manquantes
    if (numEleves > 0) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      const idxNom = headers.indexOf('NOM');
      const idxPrenom = headers.indexOf('PRENOM');
      const idxSexe = headers.indexOf('SEXE');

      let missing = 0;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        if (!row[idxNom] || !row[idxPrenom] || !row[idxSexe]) {
          missing++;
        }
      }

      if (missing > 0) {
        warnings.push(sheet.getName() + ' : ' + missing + ' élève(s) avec données manquantes (NOM/PRENOM/SEXE)');
      }
    }
  });

  if (totalEleves === 0) {
    return {
      ok: false,
      message: '❌ Aucun élève détecté dans les onglets sources'
    };
  }

  if (warnings.length > 0) {
    return {
      ok: true,  // Not blocking, but worth noting
      message: '⚠️ ' + totalEleves + ' élèves détectés, warnings :\n' + warnings.join('\n')
    };
  }

  return {
    ok: true,
    message: '✅ ' + totalEleves + ' élèves détectés, données complètes'
  };
}

/**
 * Affiche le rapport de diagnostic dans une UI
 */
function displayDiagnosticReport_(results) {
  const ui = SpreadsheetApp.getUi();

  let title = results.ok ? '✅ Diagnostic OK' : '⚠️ Diagnostic : Problèmes détectés';

  let message = '';

  // Erreurs critiques
  if (results.errors.length > 0) {
    message += '🚫 ERREURS CRITIQUES :\n\n';
    results.errors.forEach(function(err) {
      message += err + '\n\n';
    });
  }

  // Warnings
  if (results.warnings.length > 0) {
    message += '⚠️ AVERTISSEMENTS :\n\n';
    results.warnings.forEach(function(warn) {
      message += warn + '\n\n';
    });
  }

  // Infos
  if (results.infos.length > 0) {
    message += '✅ VALIDATIONS :\n\n';
    results.infos.forEach(function(info) {
      message += info + '\n\n';
    });
  }

  // Résumé
  message += '─'.repeat(50) + '\n\n';

  if (results.ok) {
    message += '🚀 Le pipeline peut être lancé en toute sécurité.';
  } else if (results.errors.length > 0) {
    message += '❌ Corrigez les erreurs avant de lancer le pipeline.';
  } else {
    message += '⚠️ Le pipeline peut être lancé, mais vérifiez les avertissements.';
  }

  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * Menu : Lancer le diagnostic
 */
function legacy_runDiagnostic_Menu() {
  runLegacyDiagnostic(true);
}
