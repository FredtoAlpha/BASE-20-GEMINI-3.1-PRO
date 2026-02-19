/**
 * ===================================================================
 * DEBUG LOGGER - Diagnostic pour vérifier l'état de Logger
 * ===================================================================
 */

function DEBUG_checkLogger() {
  var report = [];

  report.push('=== DIAGNOSTIC LOGGER ===');
  report.push('');

  // 1. Logger existe-t-il ?
  report.push('1. Logger existe: ' + (typeof Logger !== 'undefined' ? 'OUI ✅' : 'NON ❌'));

  if (typeof Logger !== 'undefined') {
    // 2. Logger est-il un objet ?
    report.push('2. Logger est un objet: ' + (typeof Logger === 'object' ? 'OUI ✅' : 'NON ❌ (type: ' + typeof Logger + ')'));

    // 3. Lister toutes les méthodes de Logger
    report.push('');
    report.push('3. Méthodes disponibles dans Logger:');
    try {
      var methods = Object.keys(Logger);
      if (methods.length === 0) {
        report.push('   ❌ Aucune méthode trouvée !');
      } else {
        methods.forEach(function(method) {
          var type = typeof Logger[method];
          var icon = type === 'function' ? '✅' : '⚠️';
          report.push('   ' + icon + ' ' + method + ' (' + type + ')');
        });
      }
    } catch (e) {
      report.push('   ❌ Erreur: ' + e.toString());
    }

    // 4. Vérifier spécifiquement Logger.log
    report.push('');
    report.push('4. Logger.log existe: ' + (typeof Logger.log !== 'undefined' ? 'OUI ✅' : 'NON ❌'));
    report.push('   Type: ' + typeof Logger.log);

    // 5. Vérifier les autres méthodes
    report.push('');
    report.push('5. Autres méthodes:');
    ['trace', 'debug', 'info', 'warn', 'error'].forEach(function(method) {
      report.push('   Logger.' + method + ': ' + typeof Logger[method]);
    });

    // 6. Essayer d'appeler Logger.log
    report.push('');
    report.push('6. Test d\'appel Logger.log():');
    try {
      Logger.log('Test message');
      report.push('   ✅ Succès !');
    } catch (e) {
      report.push('   ❌ ERREUR: ' + e.toString());
    }
  }

  // Afficher le rapport
  var fullReport = report.join('\n');
  console.log(fullReport);

  // Retourner aussi pour l'interface
  return fullReport;
}

/**
 * Version pour Interface V2
 */
function DEBUG_checkLoggerForInterface() {
  try {
    var report = DEBUG_checkLogger();
    return {
      success: true,
      report: report
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
      stack: e.stack
    };
  }
}
