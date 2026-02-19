/**
 * ===================================================================
 * LOGGER SERVER RECEIVER - Reçoit les logs du client
 * ===================================================================
 *
 * Fonction appelée par Logger_Client.html via google.script.run
 * Enregistre les logs client dans StackDriver et _ERROR_LOG
 */

/**
 * Reçoit un log du client et l'enregistre côté serveur
 * @param {Object} payload - Log payload depuis le client
 * @param {string} payload.timestamp - Timestamp ISO
 * @param {string} payload.level - Niveau de log
 * @param {string} payload.message - Message
 * @param {string} payload.data - Données contextuelles (JSON string)
 * @param {string} payload.error - Erreur (string)
 * @param {string} payload.userAgent - User agent du navigateur
 * @param {string} payload.url - URL de la page
 */
function logFromClient(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      Logger.warn('logFromClient: Invalid payload received', { payload: payload });
      return;
    }

    // Mapper le niveau
    var levelMap = {
      'TRACE': Logger.LEVELS.TRACE,
      'DEBUG': Logger.LEVELS.DEBUG,
      'INFO': Logger.LEVELS.INFO,
      'WARN': Logger.LEVELS.WARN,
      'ERROR': Logger.LEVELS.ERROR
    };

    var level = levelMap[payload.level] || Logger.LEVELS.INFO;
    var message = '[CLIENT] ' + (payload.message || 'No message');

    // Construire les données contextuelles
    var contextData = {
      clientTimestamp: payload.timestamp,
      userAgent: payload.userAgent,
      url: payload.url
    };

    // Ajouter les données du payload si présentes
    if (payload.data) {
      try {
        contextData.clientData = JSON.parse(payload.data);
      } catch (e) {
        contextData.clientData = payload.data; // Garder en string si parsing échoue
      }
    }

    // Logger selon le niveau
    if (level === Logger.LEVELS.ERROR) {
      Logger.error(message, contextData, payload.error ? new Error(payload.error) : null);
    } else if (level === Logger.LEVELS.WARN) {
      Logger.warn(message, contextData);
    } else if (level === Logger.LEVELS.INFO) {
      Logger.info(message, contextData);
    } else {
      Logger.debug(message, contextData);
    }

    return { success: true };

  } catch (error) {
    // Fallback en cas d'erreur dans le logger lui-même
    console.error('logFromClient: Critical error', error);
    return { success: false, error: error.message };
  }
}

/**
 * Nettoie la feuille _ERROR_LOG si elle devient trop grande
 */
function cleanupErrorLog() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('_ERROR_LOG');

    if (!sheet) {
      Logger.info('cleanupErrorLog: No _ERROR_LOG sheet found');
      return;
    }

    var lastRow = sheet.getLastRow();
    var maxRows = 1000; // Configurable

    if (lastRow > maxRows) {
      var rowsToDelete = lastRow - maxRows;
      sheet.deleteRows(2, rowsToDelete); // Garder la ligne d'en-tête
      Logger.info('cleanupErrorLog: Deleted old rows', { deleted: rowsToDelete });
    } else {
      Logger.debug('cleanupErrorLog: No cleanup needed', { rows: lastRow, max: maxRows });
    }

  } catch (error) {
    Logger.error('cleanupErrorLog: Failed to cleanup', error);
  }
}

/**
 * Configuration de l'environnement (production vs development)
 * À appeler manuellement ou via trigger
 */
function configureLoggerEnvironment(env) {
  env = env || 'development';

  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('ENVIRONMENT', env);

    if (env === 'production') {
      Logger.enableProduction();
    } else {
      Logger.enableDevelopment();
    }

    Logger.info('Logger environment configured', { environment: env });

  } catch (error) {
    Logger.error('configureLoggerEnvironment: Failed', error);
  }
}

/**
 * Test du système de logging
 */
function testLogger() {
  Logger.info('=== DÉBUT DU TEST LOGGER ===');

  // Test des différents niveaux
  Logger.trace('Test TRACE', { value: 1 });
  Logger.debug('Test DEBUG', { value: 2 });
  Logger.info('Test INFO', { value: 3 });
  Logger.warn('Test WARN', { value: 4 });
  Logger.error('Test ERROR', { value: 5 });

  // Test avec erreur
  try {
    throw new Error('Test exception');
  } catch (error) {
    Logger.error('Exception capturée', { context: 'test' }, error);
  }

  // Test timer
  var timer = Logger.startTimer('testOperation');
  Utilities.sleep(100);
  timer.end('Test opération terminée');

  Logger.info('=== FIN DU TEST LOGGER ===');
}
