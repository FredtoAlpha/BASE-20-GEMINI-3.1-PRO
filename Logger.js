/**
 * ===================================================================
 * LOGGER CENTRALISÉ - Production-ready logging system
 * ===================================================================
 *
 * Features:
 * - 5 log levels: TRACE, DEBUG, INFO, WARN, ERROR
 * - Configurable log level (production vs development)
 * - Error tracking via StackDriver + Sheet logging
 * - Timestamp + context + structured logging
 * - Performance metrics
 * - Memory-safe circular reference handling
 *
 * Usage:
 *   Logger.debug('User action', { userId: 123 });
 *   Logger.info('Process completed', { duration: 1234 });
 *   Logger.warn('Slow query detected', { query: 'SELECT...' });
 *   Logger.error('Fatal error', error, { context: 'Phase1' });
 *
 * Configuration:
 *   Logger.setLevel(Logger.LEVELS.INFO); // Production
 *   Logger.setLevel(Logger.LEVELS.DEBUG); // Development
 */

// ===================================================================
// CONFIGURATION GLOBALE
// ===================================================================

var Logger = (function() {
  // NOTE: 'use strict' retiré car incompatible avec exports globaux Apps Script
  // En mode strict, this = undefined dans scope global, empêchant this.Logger = Logger

  // Niveaux de log (plus bas = plus verbeux)
  var LEVELS = {
    TRACE: 0,  // Debug ultra-détaillé (ex: chaque iteration de boucle)
    DEBUG: 1,  // Debug général (ex: "Entering function X")
    INFO: 2,   // Informations importantes (ex: "Process started")
    WARN: 3,   // Avertissements (ex: "Slow query detected")
    ERROR: 4   // Erreurs critiques (ex: "Fatal exception")
  };

  var LEVEL_NAMES = ['TRACE', 'DEBUG', 'INFO', 'WARN', 'ERROR'];

  // Configuration par défaut
  var config = {
    currentLevel: LEVELS.INFO,        // Niveau minimum à logger (production)
    enableConsole: true,              // Activer console.log
    enableStackDriver: true,          // Activer StackDriver (Apps Script)
    enableSheetLogging: false,        // Activer logging dans feuille (optionnel)
    errorSheetName: '_ERROR_LOG',     // Nom de la feuille pour erreurs
    maxSheetRows: 1000,               // Limite de lignes dans feuille
    timestampFormat: 'ISO',           // 'ISO' ou 'LOCALE'
    includeStackTrace: true,          // Inclure stack trace pour ERROR
    contextPrefix: '',                // Préfixe optionnel (ex: '[PROD]')
  };

  // ===================================================================
  // UTILITIES
  // ===================================================================

  /**
   * Formatte un timestamp
   */
  function formatTimestamp() {
    var now = new Date();
    if (config.timestampFormat === 'ISO') {
      return now.toISOString();
    }
    return now.toLocaleString('fr-FR');
  }

  /**
   * Sérialise un objet en JSON (gère les références circulaires)
   */
  function safeStringify(obj, maxDepth) {
    maxDepth = maxDepth || 3;
    var seen = [];

    return JSON.stringify(obj, function(key, value) {
      // Limiter la profondeur
      if (typeof value === 'object' && value !== null) {
        if (seen.indexOf(value) !== -1) {
          return '[Circular]';
        }
        seen.push(value);
      }
      return value;
    }, 2);
  }

  /**
   * Formatte un message de log
   */
  function formatMessage(level, message, data, error) {
    var timestamp = formatTimestamp();
    var levelName = LEVEL_NAMES[level];
    var prefix = config.contextPrefix ? config.contextPrefix + ' ' : '';

    var parts = [
      '[' + timestamp + ']',
      prefix + '[' + levelName + ']',
      message
    ];

    // Ajouter les données contextuelles
    if (data && Object.keys(data).length > 0) {
      try {
        parts.push('| Data: ' + safeStringify(data));
      } catch (e) {
        parts.push('| Data: [Serialization Error]');
      }
    }

    // Ajouter l'erreur si présente
    if (error) {
      if (error instanceof Error) {
        parts.push('| Error: ' + error.message);
        if (config.includeStackTrace && error.stack) {
          parts.push('| Stack: ' + error.stack);
        }
      } else {
        try {
          parts.push('| Error: ' + safeStringify(error));
        } catch (e) {
          parts.push('| Error: [Unknown]');
        }
      }
    }

    return parts.join(' ');
  }

  /**
   * Tronque un message si trop long
   */
  function truncate(str, maxLength) {
    maxLength = maxLength || 5000;
    if (str.length <= maxLength) return str;
    return str.substring(0, maxLength) + '... [truncated]';
  }

  // ===================================================================
  // OUTPUTS
  // ===================================================================

  /**
   * Log vers la console
   */
  function logToConsole(level, formattedMessage) {
    if (!config.enableConsole) return;

    var consoleMethod = console.log; // Par défaut

    if (level === LEVELS.ERROR && console.error) {
      consoleMethod = console.error;
    } else if (level === LEVELS.WARN && console.warn) {
      consoleMethod = console.warn;
    } else if (level === LEVELS.INFO && console.info) {
      consoleMethod = console.info;
    } else if (level === LEVELS.DEBUG && console.debug) {
      consoleMethod = console.debug;
    }

    consoleMethod(formattedMessage);
  }

  /**
   * Log vers StackDriver (Google Cloud Logging)
   * Disponible automatiquement dans Apps Script
   */
  function logToStackDriver(level, message, data, error) {
    if (!config.enableStackDriver) return;

    try {
      // Apps Script utilise console.log/warn/error pour StackDriver
      // Les messages apparaissent dans Cloud Console > Logging

      var payload = {
        timestamp: new Date().toISOString(),
        level: LEVEL_NAMES[level],
        message: message,
        data: data || null,
        error: error ? (error.message || String(error)) : null
      };

      if (level === LEVELS.ERROR) {
        console.error('[STACKDRIVER]', JSON.stringify(payload));
      } else if (level === LEVELS.WARN) {
        console.warn('[STACKDRIVER]', JSON.stringify(payload));
      } else {
        console.log('[STACKDRIVER]', JSON.stringify(payload));
      }
    } catch (e) {
      // Fallback silencieux
      console.error('Logger: Failed to log to StackDriver', e);
    }
  }

  /**
   * Log vers une feuille Google Sheets
   */
  function logToSheet(level, message, data, error) {
    if (!config.enableSheetLogging) return;
    if (level < LEVELS.WARN) return; // Seulement WARN et ERROR dans la feuille

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return;

      var sheet = ss.getSheetByName(config.errorSheetName);

      // Créer la feuille si elle n'existe pas
      if (!sheet) {
        sheet = ss.insertSheet(config.errorSheetName);
        sheet.getRange(1, 1, 1, 5).setValues([[
          'Timestamp', 'Level', 'Message', 'Data', 'Error'
        ]]);
        sheet.getRange(1, 1, 1, 5)
          .setFontWeight('bold')
          .setBackground('#EA4335')
          .setFontColor('#FFFFFF');
      }

      // Limiter le nombre de lignes
      var lastRow = sheet.getLastRow();
      if (lastRow > config.maxSheetRows) {
        sheet.deleteRows(2, lastRow - config.maxSheetRows); // Garder les plus récents
      }

      // Ajouter la ligne
      var row = [
        formatTimestamp(),
        LEVEL_NAMES[level],
        truncate(message, 500),
        data ? truncate(safeStringify(data), 500) : '',
        error ? truncate(String(error.message || error), 500) : ''
      ];

      sheet.appendRow(row);

      // Formater la ligne selon le niveau
      var newRow = sheet.getLastRow();
      var color = level === LEVELS.ERROR ? '#FFEBEE' : '#FFF9C4';
      sheet.getRange(newRow, 1, 1, 5).setBackground(color);

    } catch (e) {
      // Fallback silencieux (ne pas créer de boucle infinie)
      console.error('Logger: Failed to log to sheet', e);
    }
  }

  // ===================================================================
  // API PRINCIPALE
  // ===================================================================

  /**
   * Fonction de log interne
   */
  function log(level, message, dataOrError, error) {
    // Ignorer si niveau trop bas
    if (level < config.currentLevel) return;

    // Séparer data et error
    var data = null;
    var err = null;

    if (dataOrError instanceof Error) {
      err = dataOrError;
    } else if (dataOrError && typeof dataOrError === 'object') {
      data = dataOrError;
    }

    if (error instanceof Error) {
      err = error;
    }

    // Formater le message
    var formattedMessage = formatMessage(level, message, data, err);

    // Envoyer vers les différents outputs
    logToConsole(level, formattedMessage);
    logToStackDriver(level, message, data, err);
    logToSheet(level, message, data, err);
  }

  // ===================================================================
  // API PUBLIQUE
  // ===================================================================

  return {
    // Constantes
    LEVELS: LEVELS,

    // Méthodes de logging
    trace: function(message, data, error) {
      log(LEVELS.TRACE, message, data, error);
    },

    debug: function(message, data, error) {
      log(LEVELS.DEBUG, message, data, error);
    },

    info: function(message, data, error) {
      log(LEVELS.INFO, message, data, error);
    },

    warn: function(message, data, error) {
      log(LEVELS.WARN, message, data, error);
    },

    error: function(message, data, error) {
      log(LEVELS.ERROR, message, data, error);
    },

    /**
     * Méthode log() pour compatibilité avec Logger natif Apps Script
     * Alias vers info() pour maintenir la compatibilité du code existant
     */
    log: function(message, data, error) {
      log(LEVELS.INFO, message, data, error);
    },

    // Configuration
    setLevel: function(level) {
      if (typeof level === 'number' && level >= 0 && level <= 4) {
        config.currentLevel = level;
        this.info('Logger level changed', { level: LEVEL_NAMES[level] });
      }
    },

    getLevel: function() {
      return config.currentLevel;
    },

    setConfig: function(newConfig) {
      for (var key in newConfig) {
        if (config.hasOwnProperty(key)) {
          config[key] = newConfig[key];
        }
      }
    },

    getConfig: function() {
      return JSON.parse(JSON.stringify(config)); // Deep copy
    },

    // Utilitaires
    enableProduction: function() {
      this.setLevel(LEVELS.INFO);
      config.enableSheetLogging = false;
      this.info('Logger configured for PRODUCTION');
    },

    enableDevelopment: function() {
      this.setLevel(LEVELS.DEBUG);
      config.enableSheetLogging = true;
      this.info('Logger configured for DEVELOPMENT');
    },

    // Performance helpers
    startTimer: function(label) {
      var start = Date.now();
      return {
        end: function(message) {
          var duration = Date.now() - start;
          Logger.info(message || label + ' completed', { duration: duration + 'ms' });
          return duration;
        }
      };
    },

    // Wrapper pour fonctions
    wrap: function(fn, context) {
      return function() {
        var label = fn.name || 'Anonymous';
        Logger.debug('Entering: ' + label);
        try {
          var result = fn.apply(context || this, arguments);
          Logger.debug('Exiting: ' + label);
          return result;
        } catch (error) {
          Logger.error('Exception in: ' + label, error);
          throw error;
        }
      };
    }
  };
})();

// ===================================================================
// INITIALISATION PAR DÉFAUT
// ===================================================================

// Configuration automatique selon environnement
try {
  // Détecter si on est en production (via propriétés du script)
  var scriptProperties = PropertiesService.getScriptProperties();
  var env = scriptProperties.getProperty('ENVIRONMENT') || 'development';

  if (env === 'production') {
    Logger.enableProduction();
  } else {
    Logger.enableProduction(); // Mode PRODUCTION par défaut (moins verbeux)
  }
} catch (e) {
  // Fallback: mode production par défaut (moins verbeux)
  Logger.enableProduction();
}

// ===================================================================
// COMPATIBILITÉ LEGACY
// ===================================================================

// logLine() → supprimée (définition canonique dans App.Core.js)
// Pour le logging avancé, utiliser Logger.info() directement.

// ===================================================================
// EXPORTS
// ===================================================================

// Global export (Apps Script)
// Sans 'use strict', this pointe vers le scope global dans Apps Script
this.Logger = Logger;
this.logLine = logLine;
