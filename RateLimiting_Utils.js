/**
 * ===================================================================
 * RATE LIMITING UTILITIES - Gestion des limites Google
 * ===================================================================
 * 
 * Fonctions pour gérer les erreurs HTTP 429 (Too Many Requests)
 * et éviter les dépassements de quota.
 */

/**
 * Exécute une fonction avec backoff exponentiel en cas d'erreur 429
 * 
 * @param {Function} fn - Fonction à exécuter
 * @param {Array} args - Arguments de la fonction
 * @param {string} label - Label pour le debug
 * @param {number} maxRetries - Nombre max de tentatives (défaut: 6)
 * @returns {*} Résultat de la fonction
 */
function backoff_(fn, args, label, maxRetries) {
  var tries = 0;
  var wait = 400; // ms initial
  maxRetries = maxRetries || 6; // ~0.4s, 0.8s, 1.6s, 3.2s, 6.4s, cap 8s
  
  while (true) {
    try {
      return fn.apply(null, args || []);
    } catch (e) {
      var msg = (e && e.message) || String(e);
      
      // Vérifier si c'est une erreur 429
      if (msg.indexOf('429') === -1 || tries >= maxRetries) {
        // Pas du rate-limit, ou trop de tentatives
        logLine('ERROR', '❌ Erreur ' + (label || 'backoff') + ': ' + msg);
        throw e;
      }
      
      // Log de la tentative
      tries++;
      var jitter = Math.floor(Math.random() * 250);
      var sleepTime = wait + jitter;
      
      logLine('WARN', '⚠️ Rate limit (429) sur ' + (label || 'fonction') + 
              ' - Tentative ' + tries + '/' + maxRetries + 
              ' - Attente ' + sleepTime + 'ms');
      
      Utilities.sleep(sleepTime);
      wait = Math.min(wait * 2, 8000); // Cap à 8s
    }
  }
}

/**
 * Throttle doux : évite les rafales en espaçant les appels
 * 
 * @param {string} key - Clé unique pour identifier l'opération
 * @param {number} minMs - Délai minimum entre deux appels (ms)
 */
function throttle_(key, minMs) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) {
      logLine('WARN', '⚠️ Throttle lock timeout pour ' + key);
      return;
    }
  } catch(e) {
    logLine('WARN', '⚠️ Throttle lock error: ' + e);
    return;
  }
  
  try {
    var props = PropertiesService.getScriptProperties();
    var last = Number(props.getProperty('throttle_' + key) || 0);
    var now = Date.now();
    var elapsed = now - last;
    
    if (elapsed < minMs) {
      var waitTime = minMs - elapsed;
      logLine('DEBUG', '⏱️ Throttle ' + key + ': attente ' + waitTime + 'ms');
      Utilities.sleep(waitTime);
    }
    
    props.setProperty('throttle_' + key, String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Micro-pause pour éviter les rafales
 * À utiliser dans les boucles de traitement
 * 
 * @param {number} count - Compteur d'itérations
 * @param {number} every - Pause tous les N éléments (défaut: 20)
 * @param {number} pauseMs - Durée de la pause (défaut: 120ms)
 */
function microPause_(count, every, pauseMs) {
  every = every || 20;
  pauseMs = pauseMs || 120;
  
  if (count > 0 && count % every === 0) {
    logLine('DEBUG', '⏸️ Micro-pause après ' + count + ' opérations (' + pauseMs + 'ms)');
    Utilities.sleep(pauseMs);
  }
}

/**
 * Wrapper sécurisé pour les écritures Sheets
 * Utilise backoff + throttle
 * 
 * @param {string} sheetName - Nom de la feuille
 * @param {Array} rows - Lignes à écrire
 */
function safeWriteToCache_(sheetName, rows) {
  // Throttle avant écriture
  throttle_('WRITE_' + sheetName, 150);
  
  // Écriture avec backoff
  return backoff_(
    writeBatchToCache_,
    [sheetName, rows],
    'writeBatchToCache[' + sheetName + ']',
    6
  );
}

/**
 * Wrapper sécurisé pour les lectures Sheets
 * Utilise backoff + throttle
 * 
 * @param {Array} levels - Niveaux à lire
 * @returns {Array} Élèves lus
 */
function safeReadFromCache_(levels) {
  // Throttle avant lecture
  throttle_('READ_CACHE', 100);
  
  // Lecture avec backoff
  return backoff_(
    readElevesFromCache_,
    [levels],
    'readElevesFromCache',
    6
  );
}

/**
 * Wrapper sécurisé pour le marquage dans _BASEOPTI
 * Utilise backoff + throttle
 * 
 * @param {Array} ids - IDs à marquer
 * @param {string} phase - Phase (P1/P2/P3)
 * @param {string} targetClass - Classe cible
 */
function safeMarkPlaced_(ids, phase, targetClass) {
  if (!ids || ids.length === 0) return;
  
  // Throttle avant marquage
  throttle_('MARK_PLACED', 150);
  
  // Marquage avec backoff
  return backoff_(
    baseMarkPlaced_,
    [ids, phase, targetClass],
    'baseMarkPlaced[' + phase + ']',
    6
  );
}

/**
 * Wrapper sécurisé pour la lecture de _BASEOPTI
 * Utilise backoff + throttle
 * 
 * @param {Object} filters - Filtres optionnels
 * @returns {Array} Élèves lus
 */
function safeReadBaseOpti_(filters) {
  // Throttle avant lecture
  throttle_('READ_BASEOPTI', 100);
  
  // Lecture avec backoff
  return backoff_(
    readBaseOpti_,
    [filters],
    'readBaseOpti',
    6
  );
}

/**
 * Wrapper sécurisé pour baseGetFree_
 * Utilise backoff + throttle
 * 
 * @returns {Array} Élèves libres
 */
function safeGetFree_() {
  // Throttle avant lecture
  throttle_('GET_FREE', 100);
  
  // Lecture avec backoff
  return backoff_(
    baseGetFree_,
    [],
    'baseGetFree',
    6
  );
}

/**
 * Batch writer : regroupe les écritures par classe
 * Évite les écritures multiples
 * 
 * @param {Object} batchByClass - { '6°1': [rows], '6°2': [rows], ... }
 */
function batchWriteToCache_(batchByClass) {
  var classes = Object.keys(batchByClass);
  var totalRows = 0;
  
  logLine('INFO', '📦 Batch write: ' + classes.length + ' classes');
  
  for (var i = 0; i < classes.length; i++) {
    var cls = classes[i];
    var rows = batchByClass[cls];
    
    if (rows && rows.length > 0) {
      totalRows += rows.length;
      
      // Micro-pause entre les classes
      if (i > 0) {
        microPause_(i, 1, 150);
      }
      
      // Écriture sécurisée
      safeWriteToCache_(cls + 'CACHE', rows);
    }
  }
  
  logLine('INFO', '  ✅ ' + totalRows + ' lignes écrites dans ' + classes.length + ' classes');
}
