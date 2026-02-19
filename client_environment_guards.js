/**
 * CLIENT ENVIRONMENT GUARDS
 * Protection contre l'exécution de code DOM côté serveur
 *
 * À inclure au début de tout fichier HTML qui utilise le DOM
 */

(function(global) {
  'use strict';

  /**
   * Détecte si on est côté client (navigateur)
   */
  const isClient = typeof document !== 'undefined' && typeof window !== 'undefined';

  /**
   * Détecte si on est côté serveur (Apps Script)
   */
  const isServer = !isClient;

  /**
   * Namespace global pour les guards
   */
  global.ClientGuards = {
    /**
     * Indique si on est côté client
     */
    isClient: isClient,

    /**
     * Indique si on est côté serveur
     */
    isServer: isServer,

    /**
     * Exécute une fonction uniquement côté client
     * @param {Function} fn - Fonction à exécuter
     * @param {any} context - Contexte d'exécution (this)
     * @returns {any} Résultat de la fonction ou undefined si serveur
     */
    runOnClient: function(fn, context) {
      if (!isClient) {
        console.warn('[ClientGuards] Skipping function execution in server context');
        return undefined;
      }
      return fn.call(context || global);
    },

    /**
     * Exécute une fonction uniquement côté serveur
     * @param {Function} fn - Fonction à exécuter
     * @param {any} context - Contexte d'exécution (this)
     * @returns {any} Résultat de la fonction ou undefined si client
     */
    runOnServer: function(fn, context) {
      if (!isServer) {
        console.warn('[ClientGuards] Skipping function execution in client context');
        return undefined;
      }
      return fn.call(context || global);
    },

    /**
     * Safe getElementById - retourne null si serveur
     */
    safeGetElementById: function(id) {
      return isClient ? document.getElementById(id) : null;
    },

    /**
     * Safe querySelector - retourne null si serveur
     */
    safeQuerySelector: function(selector) {
      return isClient ? document.querySelector(selector) : null;
    },

    /**
     * Safe querySelectorAll - retourne tableau vide si serveur
     */
    safeQuerySelectorAll: function(selector) {
      return isClient ? Array.from(document.querySelectorAll(selector)) : [];
    },

    /**
     * Safe addEventListener - ne fait rien si serveur
     */
    safeAddEventListener: function(element, event, handler, options) {
      if (isClient && element && element.addEventListener) {
        element.addEventListener(event, handler, options);
      }
    },

    /**
     * Safe window property access
     */
    safeWindowProperty: function(propertyName, defaultValue) {
      return isClient && global[propertyName] !== undefined
        ? global[propertyName]
        : defaultValue;
    }
  };

  // Log du statut au chargement
  if (isClient) {
    console.log('✅ Client Guards: Running in client context (browser)');
  } else {
    console.warn('⚠️ Client Guards: Running in server context, guards active');
  }

})(typeof window !== 'undefined' ? window : this);
