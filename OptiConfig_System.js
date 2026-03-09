/**
 * ===================================================================
 * SYSTÈME _OPTI_CONFIG - Configuration UI Optimisation (V2)
 * ===================================================================
 *
 * Ce système fonctionne EN PARALLÈLE avec le système legacy (_STRUCTURE).
 *
 * DEUX SYSTÈMES COHABITENT :
 *
 * 1. SYSTÈME LEGACY :
 *    - Interface Google Sheets classique
 *    - Lit depuis _STRUCTURE
 *    - Fonctions : makeCtxFromUI_(), readQuotasFromUI_(), etc.
 *    - Phases anciennes
 *
 * 2. SYSTÈME NOUVEAU (ce fichier) :
 *    - UI Optimisation (SCORE INTERFACE)
 *    - Lit depuis _OPTI_CONFIG
 *    - Fonctions : getOptimizationContext_V2(), buildCtx_V2()
 *    - Phases BASEOPTI
 *
 * Architecture _OPTI_CONFIG :
 * | KEY                 | VALUE                              | SCOPE  | UPDATED_AT |
 * |---------------------|------------------------------------|--------|------------|
 * | mode.selected       | TEST                               | GLOBAL | 2025-01-20 |
 * | weights             | {"com":0.4,"tra":0.2,...}          | GLOBAL | 2025-01-20 |
 * | parity.tolerance    | 2                                  | GLOBAL | 2025-01-20 |
 * | swaps.max           | 50                                 | GLOBAL | 2025-01-20 |
 * | targets.override.6°1| 26                                 | GLOBAL | 2025-01-20 |
 */

// ===================================================================
// SECTION 1 : API KV (_OPTI_CONFIG)
// ===================================================================

/**
 * S'assure que la feuille _OPTI_CONFIG existe (cachée)
 * Crée les en-têtes si nécessaire
 */
function ensureConfigSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('_OPTI_CONFIG');

  if (!sh) {
    sh = ss.insertSheet('_OPTI_CONFIG');
    sh.hideSheet();

    // En-têtes
    const headers = ['KEY', 'VALUE', 'SCOPE', 'UPDATED_AT'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Formater l'en-tête
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#9C27B0')
      .setFontColor('#FFFFFF');

    logLine('INFO', '✅ _OPTI_CONFIG créé et caché');
  }

  return sh;
}

/**
 * Définit une clé/valeur dans _OPTI_CONFIG
 *
 * @param {string} key - Clé (ex: "mode.selected", "weights")
 * @param {string|number|object} value - Valeur (sera converti en string, JSON si objet)
 * @param {string} scope - Scope (GLOBAL | PROFILE:<nom> | CLASS:<code>)
 */
function kvSet_(key, value, scope) {
  scope = scope || 'GLOBAL';

  const sh = ensureConfigSheet_();
  const last = sh.getLastRow();
  const now = new Date();

  // Convertir valeur en string (JSON si objet)
  let valueStr = value;
  if (typeof value === 'object') {
    valueStr = JSON.stringify(value);
  } else {
    valueStr = String(value);
  }

  // Chercher si la clé existe déjà
  let row = -1;
  if (last > 1) {
    const data = sh.getRange(2, 1, last - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key && data[i][2] === scope) {
        row = i + 2;
        break;
      }
    }
  }

  // Si pas trouvé, ajouter à la fin
  if (row === -1) {
    row = last + 1;
  }

  // Écrire ou mettre à jour
  sh.getRange(row, 1, 1, 4).setValues([[key, valueStr, scope, now]]);

  logLine('INFO', '💾 _OPTI_CONFIG: ' + key + ' = ' + valueStr + ' (scope=' + scope + ')');
}

/**
 * Récupère une clé depuis _OPTI_CONFIG
 *
 * @param {string} key - Clé à récupérer
 * @param {string} scope - Scope (GLOBAL par défaut)
 * @param {*} defaultValue - Valeur par défaut si clé introuvable
 * @returns {string|null}
 */
function kvGet_(key, scope, defaultValue) {
  scope = scope || 'GLOBAL';
  defaultValue = defaultValue !== undefined ? defaultValue : null;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('_OPTI_CONFIG');

  if (!sh || sh.getLastRow() <= 1) {
    return defaultValue;
  }

  const last = sh.getLastRow();
  const data = sh.getRange(2, 1, last - 1, 3).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key && data[i][2] === scope) {
      return data[i][1];
    }
  }

  return defaultValue;
}

/**
 * Récupère toutes les clés/valeurs de _OPTI_CONFIG
 *
 * @returns {Object} Objet avec toutes les clés/valeurs
 */
function kvGetAll_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('_OPTI_CONFIG');

  const out = {};

  if (!sh || sh.getLastRow() <= 1) {
    return out;
  }

  const last = sh.getLastRow();
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();

  for (let i = 0; i < rows.length; i++) {
    const k = rows[i][0];
    const v = rows[i][1];
    const scope = rows[i][2];

    const target = (scope && scope !== 'GLOBAL') ? scope + ':' + k : k;
    out[target] = v;
  }

  return out;
}

// ===================================================================
// SECTION 2 : LECTURE CONTEXTE OPTIMISATION (V2)
// ===================================================================

/**
 * Construit le contexte d'optimisation depuis _OPTI_CONFIG
 *
 * HIÉRARCHIE DE LECTURE :
 * 1. _OPTI_CONFIG.targets.override.<classe> (tests utilisateur)
 * 2. _STRUCTURE.EFFECTIF (plan établissement)
 * 3. Fallback calculé : ceil(total / nbClasses)
 *
 * @returns {Object} Contexte avec mode, offersByClass, targetsByClass, weights, etc.
 */
function getOptimizationContext_V2() {
  logLine('INFO', '📋 Construction contexte depuis _OPTI_CONFIG (V2)...');

  // ===== 1. LIRE _OPTI_CONFIG =====
  const mode = kvGet_('mode.selected', 'GLOBAL', 'TEST');
  const weightsJson = kvGet_('weights', 'GLOBAL', null);
  const weights = weightsJson ? JSON.parse(weightsJson) : {
    parity: 0.3,
    com: 0.4,
    tra: 0.1,
    part: 0.1,
    abs: 0.1,
    effectif: 2.0,
    profiles: 2.0
  };
  const parityTolerance = Number(kvGet_('parity.tolerance', 'GLOBAL', 2));
  const maxSwaps = Number(kvGet_('swaps.max', 'GLOBAL', 50));
  const runtimeSec = Number(kvGet_('swaps.runtime', 'GLOBAL', 180)); // Budget temps P4 (défaut: 3 min)

  logLine('INFO', '  📌 Mode: ' + mode);
  logLine('INFO', '  📌 Poids: ' + JSON.stringify(weights));
  logLine('INFO', '  📌 Tolérance parité: ' + parityTolerance);
  logLine('INFO', '  📌 Max swaps: ' + maxSwaps);
  logLine('INFO', '  📌 Runtime P4: ' + runtimeSec + 's');

  // ===== 2. LIRE QUOTAS DEPUIS _OPTI_CONFIG (100% AUTONOME) =====
  // ✅ PIPELINE OPTI AUTONOME : Lit UNIQUEMENT depuis _OPTI_CONFIG
  // Les quotas sont sauvegardés via saveOptiConfigFromUI() depuis OptimizationPanel.html

  const offersJson = kvGet_('offers.byClass', 'GLOBAL', null);
  let offersByClass = offersJson ? JSON.parse(offersJson) : {};

  if (Object.keys(offersByClass).length > 0) {
    logLine('INFO', '  📊 Quotas lus depuis _OPTI_CONFIG: ' + Object.keys(offersByClass).length + ' classes');

    // Log détaillé des quotas
    for (const classe in offersByClass) {
      const quotas = offersByClass[classe];
      if (Object.keys(quotas).length > 0) {
        const quotaStr = Object.keys(quotas).map(function(opt) {
          return opt + '=' + quotas[opt];
        }).join(', ');
        logLine('INFO', '  📌 ' + classe + ' : ' + quotaStr);
      }
    }
  } else {
    logLine('INFO', '  ℹ️ Aucun quota dans _OPTI_CONFIG (configurez dans l\'interface Optimisation)');
    offersByClass = {};
  }

  // ===================================================================
  // ✅ DÉTECTION DYNAMIQUE DES CLASSES (PLUS DE CONSTANTE CODÉE EN DUR)
  // ===================================================================
  // Détecter les classes réellement présentes depuis les onglets CACHE/TEST/FINAL
  const detectedClasses = detectRealClasses_();
  
  // Combiner : classes détectées + classes de quotas
  const allClassesSet = new Set(detectedClasses);
  Object.keys(offersByClass).forEach(function(c) { allClassesSet.add(c); });
  const allClasses = Array.from(allClassesSet).sort();

  logLine('INFO', '  📋 Univers des classes: ' + allClasses.join(', '));
  logLine('INFO', '  📌 Source classes: Détection dynamique (onglets CACHE/TEST/FINAL)');
  logLine('INFO', '  📌 Source quotas: _OPTI_CONFIG (autonome)');

  // ✅ S'assurer que toutes les classes ont une entrée (vide si pas de quotas)
  allClasses.forEach(function(classe) {
    if (!offersByClass[classe]) {
      offersByClass[classe] = {};
    }
  });

  // ✅ CORRECTIF TARGETS DYNAMIQUES : Calculer depuis le nombre réel d'élèves
  const totalStudents = countStudentsFromBaseopti_();
  logLine('INFO', '  📊 Total élèves dans _BASEOPTI: ' + totalStudents);
  
  let targetsByClass = {};
  
  if (totalStudents > 0) {
    // Calcul équitable : répartir les élèves sur toutes les classes
    targetsByClass = computeTargetsFromUI_(allClasses, totalStudents);
    logLine('INFO', '  ✅ Targets calculées automatiquement (répartition équitable)');
    
    // Afficher les targets calculées
    allClasses.forEach(function(classe) {
      logLine('INFO', '  📊 ' + classe + ' effectif = ' + targetsByClass[classe] + ' (calculé depuis total=' + totalStudents + ')');
    });
  } else {
    // Fallback si _BASEOPTI est vide : utiliser _STRUCTURE ou 25
    allClasses.forEach(function(classe) {
      const override = kvGet_('targets.override.' + classe, 'GLOBAL', null);
      if (override) {
        targetsByClass[classe] = Number(override);
        logLine('INFO', '  🔧 Override: ' + classe + ' effectif = ' + override + ' (depuis _OPTI_CONFIG)');
      } else {
        targetsByClass[classe] = 25;
        logLine('INFO', '  ⚙️ ' + classe + ' effectif = 25 (fallback)');
      }
    });
  }

  logLine('INFO', '✅ Contexte V2 construit: ' + Object.keys(targetsByClass).length + ' classes');

  return {
    mode: mode,
    offersByClass: offersByClass,
    targetsByClass: targetsByClass,
    weights: weights,
    parityTolerance: parityTolerance,
    maxSwaps: maxSwaps,
    runtimeSec: runtimeSec
  };
}

/**
 * Détecte dynamiquement les classes réellement présentes dans le spreadsheet
 * 
 * Stratégie de détection (par ordre de priorité) :
 * 1. Onglets CACHE existants (ex: "6°1CACHE", "6°2CACHE")
 * 2. Onglets TEST existants (ex: "6°1TEST", "6°2TEST")
 * 3. Onglets FINAL existants (ex: "6°1FINAL", "6°2FINAL")
 * 4. Classes dans _STRUCTURE (si disponible)
 * 5. Classes dans _BASEOPTI (colonne CLASSE)
 * 
 * @returns {Array<string>} Liste des classes détectées (ex: ["6°1", "6°2", "5°1"])
 */
function detectRealClasses_() {
  logLine('INFO', '🔍 Détection dynamique des classes réellement présentes...');
  
  const ss = SpreadsheetApp.getActive();
  const allSheets = ss.getSheets();
  const classesSet = new Set();
  
  // ===================================================================
  // STRATÉGIE 1 : Détecter depuis les onglets CACHE
  // ===================================================================
  allSheets.forEach(function(sheet) {
    const name = sheet.getName();
    if (name.endsWith('CACHE')) {
      const className = name.replace('CACHE', '').trim();
      if (className) {
        classesSet.add(className);
        logLine('INFO', '  ✅ Classe détectée depuis CACHE: ' + className);
      }
    }
  });
  
  // ===================================================================
  // STRATÉGIE 2 : Détecter depuis les onglets TEST
  // ===================================================================
  if (classesSet.size === 0) {
    allSheets.forEach(function(sheet) {
      const name = sheet.getName();
      if (name.endsWith('TEST')) {
        const className = name.replace('TEST', '').trim();
        if (className) {
          classesSet.add(className);
          logLine('INFO', '  ✅ Classe détectée depuis TEST: ' + className);
        }
      }
    });
  }
  
  // ===================================================================
  // STRATÉGIE 3 : Détecter depuis les onglets FINAL
  // ===================================================================
  if (classesSet.size === 0) {
    allSheets.forEach(function(sheet) {
      const name = sheet.getName();
      if (name.endsWith('FINAL')) {
        const className = name.replace('FINAL', '').trim();
        if (className) {
          classesSet.add(className);
          logLine('INFO', '  ✅ Classe détectée depuis FINAL: ' + className);
        }
      }
    });
  }
  
  // ===================================================================
  // STRATÉGIE 4 : Détecter depuis _STRUCTURE
  // ===================================================================
  if (classesSet.size === 0) {
    const structSheet = ss.getSheetByName('_STRUCTURE');
    if (structSheet && structSheet.getLastRow() > 1) {
      const data = structSheet.getDataRange().getValues();
      const headers = data[0];
      const idxClasse = headers.indexOf('CLASSE');
      
      if (idxClasse >= 0) {
        for (let i = 1; i < data.length; i++) {
          const classe = String(data[i][idxClasse] || '').trim();
          if (classe && classe !== '') {
            classesSet.add(classe);
            logLine('INFO', '  ✅ Classe détectée depuis _STRUCTURE: ' + classe);
          }
        }
      }
    }
  }
  
  // ===================================================================
  // STRATÉGIE 5 : Détecter depuis _BASEOPTI
  // ===================================================================
  if (classesSet.size === 0) {
    const baseoptiSheet = ss.getSheetByName('_BASEOPTI');
    if (baseoptiSheet && baseoptiSheet.getLastRow() > 1) {
      const data = baseoptiSheet.getDataRange().getValues();
      const headers = data[0];
      const idxClasse = headers.indexOf('CLASSE');
      
      if (idxClasse >= 0) {
        for (let i = 1; i < data.length; i++) {
          const classe = String(data[i][idxClasse] || '').trim();
          if (classe && classe !== '') {
            classesSet.add(classe);
            logLine('INFO', '  ✅ Classe détectée depuis _BASEOPTI: ' + classe);
          }
        }
      }
    }
  }
  
  // ===================================================================
  // RÉSULTAT
  // ===================================================================
  const classes = Array.from(classesSet).sort();
  
  if (classes.length > 0) {
    logLine('INFO', '✅ ' + classes.length + ' classe(s) détectée(s): ' + classes.join(', '));
  } else {
    logLine('WARN', '⚠️ Aucune classe détectée ! Vérifiez que des onglets CACHE/TEST/FINAL existent.');
  }
  
  return classes;
}

/**
 * Compte le nombre total d'élèves dans _BASEOPTI
 * Utilisé pour calculer l'effectif moyen en fallback
 */
function countStudentsFromBaseopti_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('_BASEOPTI');

  if (!sh || sh.getLastRow() <= 1) {
    return 0;
  }

  // Compter uniquement les lignes avec ID en colonne A
  const data = sh.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const idColA = data[i][0];
    if (idColA && String(idColA).trim() !== '') {
      count++;
    }
  }

  return count;
}

/**
 * ✅ CORRECTIF TARGETS DYNAMIQUES
 * Calcule les effectifs cibles équitables depuis le nombre total d'élèves
 * @param {Array<string>} classes - Liste des classes (ex: ["6°1", "6°2", ...])
 * @param {number} totalStudents - Nombre total d'élèves
 * @returns {Object} - { "6°1": 24, "6°2": 25, ... }
 */
function computeTargetsFromUI_(classes, totalStudents) {
  const k = classes.length;
  if (k <= 0) throw new Error("Aucune classe UI.");
  if (totalStudents < 0) throw new Error("Total élèves négatif.");
  
  const avg = Math.floor(totalStudents / k);
  let rem = totalStudents - avg * k; // nombre de classes à avg+1

  const targets = {};
  for (const cls of classes) {
    targets[cls] = avg;
  }
  
  // Attribuer +1 aux 'rem' premières classes
  for (let i = 0; i < classes.length && rem > 0; i++, rem--) {
    targets[classes[i]] += 1;
  }
  
  return targets;
}

// ===================================================================
// SECTION 3 : SAUVEGARDE DEPUIS UI
// ===================================================================

/**
 * Sauvegarde le profil d'optimisation depuis l'UI
 * Appelé par SCORE INTERFACE quand l'utilisateur valide
 *
 * @param {Object} payload - { offersByClass, targetsByClass, weights, parityTolerance, maxSwaps, mode }
 * @returns {Object} { ok: true }
 */
function saveOptimizationProfileFromUI(payload) {
  logLine('INFO', '💾 Sauvegarde profil optimisation depuis UI...');

  try {
    // Mode sélectionné
    if (payload.mode) {
      kvSet_('mode.selected', payload.mode);
    }

    // Quotas par classe (offersByClass)
    if (payload.offersByClass) {
      kvSet_('offers.byClass', JSON.stringify(payload.offersByClass));
    }

    // Effectifs cibles par classe (targetsByClass)
    if (payload.targetsByClass) {
      // Option 1 : Stocker en JSON complet
      kvSet_('targets.byClass', JSON.stringify(payload.targetsByClass));

      // Option 2 : Stocker chaque classe séparément (pour overrides)
      for (const classe in payload.targetsByClass) {
        kvSet_('targets.override.' + classe, payload.targetsByClass[classe]);
      }
    }

    // Poids (weights)
    if (payload.weights) {
      kvSet_('weights', JSON.stringify(payload.weights));
    }

    // Tolérance parité
    if (payload.parityTolerance !== undefined) {
      kvSet_('parity.tolerance', String(payload.parityTolerance));
    }

    // Max swaps
    if (payload.maxSwaps !== undefined) {
      kvSet_('swaps.max', String(payload.maxSwaps));
    }

    SpreadsheetApp.flush();

    logLine('INFO', '✅ Profil optimisation sauvegardé dans _OPTI_CONFIG');

    return { ok: true };

  } catch (e) {
    logLine('ERROR', '❌ Erreur sauvegarde profil: ' + e);
    return { ok: false, error: e.toString() };
  }
}

/**
 * Récupère le profil d'optimisation depuis _OPTI_CONFIG
 * Utilisé par l'UI pour afficher les valeurs sauvegardées
 *
 * @returns {Object} Profil avec mode, offersByClass, targetsByClass, weights, etc.
 */
function getOptimizationProfileFromUI() {
  const ctx = getOptimizationContext_V2();

  return {
    mode: ctx.mode,
    offersByClass: ctx.offersByClass,
    targetsByClass: ctx.targetsByClass,
    weights: ctx.weights,
    parityTolerance: ctx.parityTolerance,
    maxSwaps: ctx.maxSwaps
  };
}

/**
 * Récupère la configuration d'optimisation pour l'interface UI
 * Appelé par OptimizationPanel.html au chargement de l'interface
 *
 * @returns {Object} { success: true, config: {...} } ou { success: false, error: "..." }
 */
function getOptiConfigForUI() {
  try {
    logLine('INFO', '📖 getOptiConfigForUI: Récupération configuration pour UI...');

    const ctx = getOptimizationContext_V2();

    const config = {
      mode: ctx.mode || 'TEST',
      weights: ctx.weights || {
        parity: 0.3,
        com: 0.4,
        tra: 0.1,
        part: 0.1,
        abs: 0.1,
        effectif: 2.0,
        profiles: 2.0
      },
      parityTolerance: ctx.parityTolerance || 2,
      maxSwaps: ctx.maxSwaps || 50,
      runtimeSec: ctx.runtimeSec || 180,
      offersByClass: ctx.offersByClass || {},
      targetsByClass: ctx.targetsByClass || {}
    };

    logLine('INFO', '✅ Configuration UI récupérée avec succès');

    return {
      success: true,
      config: config
    };

  } catch (e) {
    logLine('ERROR', '❌ Erreur lors de la récupération de la configuration UI: ' + e);
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Sauvegarde la configuration d'optimisation depuis l'interface UI
 * Appelé par OptimizationPanel.html lors de la sauvegarde
 *
 * @param {Object} config - Configuration complète depuis l'UI
 * @returns {Object} { success: true } ou { success: false, error: "..." }
 */
function saveOptiConfigFromUI(config) {
  try {
    logLine('INFO', '💾 saveOptiConfigFromUI: Sauvegarde configuration depuis UI...');

    // Préparer le payload pour saveOptimizationProfileFromUI
    const payload = {
      mode: config.mode || 'TEST',
      offersByClass: config.quotas || config.offersByClass || {},
      targetsByClass: config.targetsByClass || {},
      weights: config.weights || {
        parity: 0.3,
        com: 0.4,
        tra: 0.1,
        part: 0.1,
        abs: 0.1,
        effectif: 2.0,
        profiles: 2.0
      },
      parityTolerance: config.parityTolerance || 2,
      maxSwaps: config.maxSwaps || 50
    };

    // Utiliser la fonction existante
    const result = saveOptimizationProfileFromUI(payload);

    if (result.ok) {
      logLine('INFO', '✅ Configuration UI sauvegardée avec succès');
      return { success: true };
    } else {
      logLine('ERROR', '❌ Échec de la sauvegarde: ' + result.error);
      return { success: false, error: result.error };
    }

  } catch (e) {
    logLine('ERROR', '❌ Erreur lors de la sauvegarde de la configuration UI: ' + e);
    return {
      success: false,
      error: e.toString()
    };
  }
}

// ===================================================================
// SECTION 4 : CONSTRUCTION CONTEXTE POUR PHASES (V2)
// ===================================================================

/**
 * Lit les quotas depuis _STRUCTURE (source de vérité unique)
 * Retourne { "6°1": { "ITA": 6 }, "6°5": { "CHAV": 10 }, ... }
 * ✅ RENOMMÉ pour éviter conflit avec Orchestration_V14I.gs
 */
function readQuotasFromStructureV2_() {
  logLine('INFO', '📖 Lecture des quotas depuis _STRUCTURE (référence V2)...');
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('_STRUCTURE');
  
  if (!sheet) {
    logLine('WARN', '⚠️ Feuille _STRUCTURE introuvable, quotas vides');
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    logLine('WARN', '⚠️ _STRUCTURE vide, quotas vides');
    return {};
  }
  
  // ✅ CORRECTION : Recherche dynamique de l'en-tête (tolère lignes de garde/metadata)
  let headerRow = -1;
  for (let i = 0; i < Math.min(20, data.length); i++) {
    const row = data[i];
    // Chercher une ligne contenant CLASSE_DEST ou CLASSE_ORIGINE
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j] || '').trim().toUpperCase();
      if (cell === 'CLASSE_DEST' || cell === 'CLASSE_ORIGINE') {
        headerRow = i;
        break;
      }
    }
    if (headerRow !== -1) break;
  }
  
  if (headerRow === -1) {
    logLine('WARN', '⚠️ En-têtes non trouvés dans _STRUCTURE (cherché dans les 20 premières lignes)');
    logLine('WARN', '   Attendu: ligne contenant "CLASSE_DEST" ou "CLASSE_ORIGINE"');
    return {};
  }
  
  logLine('INFO', '✅ En-tête trouvé à la ligne ' + (headerRow + 1));
  
  const headers = data[headerRow];
  const colDest = headers.indexOf('CLASSE_DEST');
  const colOptions = headers.indexOf('OPTIONS');
  
  if (colDest === -1 || colOptions === -1) {
    logLine('WARN', '⚠️ Colonnes CLASSE_DEST ou OPTIONS non trouvées');
    return {};
  }
  
  const quotas = {};
  
  // Parcourir les lignes
  for (let i = headerRow + 1; i < data.length; i++) {
    const classeDest = String(data[i][colDest] || '').trim();
    const optionsStr = String(data[i][colOptions] || '').trim();
    
    if (!classeDest) continue;
    
    quotas[classeDest] = {};
    
    // Parser OPTIONS (format: "ITA=6,CHAV=10,LATIN=3")
    if (optionsStr) {
      const pairs = optionsStr.split(',');
      pairs.forEach(function(pair) {
        const parts = pair.trim().split('=');
        if (parts.length === 2) {
          const key = parts[0].trim().toUpperCase();
          const value = parseInt(parts[1].trim()) || 0;
          if (value > 0) {
            quotas[classeDest][key] = value;
          }
        }
      });
    }
  }
  
  logLine('INFO', '✅ Quotas lus depuis _STRUCTURE: ' + JSON.stringify(quotas));
  return quotas;
}

/**
 * Construit le contexte complet pour l'optimisation V2
 * ✅ ARCHITECTURE HYBRIDE :
 *    - Quotas depuis _STRUCTURE (source de vérité unique)
 *    - Poids/paramètres depuis _OPTI_CONFIG (historique)
 *
 * IMPORTANT : Cette fonction est utilisée par le SYSTÈME NOUVEAU uniquement
 * Le système legacy utilise buildCtx_() (Orchestration_V14I_Stream.gs)
 *
 * @param {Object} options - Options (sourceFamily, targetFamily)
 * @returns {Object} Contexte complet pour les phases
 */
function buildCtx_V2(options) {
  logLine('INFO', '🔧 buildCtx_V2: Construction contexte HYBRIDE (_STRUCTURE + _OPTI_CONFIG)...');

  const ss = SpreadsheetApp.getActive();

  // Lire le contexte depuis _OPTI_CONFIG (poids, paramètres)
  const optiCtx = getOptimizationContext_V2();

  // ✅ NOUVEAU : Lire les quotas depuis _STRUCTURE (référence)
  const quotasFromStructure = readQuotasFromStructureV2_();

  // Mode source/target
  const modeSrc = (options && options.sourceFamily) || optiCtx.mode || 'TEST';
  const writeTarget = (options && options.targetFamily) || 'CACHE';

  // Classes à traiter (depuis targetsByClass)
  const niveaux = Object.keys(optiCtx.targetsByClass);

  // Onglets sources et cibles
  const srcSheets = niveaux.map(function(c) { return c + modeSrc; });
  const cacheSheets = niveaux.map(function(c) { return c + 'CACHE'; });

  logLine('INFO', '  📋 Mode: ' + modeSrc + ' → ' + writeTarget);
  logLine('INFO', '  📋 Classes: ' + niveaux.join(', '));
  logLine('INFO', '  📋 Sources: ' + srcSheets.join(', '));
  logLine('INFO', '  📋 Cibles: ' + cacheSheets.join(', '));
  logLine('INFO', '  📌 Quotas (depuis _STRUCTURE): ' + JSON.stringify(quotasFromStructure));

  // Construire le contexte final
  return {
    ss: ss,
    modeSrc: modeSrc,
    writeTarget: writeTarget,
    niveaux: niveaux,
    levels: niveaux,
    srcSheets: srcSheets,
    cacheSheets: cacheSheets,
    quotas: quotasFromStructure,  // ✅ DEPUIS _STRUCTURE (référence)
    targets: optiCtx.targetsByClass,
    effectifCible: optiCtx.targetsByClass, // Compatibilité
    tolParite: optiCtx.parityTolerance,
    maxSwaps: optiCtx.maxSwaps,
    weights: optiCtx.weights
  };
}
