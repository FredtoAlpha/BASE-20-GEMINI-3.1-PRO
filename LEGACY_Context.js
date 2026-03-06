/**
 * ===================================================================
 * 🔧 PRIME LEGACY - GESTION DU CONTEXTE
 * ===================================================================
 *
 * Gestion du contexte LEGACY pour le pipeline :
 * - Détection automatique des onglets sources (°1, °2, etc.)
 * - Lecture de la configuration depuis _STRUCTURE
 * - Génération des onglets TEST de destination
 *
 * ISOLATION COMPLÈTE :
 * - OPTI : _BASEOPTI → _CACHE
 * - LEGACY : Sources (°1, °2) → TEST
 *
 * Date : 2025-11-13
 * Branche : claude/PRIME-LEGACY-01SJDcJv7zHGGBXWhHpzfnxr
 *
 * ===================================================================
 */

// ===================================================================
// CONSTRUCTION DU CONTEXTE LEGACY
// ===================================================================

/**
 * Détecte automatiquement les onglets sources existants et crée un contexte
 * pour le pipeline LEGACY initial (Sources → TEST)
 *
 * @returns {Object} Contexte prêt pour les 4 phases LEGACY
 */
function makeCtxFromSourceSheets_LEGACY() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  logLine('INFO', '🔍 Détection des onglets sources LEGACY...');

  // ========== ÉTAPE 1 : DÉTECTER ONGLETS SOURCES ==========
  const sourceSheets = [];

  // Pattern pour onglets sources : 6°1, ALBEXT°7, BONHOURE°2, etc. (toujours avec °)
  const sourcePattern = /^[A-Za-z0-9_-]+°\d+$/;
  // ❌ Exclure les onglets TEST, CACHE, DEF, FIN, etc.
  const excludePattern = /TEST|CACHE|DEF|FIN|SRC|SOURCE|_CONFIG|_STRUCTURE|_LOG/i;

  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName();
    if (sourcePattern.test(name) && !excludePattern.test(name)) {
      sourceSheets.push(name);
    }
  }

  if (sourceSheets.length === 0) {
    throw new Error(
      '❌ Aucun onglet source trouvé !\n\n' +
      'Formats supportés pour les onglets sources :\n' +
      '• Format classique: 6°1, 6°2, 5°1, 5°2, 4°1, 4°2, 3°1, 3°2, etc.\n' +
      '• Format ECOLE: ECOLE1, ECOLE2, ECOLE3, etc.\n' +
      '• Format personnalisé: GAMARRA°4, NOMECOLE°1, etc.\n\n' +
      'Note: Le symbole ° est obligatoire pour les formats personnalisés.'
    );
  }

  sourceSheets.sort();
  logLine('INFO', '📋 Onglets sources détectés : ' + sourceSheets.join(', '));

  // ========== ÉTAPE 2 : LIRE MAPPING DEPUIS _STRUCTURE ==========
  // Mapping CLASSE_ORIGINE → CLASSE_DEST (ex: "ECOLE1" → "6°1")
  const sourceToDestMapping = readSourceToDestMapping_LEGACY();

  logLine('INFO', '🗺️ Mapping sources → destinations :');
  for (const src in sourceToDestMapping) {
    logLine('INFO', '  • ' + src + ' → ' + sourceToDestMapping[src]);
  }

  // ========== ÉTAPE 3 : GÉNÉRER NOMS ONGLETS TEST ==========
  // Utiliser le mapping si disponible, sinon fallback intelligent
  const testSheets = sourceSheets.map(function(name) {
    // Si le mapping existe pour cette source, utiliser la destination mappée
    if (sourceToDestMapping[name]) {
      return sourceToDestMapping[name] + 'TEST';
    }

    // Sinon, fallback sur l'ancien comportement
    // Extraire le niveau (6°, 5°, etc.)
    const matchNiveau = name.match(/([3-6]°\d+)/);
    if (matchNiveau) {
      return matchNiveau[1] + 'TEST';
    }

    // Si c'est ECOLE, on génère 6°X TEST
    const matchEcole = name.match(/ECOLE(\d+)/);
    if (matchEcole) {
      return '6°' + matchEcole[1] + 'TEST';
    }

    // Fallback final
    return name + 'TEST';
  });

  logLine('INFO', '📋 Onglets TEST à créer : ' + testSheets.join(', '));

  // ========== ÉTAPE 4 : GÉNÉRER NIVEAUX DESTINATION ==========
  // Niveaux de destination (sans suffixe TEST)
  const niveauxDest = sourceSheets.map(function(name) {
    return sourceToDestMapping[name] || name;
  });

  logLine('INFO', '📋 Niveaux de destination : ' + niveauxDest.join(', '));

  // ========== ÉTAPE 5 : LIRE CONFIGURATION DEPUIS _STRUCTURE ==========
  // Lire les quotas par classe depuis _STRUCTURE
  const quotas = readQuotasFromUI_LEGACY();

  logLine('INFO', '📊 Quotas lus :');
  for (const classe in quotas) {
    const opts = quotas[classe];
    if (Object.keys(opts).length > 0) {
      logLine('INFO', '  • ' + classe + ' : ' + JSON.stringify(opts));
    }
  }

  // Lire les cibles d'effectifs par classe
  const targets = readTargetsFromUI_LEGACY();

  logLine('INFO', '🎯 Effectifs cibles :');
  for (const classe in targets) {
    logLine('INFO', '  • ' + classe + ' : ' + targets[classe] + ' élèves');
  }

  // Lire la tolérance de parité
  const tolParite = readParityToleranceFromUI_LEGACY() || 2;

  // Lire le nombre max de swaps
  const maxSwaps = readMaxSwapsFromUI_LEGACY() || 500;

  // Lire les autorisations de classes pour options/LV2
  const autorisations = readClassAuthorizationsFromUI_LEGACY();

  // ========== ÉTAPE 6 : CONSTRUIRE LE CONTEXTE ==========
  const ctx = {
    ss: ss,
    modeSrc: '',  // ✅ Mode vide pour LEGACY car les sources n'ont pas de suffixe
    writeTarget: 'TEST',  // ✅ Écrire vers onglets TEST
    niveaux: niveauxDest,  // ✅ Niveaux de destination (5°1, 5°2, etc.)
    levels: niveauxDest,  // ✅ ALIAS pour compatibilité Phase4_BASEOPTI_V2
    srcSheets: sourceSheets,  // ✅ Onglets sources réels (6°1, 6°2, etc.)
    cacheSheets: testSheets,  // ✅ Onglets TEST à créer (5°1TEST, 5°2TEST, etc.)
    sourceToDestMapping: sourceToDestMapping,  // ✅ Mapping pour utilisation dans les phases
    quotas: quotas,
    targets: targets,
    tolParite: tolParite,
    maxSwaps: maxSwaps,
    autorisations: autorisations,
    weights: {
      parity: 1.0,
      com: 1.0,
      tra: 0.5,
      part: 0.3,
      abs: 0.2
    }
  };

  logLine('INFO', '✅ Contexte LEGACY créé avec succès');
  logLine('INFO', '  • Sources : ' + ctx.srcSheets.length + ' onglets');
  logLine('INFO', '  • Destinations TEST : ' + ctx.cacheSheets.length + ' onglets');
  logLine('INFO', '  • Tolérance parité : ±' + ctx.tolParite);
  logLine('INFO', '  • Max swaps : ' + ctx.maxSwaps);

  return ctx;
}

// ===================================================================
// FONCTIONS DE LECTURE DE CONFIGURATION
// ===================================================================

/**
 * Lit le mapping CLASSE_ORIGINE → CLASSE_DEST depuis _STRUCTURE
 * @returns {Object} Mapping { "ECOLE1": "6°1", "6°1": "5°1", ... }
 */
function readSourceToDestMapping_LEGACY() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structSheet = ss.getSheetByName('_STRUCTURE');
  const mapping = {};

  if (!structSheet) {
    logLine('WARN', '⚠️ Onglet _STRUCTURE introuvable, mapping vide');
    return mapping;
  }

  try {
    const data = structSheet.getDataRange().getValues();

    // Recherche de l'en-tête (tolérer lignes de garde/metadata)
    let headerRow = -1;
    let isV3Format = false; // Format V3 (Type, Nom Classe) ou LEGACY (CLASSE_ORIGINE, CLASSE_DEST)

    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] || '').trim().toUpperCase();
        // ✅ Détecter format V3
        if (cell === 'TYPE' || cell === 'NOM CLASSE') {
          headerRow = i;
          isV3Format = true;
          break;
        }
        // Détecter format LEGACY
        if (cell === 'CLASSE_DEST' || cell === 'CLASSE_ORIGINE') {
          headerRow = i;
          isV3Format = false;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1) {
      logLine('WARN', '⚠️ En-têtes non trouvés dans _STRUCTURE (cherché dans les 20 premières lignes)');
      return mapping;
    }

    logLine('INFO', '  ✅ En-tête _STRUCTURE trouvé à la ligne ' + (headerRow + 1) + ' (format: ' + (isV3Format ? 'V3' : 'LEGACY') + ')');

    const headers = data[headerRow];

    if (isV3Format) {
      // ✅ FORMAT V3 : Type | Nom Classe | Capacité Max | Options (Quotas)
      let typeCol = -1;
      let nomCol = -1;

      // Trouver les colonnes avec normalisation
      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j] || '').trim().toUpperCase();
        if (h === 'TYPE') typeCol = j;
        if (h === 'NOM CLASSE' || h === 'NOM' || h === 'CLASSE') nomCol = j;
      }

      if (typeCol === -1 || nomCol === -1) {
        logLine('WARN', '⚠️ Colonnes Type ou Nom Classe introuvables dans format V3');
        return mapping;
      }

      // Lire les lignes SOURCE et TEST pour construire le mapping
      const sourceClasses = [];
      const testClasses = [];

      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        const type = String(row[typeCol] || '').trim().toUpperCase();
        const nom = String(row[nomCol] || '').trim();

        if (!nom) continue;

        if (type === 'SOURCE') {
          sourceClasses.push(nom);
        } else if (type === 'TEST') {
          testClasses.push(nom);
        }
      }

      // Mapper : 1er SOURCE → 1er TEST, 2ème SOURCE → 2ème TEST, etc.
      for (let i = 0; i < Math.max(sourceClasses.length, testClasses.length); i++) {
        if (sourceClasses[i] && testClasses[i]) {
          mapping[sourceClasses[i]] = testClasses[i];
        }
      }

    } else {
      // FORMAT LEGACY : CLASSE_ORIGINE | CLASSE_DEST | EFFECTIF | OPTIONS
      const origineCol = headers.indexOf('CLASSE_ORIGINE');
      const destCol = headers.indexOf('CLASSE_DEST');

      if (origineCol === -1 || destCol === -1) {
        logLine('WARN', '⚠️ Colonnes CLASSE_ORIGINE ou CLASSE_DEST introuvables');
        return mapping;
      }

      // Lire le mapping
      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        const origine = String(row[origineCol] || '').trim();
        const dest = String(row[destCol] || '').trim();

        if (origine && dest) {
          mapping[origine] = dest;
        }
      }
    }

  } catch (e) {
    logLine('WARN', '⚠️ Erreur lecture mapping depuis _STRUCTURE : ' + e.message);
  }

  return mapping;
}

/**
 * Lit les quotas par classe depuis l'interface
 * Format attendu : { "6°1": { ITA: 6, CHAV: 0, LV2_ESP: 3, ... }, ... }
 */
function readQuotasFromUI_LEGACY() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Essayer de lire depuis _STRUCTURE
  const structSheet = ss.getSheetByName('_STRUCTURE');
  if (structSheet) {
    return readQuotasFromStructure_LEGACY(structSheet);
  }

  // Sinon, retour valeurs par défaut vides
  logLine('WARN', '⚠️ _STRUCTURE introuvable, quotas par défaut (vides)');
  return {};
}

/**
 * Lit les quotas depuis la feuille _STRUCTURE
 * Parse la colonne OPTIONS au format "ITA=6,CHAV=10,ESP=5"
 */
function readQuotasFromStructure_LEGACY(sheet) {
  const quotas = {};

  try {
    const data = sheet.getDataRange().getValues();

    // ✅ Recherche dynamique de l'en-tête (tolère lignes de garde/metadata)
    let headerRow = -1;
    let isV3Format = false;

    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] || '').trim().toUpperCase();
        // Détecter format V3
        if (cell === 'TYPE' || cell === 'NOM CLASSE') {
          headerRow = i;
          isV3Format = true;
          break;
        }
        // Détecter format LEGACY
        if (cell === 'CLASSE_DEST' || cell === 'CLASSE_ORIGINE') {
          headerRow = i;
          isV3Format = false;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1) {
      logLine('WARN', '⚠️ En-têtes non trouvés dans _STRUCTURE');
      return quotas;
    }

    const headers = data[headerRow];

    if (isV3Format) {
      // ✅ FORMAT V3 : Type | Nom Classe | Capacité Max | Options (Quotas)
      let typeCol = -1;
      let nomCol = -1;
      let optionsCol = -1;

      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j] || '').trim().toUpperCase();
        if (h === 'TYPE') typeCol = j;
        if (h === 'NOM CLASSE' || h === 'NOM' || h === 'CLASSE') nomCol = j;
        if (h.indexOf('OPTIONS') !== -1 || h.indexOf('QUOTAS') !== -1) optionsCol = j;
      }

      if (nomCol === -1 || optionsCol === -1) {
        logLine('WARN', '⚠️ Colonnes Nom Classe ou Options introuvables dans format V3');
        return quotas;
      }

      // Lire les quotas depuis les lignes TEST (ce sont les classes de destination)
      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        const type = String(row[typeCol] || '').trim().toUpperCase();
        const nom = String(row[nomCol] || '').trim();
        const optionsStr = String(row[optionsCol] || '').trim();

        if (!nom) continue;

        // ✅ Lire les quotas pour les classes TEST (= destinations)
        if (type === 'TEST') {
          quotas[nom] = {};

          // Parser le format "ITA=6,CHAV=10,ESP=5"
          if (optionsStr) {
            optionsStr.split(',').forEach(function(pair) {
              const parts = pair.split('=');
              if (parts.length === 2) {
                const optName = parts[0].trim().toUpperCase();
                const optValue = parseInt(parts[1].trim()) || 0;
                quotas[nom][optName] = optValue;
              }
            });
          }
        }
      }

    } else {
      // FORMAT LEGACY : CLASSE_ORIGINE | CLASSE_DEST | EFFECTIF | OPTIONS
      const classeCol = headers.indexOf('CLASSE_DEST');
      const optionsCol = headers.indexOf('OPTIONS');

      if (classeCol === -1 || optionsCol === -1) {
        logLine('WARN', '⚠️ Colonnes CLASSE_DEST ou OPTIONS introuvables dans _STRUCTURE');
        return quotas;
      }

      // Parcourir les lignes (à partir de headerRow + 1)
      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        const classe = String(row[classeCol] || '').trim();
        if (!classe) continue;

        const optionsStr = String(row[optionsCol] || '').trim();

        quotas[classe] = {};

        // ✅ Parser le format "ITA=6,CHAV=10,ESP=5"
        if (optionsStr) {
          optionsStr.split(',').forEach(function(pair) {
            const parts = pair.split('=');
            if (parts.length === 2) {
              const optName = parts[0].trim().toUpperCase();
              const optValue = parseInt(parts[1].trim()) || 0;
              quotas[classe][optName] = optValue;
            }
          });
        }
      }
    }

  } catch (e) {
    logLine('WARN', '⚠️ Erreur lecture quotas depuis _STRUCTURE : ' + e.message);
  }

  return quotas;
}

/**
 * Lit les cibles d'effectifs par classe depuis l'interface
 * ✅ Lit depuis _STRUCTURE si disponible
 */
function readTargetsFromUI_LEGACY() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Essayer de lire depuis _STRUCTURE
  const structSheet = ss.getSheetByName('_STRUCTURE');
  if (structSheet) {
    return readTargetsFromStructure_LEGACY(structSheet);
  }

  // Sinon, valeurs par défaut : 25 élèves par classe
  logLine('WARN', '⚠️ _STRUCTURE introuvable, effectifs par défaut (25)');
  return {};
}

/**
 * Lit les effectifs cibles depuis _STRUCTURE
 * Lit la colonne EFFECTIF pour chaque classe
 */
function readTargetsFromStructure_LEGACY(sheet) {
  const targets = {};

  try {
    const data = sheet.getDataRange().getValues();

    // Recherche dynamique de l'en-tête
    let headerRow = -1;
    let isV3Format = false;

    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] || '').trim().toUpperCase();
        // Détecter format V3
        if (cell === 'TYPE' || cell === 'NOM CLASSE') {
          headerRow = i;
          isV3Format = true;
          break;
        }
        // Détecter format LEGACY
        if (cell === 'CLASSE_DEST' || cell === 'CLASSE_ORIGINE') {
          headerRow = i;
          isV3Format = false;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1) {
      logLine('WARN', '⚠️ En-têtes non trouvés dans _STRUCTURE');
      return targets;
    }

    const headers = data[headerRow];

    if (isV3Format) {
      // ✅ FORMAT V3 : Type | Nom Classe | Capacité Max | Options (Quotas)
      let typeCol = -1;
      let nomCol = -1;
      let capaciteCol = -1;

      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j] || '').trim().toUpperCase();
        if (h === 'TYPE') typeCol = j;
        if (h === 'NOM CLASSE' || h === 'NOM' || h === 'CLASSE') nomCol = j;
        if (h.indexOf('CAPACITÉ') !== -1 || h.indexOf('CAPACITE') !== -1 || h.indexOf('MAX') !== -1) capaciteCol = j;
      }

      if (nomCol === -1 || capaciteCol === -1) {
        logLine('WARN', '⚠️ Colonnes Nom Classe ou Capacité introuvables dans format V3');
        return targets;
      }

      // Lire les effectifs depuis les lignes TEST (ce sont les classes de destination)
      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        const type = String(row[typeCol] || '').trim().toUpperCase();
        const nom = String(row[nomCol] || '').trim();
        const capacite = parseInt(row[capaciteCol]) || 25; // Fallback 25

        if (!nom) continue;

        // ✅ Lire les effectifs pour les classes TEST (= destinations)
        if (type === 'TEST') {
          targets[nom] = capacite;
        }
      }

    } else {
      // FORMAT LEGACY : CLASSE_ORIGINE | CLASSE_DEST | EFFECTIF | OPTIONS
      const classeCol = headers.indexOf('CLASSE_DEST');
      const effectifCol = headers.indexOf('EFFECTIF');

      if (classeCol === -1 || effectifCol === -1) {
        logLine('WARN', '⚠️ Colonnes CLASSE_DEST ou EFFECTIF introuvables dans _STRUCTURE');
        return targets;
      }

      // Parcourir les lignes
      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        const classe = String(row[classeCol] || '').trim();
        if (!classe) continue;

        const effectif = parseInt(row[effectifCol]) || 25; // Fallback 25
        targets[classe] = effectif;
      }
    }

  } catch (e) {
    logLine('WARN', '⚠️ Erreur lecture effectifs depuis _STRUCTURE : ' + e.message);
  }

  return targets;
}

/**
 * Lit la tolérance de parité depuis l'interface
 * Retourne une valeur par défaut (2)
 */
function readParityToleranceFromUI_LEGACY() {
  // ✅ Valeur par défaut : tolérance de ±2
  return 2;
}

/**
 * Lit le nombre max de swaps depuis l'interface
 * Retourne une valeur par défaut (500)
 */
function readMaxSwapsFromUI_LEGACY() {
  // ✅ Valeur par défaut : 500 swaps max
  return 500;
}

/**
 * Lit les autorisations de classes par option
 * Format : { ITA: ["6°1", "6°3"], CHAV: ["6°2", "6°3"], ... }
 *
 * ✅ Calculées automatiquement depuis les quotas
 */
function readClassAuthorizationsFromUI_LEGACY() {
  const autorisations = {};

  // ✅ Calculer depuis les quotas
  const quotas = readQuotasFromUI_LEGACY();

  for (const classe in quotas) {
    const opts = quotas[classe];
    for (const optName in opts) {
      if (opts[optName] > 0) {
        if (!autorisations[optName]) {
          autorisations[optName] = [];
        }
        autorisations[optName].push(classe);
      }
    }
  }

  return autorisations;
}

// ===================================================================
// UTILITAIRES CONTEXTE
// ===================================================================

/**
 * Affiche le contexte LEGACY dans les logs
 * @param {Object} ctx - Contexte LEGACY
 */
function logContext_LEGACY(ctx) {
  logLine('INFO', '');
  logLine('INFO', '📋 CONTEXTE LEGACY :');
  logLine('INFO', '─────────────────────────────────────────────────────');
  logLine('INFO', '  • Sources : ' + (ctx.srcSheets || []).join(', '));
  logLine('INFO', '  • Destinations TEST : ' + (ctx.cacheSheets || []).join(', '));
  logLine('INFO', '  • Niveaux : ' + (ctx.niveaux || []).join(', '));
  logLine('INFO', '  • Tolérance parité : ±' + (ctx.tolParite || 2));
  logLine('INFO', '  • Max swaps : ' + (ctx.maxSwaps || 500));
  logLine('INFO', '  • Quotas : ' + Object.keys(ctx.quotas || {}).length + ' classes');
  logLine('INFO', '  • Effectifs cibles : ' + Object.keys(ctx.targets || {}).length + ' classes');
  logLine('INFO', '─────────────────────────────────────────────────────');
  logLine('INFO', '');
}
