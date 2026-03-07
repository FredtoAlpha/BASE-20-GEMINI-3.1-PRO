/**
 * ===================================================================
 * GESTION DYNAMIQUE DES NIVEAUX - Système Universel
 * ===================================================================
 *
 * Ce module permet de lire dynamiquement les niveaux depuis la configuration
 * au lieu d'utiliser des valeurs hardcodées.
 *
 * Supporte : 6°, 5°, 4°, 3° et tous les formats de _STRUCTURE
 */

/**
 * Lit le niveau scolaire depuis _CONFIG
 * @return {string} Le niveau (ex: "6°", "5°", "4°")
 */
function lireNiveauDepuisConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('_CONFIG');

    if (!configSheet) {
      logLine('WARN', '⚠️ Onglet _CONFIG introuvable, utilisation niveau par défaut: 6°');
      return '6°';
    }

    const data = configSheet.getDataRange().getValues();

    // Chercher la ligne NIVEAU
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'NIVEAU') {
        const niveau = String(data[i][1]).trim();
        if (niveau) {
          logLine('INFO', `✅ Niveau lu depuis _CONFIG: ${niveau}`);
          return niveau;
        }
      }
    }

    logLine('WARN', '⚠️ Paramètre NIVEAU introuvable dans _CONFIG, utilisation: 6°');
    return '6°';
  } catch (e) {
    logLine('ERROR', `❌ Erreur lecture niveau: ${e}`);
    return '6°';
  }
}

/**
 * Lit les noms des classes depuis _STRUCTURE (format LEGACY)
 * Format attendu : CLASSE_ORIGINE | CLASSE_DEST | EFFECTIF | OPTIONS
 * @return {string[]} Liste des noms de classes (ex: ['6°1', '6°2', '6°3'])
 */
function lireClassesDepuisStructureLegacy() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const structureSheet = ss.getSheetByName('_STRUCTURE');

    if (!structureSheet) {
      logLine('WARN', '⚠️ Onglet _STRUCTURE introuvable');
      return null;
    }

    const data = structureSheet.getDataRange().getValues();
    if (data.length <= 1) {
      logLine('WARN', '⚠️ _STRUCTURE vide');
      return null;
    }

    const headers = data[0];

    // Vérifier si c'est le format LEGACY (CLASSE_ORIGINE/CLASSE_DEST)
    const origineIdx = headers.indexOf('CLASSE_ORIGINE');
    const destIdx = headers.indexOf('CLASSE_DEST');

    if (origineIdx !== -1 && destIdx !== -1) {
      // Format LEGACY détecté
      const classesSet = new Set();

      for (let i = 1; i < data.length; i++) {
        const origine = String(data[i][origineIdx] || '').trim();
        const dest = String(data[i][destIdx] || '').trim();

        if (origine) classesSet.add(origine);
        if (dest) classesSet.add(dest);
      }

      const classes = Array.from(classesSet).sort();
      logLine('INFO', `✅ Classes lues depuis _STRUCTURE (format LEGACY): ${classes.join(', ')}`);
      return classes;
    }

    return null;
  } catch (e) {
    logLine('ERROR', `❌ Erreur lecture _STRUCTURE (LEGACY): ${e}`);
    return null;
  }
}

/**
 * Lit les noms des classes depuis _STRUCTURE (format INITIALISATION)
 * Format attendu : Type | Nom Classe | Capacité Max | Prof Principal | Observations
 * @return {string[]} Liste des noms de classes sources (ex: ['6°1', '6°2', '6°3'])
 */
function lireClassesDepuisStructureInit() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const structureSheet = ss.getSheetByName('_STRUCTURE');

    if (!structureSheet) {
      return null;
    }

    const data = structureSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return null;
    }

    const headers = data[0];

    // Vérifier si c'est le format INITIALISATION (Type/Nom Classe)
    const typeIdx = headers.indexOf('Type');
    const nomIdx = headers.indexOf('Nom Classe');

    if (typeIdx !== -1 && nomIdx !== -1) {
      // Format INITIALISATION détecté
      const classes = [];

      for (let i = 1; i < data.length; i++) {
        const type = String(data[i][typeIdx] || '').trim().toUpperCase();
        const nom = String(data[i][nomIdx] || '').trim();

        // Récupérer uniquement les classes SOURCE
        if (type === 'SOURCE' && nom) {
          classes.push(nom);
        }
      }

      classes.sort();
      logLine('INFO', `✅ Classes lues depuis _STRUCTURE (format INIT): ${classes.join(', ')}`);
      return classes;
    }

    return null;
  } catch (e) {
    logLine('ERROR', `❌ Erreur lecture _STRUCTURE (INIT): ${e}`);
    return null;
  }
}

/**
 * Génère dynamiquement la liste des niveaux à traiter
 * Remplace readNiveauxFromUI_() qui retourne des valeurs hardcodées
 *
 * Stratégie :
 * 1. Tente de lire depuis _STRUCTURE (format LEGACY)
 * 2. Sinon, tente format INITIALISATION
 * 3. Sinon, génère depuis niveau + nombre de classes
 *
 * @return {string[]} Liste des niveaux (ex: ['6°1', '6°2', '6°3', '6°4', '6°5'])
 */
function genererNiveauxDynamiques() {
  logLine('INFO', '📋 Génération dynamique des niveaux...');

  // Stratégie 1 : Lire depuis _STRUCTURE (format LEGACY)
  const classesLegacy = lireClassesDepuisStructureLegacy();
  if (classesLegacy && classesLegacy.length > 0) {
    return classesLegacy;
  }

  // Stratégie 2 : Lire depuis _STRUCTURE (format INITIALISATION)
  const classesInit = lireClassesDepuisStructureInit();
  if (classesInit && classesInit.length > 0) {
    return classesInit;
  }

  // Stratégie 3 : Générer depuis niveau + nombre par défaut
  const niveau = lireNiveauDepuisConfig();
  const nbClasses = 5; // Valeur par défaut

  const classes = [];
  for (let i = 1; i <= nbClasses; i++) {
    classes.push(`${niveau}${i}`);
  }

  logLine('WARN', `⚠️ Niveaux générés par défaut: ${classes.join(', ')}`);
  return classes;
}

/**
 * Version améliorée de readNiveauxFromUI_() qui lit dynamiquement
 * @deprecated Utiliser genererNiveauxDynamiques() à la place
 */
function readNiveauxFromUI_V2() {
  return genererNiveauxDynamiques();
}

// logLine : définition canonique dans App.Core.js (pas de fallback local)
