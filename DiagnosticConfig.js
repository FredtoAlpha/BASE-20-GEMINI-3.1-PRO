/**
 * =====================================================================
 * FONCTION DE DIAGNOSTIC DÉDIÉE - LECTURE _CONFIG
 * =====================================================================
 * Cette fonction va tracer EXPLICITEMENT chaque étape de la lecture
 * de _CONFIG et nous dire EXACTEMENT où est le problème.
 */

function DIAGNOSTIC_CONFIG() {
  const rapport = {
    etapes: [],
    erreurs: [],
    succes: false,
    config: {}
  };
  
  try {
    rapport.etapes.push("📌 ÉTAPE 1: Début du diagnostic");
    
    // ÉTAPE 2: Récupérer le spreadsheet
    rapport.etapes.push("📌 ÉTAPE 2: Récupération du spreadsheet");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      rapport.erreurs.push("❌ ERREUR ÉTAPE 2: Impossible de récupérer le spreadsheet");
      return rapport;
    }
    rapport.etapes.push("✅ ÉTAPE 2: Spreadsheet récupéré avec succès");
    
    // ÉTAPE 3: Récupérer la feuille _CONFIG
    rapport.etapes.push("📌 ÉTAPE 3: Recherche de la feuille _CONFIG");
    const configSheet = ss.getSheetByName("_CONFIG");
    if (!configSheet) {
      rapport.erreurs.push("❌ ERREUR ÉTAPE 3: Feuille _CONFIG introuvable");
      return rapport;
    }
    rapport.etapes.push("✅ ÉTAPE 3: Feuille _CONFIG trouvée");
    
    // ÉTAPE 4: Lire toutes les données
    rapport.etapes.push("📌 ÉTAPE 4: Lecture des données de _CONFIG");
    const data = configSheet.getDataRange().getValues();
    rapport.etapes.push(`✅ ÉTAPE 4: ${data.length} lignes lues`);
    
    // ÉTAPE 5: Afficher les 10 premières lignes brutes
    rapport.etapes.push("📌 ÉTAPE 5: Affichage des données brutes");
    for (let i = 0; i < Math.min(10, data.length); i++) {
      rapport.etapes.push(`   Ligne ${i+1}: [${data[i].map(v => `"${v}"`).join(', ')}]`);
    }
    
    // ÉTAPE 6: Rechercher les en-têtes
    rapport.etapes.push("📌 ÉTAPE 6: Recherche des en-têtes PARAMETRE/VALEUR");
    let headerRowIndex = -1;
    let headers = [];
    
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const potentialHeaders = data[i].map(h => String(h || '').trim().toUpperCase());
      rapport.etapes.push(`   Ligne ${i+1} en majuscules: [${potentialHeaders.join(', ')}]`);
      
      const hasParam = potentialHeaders.includes("PARAMETRE") || potentialHeaders.includes("PARAMÈTRE");
      const hasValeur = potentialHeaders.includes("VALEUR");
      
      if (hasParam && hasValeur) {
        headers = potentialHeaders;
        headerRowIndex = i;
        rapport.etapes.push(`✅ ÉTAPE 6: En-têtes trouvés à la ligne ${i+1}`);
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      rapport.erreurs.push("❌ ERREUR ÉTAPE 6: En-têtes PARAMETRE/VALEUR non trouvés");
      rapport.erreurs.push("   Les en-têtes doivent être 'PARAMETRE' et 'VALEUR' (ou 'Paramètre' et 'Valeur')");
      return rapport;
    }
    
    // ÉTAPE 7: Identifier les colonnes
    rapport.etapes.push("📌 ÉTAPE 7: Identification des index de colonnes");
    const paramIndex = headers.indexOf("PARAMETRE") !== -1 ? headers.indexOf("PARAMETRE") : headers.indexOf("PARAMÈTRE");
    const valueIndex = headers.indexOf("VALEUR");
    rapport.etapes.push(`✅ ÉTAPE 7: Colonne PARAMETRE à l'index ${paramIndex}, VALEUR à l'index ${valueIndex}`);
    
    // ÉTAPE 8: Lire les paramètres ligne par ligne
    rapport.etapes.push("📌 ÉTAPE 8: Lecture des paramètres");
    const parametres = {};
    
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const paramKeyRaw = String(data[i][paramIndex] || '').trim();
      const paramKeyUpper = paramKeyRaw.toUpperCase();
      let value = data[i][valueIndex];
      
      if (!paramKeyRaw) {
        rapport.etapes.push(`   Ligne ${i+1}: VIDE - ignorée`);
        continue;
      }
      
      // Conversion des types
      if (typeof value === 'string') {
        const valueLower = value.toLowerCase().trim();
        if (valueLower === 'true') value = true;
        else if (valueLower === 'false') value = false;
        else if (value.trim() !== '' && !isNaN(Number(value))) value = Number(value);
      }
      
      parametres[paramKeyRaw] = value;
      rapport.etapes.push(`   Ligne ${i+1}: "${paramKeyRaw}" = ${JSON.stringify(value)} (type: ${typeof value})`);
      
      // Tracer spécifiquement NB_SOURCES et NB_DEST
      if (paramKeyUpper === "NB_SOURCES" || paramKeyUpper === "NB_DEST") {
        rapport.etapes.push(`   🎯 TROUVÉ: ${paramKeyRaw} = ${value}`);
      }
    }
    
    // ÉTAPE 9: Vérifier les paramètres critiques
    rapport.etapes.push("📌 ÉTAPE 9: Vérification des paramètres critiques");
    const critiques = ["NB_SOURCES", "NB_DEST", "NIVEAU", "ADMIN_PASSWORD"];
    
    critiques.forEach(key => {
      const found = parametres[key];
      if (found !== undefined) {
        rapport.etapes.push(`✅ ${key} trouvé: ${JSON.stringify(found)}`);
        rapport.config[key] = found;
      } else {
        rapport.erreurs.push(`⚠️ ${key} NON trouvé dans _CONFIG`);
      }
    });
    
    // ÉTAPE 10: Résultat final
    rapport.etapes.push("📌 ÉTAPE 10: Diagnostic terminé");
    rapport.succes = true;
    rapport.config = parametres;
    
  } catch (e) {
    rapport.erreurs.push(`💥 EXCEPTION: ${e.message}`);
    rapport.erreurs.push(`   Stack: ${e.stack}`);
  }
  
  // AFFICHER LE RAPPORT COMPLET
  Logger.log("\n╔═══════════════════════════════════════════════════════════╗");
  Logger.log("║         RAPPORT DE DIAGNOSTIC _CONFIG                    ║");
  Logger.log("╚═══════════════════════════════════════════════════════════╝\n");
  
  rapport.etapes.forEach(etape => Logger.log(etape));
  
  if (rapport.erreurs.length > 0) {
    Logger.log("\n❌ ERREURS DÉTECTÉES:");
    rapport.erreurs.forEach(err => Logger.log(err));
  }
  
  Logger.log("\n📊 RÉSUMÉ:");
  Logger.log(`   Succès: ${rapport.succes ? "OUI ✅" : "NON ❌"}`);
  Logger.log(`   Paramètres trouvés: ${Object.keys(rapport.config).length}`);
  
  if (rapport.config.NB_SOURCES !== undefined) {
    Logger.log(`   🎯 NB_SOURCES = ${rapport.config.NB_SOURCES}`);
  }
  if (rapport.config.NB_DEST !== undefined) {
    Logger.log(`   🎯 NB_DEST = ${rapport.config.NB_DEST}`);
  }
  
  Logger.log("\n╚═══════════════════════════════════════════════════════════╝\n");
  
  return rapport;
}

/**
 * Fonction de test rapide à appeler depuis Apps Script
 */
function TEST_DIAGNOSTIC() {
  const rapport = DIAGNOSTIC_CONFIG();
  
  if (rapport.succes) {
    Logger.log("✅ DIAGNOSTIC RÉUSSI");
    Logger.log(`NB_SOURCES trouvé: ${rapport.config.NB_SOURCES !== undefined ? "OUI" : "NON"}`);
    Logger.log(`NB_DEST trouvé: ${rapport.config.NB_DEST !== undefined ? "OUI" : "NON"}`);
  } else {
    Logger.log("❌ DIAGNOSTIC ÉCHOUÉ - Voir les erreurs ci-dessus");
  }
}

/**
 * TEST COMPLET: Compare ce que lit DIAGNOSTIC_CONFIG vs ce que retourne getConfig()
 */
function TEST_COMPLET_GETCONFIG() {
  Logger.log("\n╔═══════════════════════════════════════════════════════════╗");
  Logger.log("║    TEST COMPLET: DIAGNOSTIC vs getConfig()               ║");
  Logger.log("╚═══════════════════════════════════════════════════════════╝\n");
  
  // PARTIE 1: Ce que lit directement DIAGNOSTIC_CONFIG
  Logger.log("📋 PARTIE 1: Lecture directe de _CONFIG par DIAGNOSTIC_CONFIG");
  const rapport = DIAGNOSTIC_CONFIG();
  
  Logger.log("\n📊 Ce que DIAGNOSTIC_CONFIG a trouvé:");
  Logger.log(`   NB_SOURCES = ${rapport.config.NB_SOURCES}`);
  Logger.log(`   NB_DEST = ${rapport.config.NB_DEST}`);
  Logger.log(`   NIVEAU = "${rapport.config.NIVEAU}"`);
  Logger.log(`   ADMIN_PASSWORD = "${rapport.config.ADMIN_PASSWORD}"`);
  
  // PARTIE 2: Ce que retourne getConfig()
  Logger.log("\n📋 PARTIE 2: Ce que retourne getConfig()");
  const config = getConfig();
  
  Logger.log("\n📊 Ce que getConfig() retourne:");
  Logger.log(`   config.NB_SOURCES = ${config.NB_SOURCES}`);
  Logger.log(`   config.NB_DEST = ${config.NB_DEST}`);
  Logger.log(`   config.NIVEAU = "${config.NIVEAU}"`);
  Logger.log(`   config.ADMIN_PASSWORD = "${config.ADMIN_PASSWORD}"`);
  
  // PARTIE 3: Comparaison
  Logger.log("\n╔═══════════════════════════════════════════════════════════╗");
  Logger.log("║              COMPARAISON DES RÉSULTATS                    ║");
  Logger.log("╚═══════════════════════════════════════════════════════════╝\n");
  
  const compareParam = (key) => {
    const diagnosticValue = rapport.config[key];
    const getConfigValue = config[key];
    
    if (diagnosticValue === getConfigValue) {
      Logger.log(`✅ ${key}: IDENTIQUE (${JSON.stringify(diagnosticValue)})`);
    } else {
      Logger.log(`❌ ${key}: DIFFÉRENT !`);
      Logger.log(`   DIAGNOSTIC dit: ${JSON.stringify(diagnosticValue)}`);
      Logger.log(`   getConfig() dit: ${JSON.stringify(getConfigValue)}`);
    }
  };
  
  compareParam("NB_SOURCES");
  compareParam("NB_DEST");
  compareParam("NIVEAU");
  compareParam("ADMIN_PASSWORD");
  
  // PARTIE 4: Vérifier si les valeurs sont ailleurs dans l'objet config
  Logger.log("\n📋 PARTIE 4: Recherche de NB_SOURCES/NB_DEST ailleurs dans config");
  
  const chercher = (obj, cle, chemin = "config") => {
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const valeur = obj[key];
        
        if (key === cle) {
          Logger.log(`   🔍 Trouvé ${cle} dans ${chemin}.${key} = ${JSON.stringify(valeur)}`);
        }
        
        if (typeof valeur === 'object' && valeur !== null && !Array.isArray(valeur)) {
          chercher(valeur, cle, `${chemin}.${key}`);
        }
      }
    }
  };
  
  chercher(config, "NB_SOURCES");
  chercher(config, "NB_DEST");
  
  Logger.log("\n╔═══════════════════════════════════════════════════════════╗");
  Logger.log("║                    CONCLUSION                             ║");
  Logger.log("╚═══════════════════════════════════════════════════════════╝\n");
  
  if (config.NB_SOURCES === rapport.config.NB_SOURCES && config.NB_DEST === rapport.config.NB_DEST) {
    Logger.log("✅ getConfig() retourne correctement NB_SOURCES et NB_DEST");
    Logger.log("   Le problème est AILLEURS (probablement dans v3_loadConfigForForm)");
  } else if (config.NB_SOURCES === undefined && config.NB_DEST === undefined) {
    Logger.log("❌ getConfig() NE RETOURNE PAS NB_SOURCES et NB_DEST");
    Logger.log("   Le problème est dans la logique de getConfig()");
    Logger.log("   Les valeurs sont LUES depuis _CONFIG mais pas assignées à finalConfig");
  } else {
    Logger.log("⚠️ Résultat inattendu - voir les détails ci-dessus");
  }
  
  Logger.log("\n");
}
