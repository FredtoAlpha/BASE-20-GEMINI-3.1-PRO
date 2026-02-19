/**
 * ==================================================================
 *          CONFIG.GS - Configuration Globale V13 & V2 & Constantes
 * ==================================================================
 * Ce fichier centralise :
 * - Les constantes de configuration par défaut (CONFIG)
 * - Les codes d'erreur standardisés (ERROR_CODES)
 * - Les règles de validation des données (CHECKS)
 * - La fonction robuste pour lire la configuration depuis la feuille _CONFIG (getConfig)
 * - Des utilitaires pour gérer la configuration (createDefaultConfig, updateConfig)
 * ==================================================================
 */

// =================================================
// 1. CONSTANTE DE CONFIGURATION PAR DÉFAUT
// =================================================
const CONFIG = {
  VERSION: "13.0_V2_WIP_Phase5.V12_Integration", // Version mise à jour
  ADMIN_PASSWORD_DEFAULT: "admin123",
  MAX_SWAPS: 50,
  MAX_SWAPS_EVAL: 5000,
  PARITY_TOLERANCE: 2,
  AUTO_RENAME: false,
  TEST_SUFFIX: "TEST",
  DEF_SUFFIX: "DEF",
  HEADER_MOBILITY: 'MOBILITE',

  SHEETS: {
    STRUCTURE: "_STRUCTURE",
    CONFIG: "_CONFIG",
    CONSOLIDATION: "CONSOLIDATION",
    JOURNAL: "_JOURNAL", // Utilisé par Phase5 (via Utils.logAction si configuré ainsi)
    REGISTRE: "_REGISTRE", // Alternative pour logAction si besoin de séparer
    BACKUP: "_BACKUP",
    BILAN_TEST: "BILAN TEST", // Utilisé par le module "Maquette"
    BILAN_DEF: "BILAN DEF",   // Utilisé par Phase5
    BILAN_COMPARE: "BILAN COMPARAISON", // Pour gererVisibiliteOnglets
    STATS_FINAL: "STATISTIQUES FINALES", // Pour gererVisibiliteOnglets
    // Pour Nirvana V2
    V2_OPTI_STRUCTURE: "_NirvanaV2_Calculs",
    V2_BILAN: "_Bilan_Nirvana_V2"
  },

  CRITERES: { // Structure existante, noms utilisés si COLUMN_NAMES ne sont pas prioritaires
    COM: {nom: "Comportement", defaut: 4},
    TRA: {nom: "Travail", defaut: 3},
    PART: {nom: "Participation", defaut: 2},
    ABS: {nom: "Absentéisme", defaut: 1}
  },

  OPTIONS: { // Structure existante
    "6e": ["CHAV", "ITA"],
    "5e": ["CHAV", "LATIN", "ITA"],
    "4e": ["CHAV", "LATIN", "GREC", "ITA"],
    "3e": ["CHAV", "LATIN", "GREC", "ITA"]
  },

  PROTECTED_SHEETS: ["CONSOLIDATION", "_CONFIG", "_STRUCTURE", "ACCUEIL", "_JOURNAL", "_REGISTRE", "_BACKUP"], // Ajout _REGISTRE

  // --- AJOUTS/MODIFICATIONS POUR Phase5.V12.gs et Style "Maquette" ---
  NIVEAU: "5e", // Niveau par défaut pour les exports PDF/Excel de Phase5
  OPTIONS_TO_TRACK_IN_BILAN: ["ITA", "ESP", "GREC", "LATIN", "CHAV"], // Options spécifiques pour le BILAN DEF de Phase5
  LV2_OPTIONS: ["ESP", "ITA"], // Pour distinguer LV2 des autres options lors du comptage

  LOCK_TIMEOUT_PDF: 20000,
  LOCK_TIMEOUT_EXCEL: 45000,
  LOCK_TIMEOUT_RESET: 30000,
  LOCK_TIMEOUT_FINALISER: 30000, // Timeout spécifique pour la fonction finaliser

  // Noms de colonnes CANONIQUES (utilisés si un alias spécifique n'est pas trouvé ou pour référence)
  COLUMN_NAMES: {
    ID_ELEVE: "ID_ELEVE",
    NOM_PRENOM: "NOM & PRENOM",
    SEXE: "SEXE",
    LV2: "LV2",
    OPT: "OPT",
    COM: "COM", // Si les en-têtes sont juste "COM", "TRA"...
    TRA: "TRA",
    PART: "PART",
    ABS: "ABS",
    INDICATEUR: "INDICATEUR",
    ASSO: "ASSO",
    DISSO: "DISSO",
    CLASSE_DEF: "CLASSE DEF",
    SOURCE: "SOURCE",
    MOBILITE: "MOBILITE", // Redondant avec HEADER_MOBILITY mais bon pour la centralisation
    RAISON_PLACEMENT: "RAISON PLACEMENT"
  },

  // ALIAS pour les colonnes (Phase5 utilisera ces listes pour trouver les colonnes)
  COLUMN_ALIASES: {
    ID_ELEVE:   ["ID_ELEVE", "ID", "IDENTIFIANT", "IDELEVE", "Num Elève"],
    NOM_PRENOM: ["NOM & PRENOM", "NOM_PRENOM", "NOM PRENOM", "ELEVE", "NOM ET PRENOM", "NOM"],
    SEXE:       ["SEXE", "GENRE", "S"],
    LV2:        ["LV2", "LANGUE", "LANGUE VIVANTE 2", "LV"],
    OPT:        ["OPT", "OPTION", "OPTIONS"],
    COM:        ["COM", "COMPORTEMENT", "COMP", "H"],
    TRA:        ["TRA", "TRAVAIL", "TRV", "I"],
    PART:       ["PART", "PARTICIPATION", "PARTI", "J"],
    ABS:        ["ABS", "ABSENCES", "ASSIDUITE", "K"],
    INDICATEUR: ["INDICATEUR", "IND", "L", "IND."],
    ASSO:       ["ASSO", "ASSOCIATION", "ASSOC", "M"],
    DISSO:      ["DISSO", "DISSOCIATION", "DISSOC", "N"],
    CLASSE_DEF: ["CLASSE DEF", "CLASSE DEFINITIVE", "CLASSE_DEF", "CLASSE FINALE", "CLASSE_FINALE", "R"], // R pour colonne Maquette
    SOURCE:     ["SOURCE", "ORIGINE"],
    MOBILITE:   ["MOBILITE", 'MOBILITE'] // S'assurer que HEADER_MOBILITY est inclus
  },

  // Styles pour Phase5 inspirés de "Maquette"
  STYLE: {
    // Général onglet
    HEADER_BG: "#C6E0B4", // Vert clair Maquette
    HEADER_FONT_WEIGHT: "bold",
    HEADER_H_ALIGN: "center",
    HEADER_V_ALIGN: "middle",
    EVEN_ROW_BG: "#f8f9fa",    // Zébrure standard
    SEPARATOR_ROW_BG: "#dcdcdc", // Ligne de séparation grise
    DEF_TAB_COLOR: "#1a73e8",    // Bleu pour onglet DEF Phase5
    NOTE_FONT_COLOR: "#666666",  // Couleur pour la note "Classe définitive..."

    // Largeurs de colonnes (pixels)
    WIDTH_NOM_PRENOM: 200,
    WIDTH_INFO_SHORT: 70,
    WIDTH_CRITERE: 60,
    WIDTH_CLASSE_DEF: 100,

    // Formatage conditionnel "Maquette" (utilisé par Phase5_applyConditionalFormattingMaquetteStyle)
    SEXE_M_COLOR: "#4F81BD", // Bleu Maquette pour Garçon
    SEXE_F_COLOR: "#F28EA8", // Rose Maquette pour Fille
    LV2_COLORS: { // Utilisé pour formattage conditionnel LV2
        ESP: "#E59838",   // Orange Maquette pour ESP
        ITA: "#73C6B6",   // Vert d'eau pour ITA (de miseEnFormeOngletUnique)
        AUTRE: "#A3E4D7"  // Bleu-vert clair (suggéré pour autres LV2 si besoin)
    },
    OPTION_FORMATS: [ // Structure pour le formatage des options
      { text:"GREC",  bgColor:"#C0392B", fgColor:"#FFFFFF" },
      { text:"LATIN", bgColor:"#641E16", fgColor:"#FFFFFF" },
      { text:"CHAV",  bgColor:"#6C3483", fgColor:"#FFFFFF" },
      { text:"UPE2A", bgColor:"#D5D8DC", fgColor:"#000000" }, // Noir sur gris
      { text:"LLCA",  bgColor:"#F4B084", fgColor:"#000000" }  // Exemple si LLCA existe
    ],
    SCORE_COLORS: { // Couleurs de fond pour les scores 1-4
        S1: "#FF0000", // Rouge
        S2: "#FFD966", // Jaune
        S3: "#3CB371", // Vert moyen
        S4: "#006400"  // Vert foncé
    },
    SCORE_FONT_COLORS: { // Couleurs de police pour les scores (pour contraste)
        S1: "#FFFFFF",
        S2: "#000000",
        S3: "#FFFFFF", // #000000 si #3CB371 est trop clair
        S4: "#FFFFFF"
    },

    // Style pour les stats écrites SOUS la feuille (inspiré de calculerStatistiquesTEST et writeSheetStats Maquette)
    STATS_TITLE_BG: '#b6d7a8',          // Vert clair (comme header Maquette)
    STATS_TITLE_FONT_WEIGHT: 'bold',
    STATS_ROW_COUNT: 7,                 // Nb de lignes pour les stats affichées (Titre + 5 lignes de données + Moyenne)
    STATS_BORDER_COLOR: '#000000',      // Bordure noire pour le bloc stats
    // Styles spécifiques par cellule de stat (pour _phase5_ajouterStatsStyleMaquette)
    STATS_SEXE_F_CELL:      { align: 'center', bg: '#F28EA8', fg: '#000000', bold: true }, // Rose Maquette, police noire
    STATS_SEXE_M_CELL:      { align: 'center', bg: '#4F81BD', fg: '#FFFFFF', bold: true }, // Bleu Maquette, police blanche
    STATS_LV2_ESP_CELL:     { align: 'center', bg: '#E59838', fg: '#000000', bold: true }, // Orange Maquette
    STATS_LV2_AUTRE_CELL:   { align: 'center', bg: '#A3E4D7', fg: '#000000', bold: true }, // Autre LV2
    STATS_OPTIONS_TOTAL_CELL:{ align: 'center', bold: true },                               // Options total
    // Pour les scores, utiliser les SCORE_COLORS et SCORE_FONT_COLORS définis plus haut
    // Exemple pour S4 (le reste sera dérivé dans la fonction d'écriture)
    SCORE_CELL_S4:          { align: 'center', bg: '#006400', fg: '#FFFFFF', bold: true },
    SCORE_CELL_S3:          { align: 'center', bg: '#3CB371', fg: '#FFFFFF', bold: true },
    SCORE_CELL_S2:          { align: 'center', bg: '#FFD966', fg: '#000000', bold: true },
    SCORE_CELL_S1:          { align: 'center', bg: '#FF0000', fg: '#FFFFFF', bold: true },
    STATS_AVERAGE_CELL:     { align: 'center', bold: true, fmt: '0.00', bg: '#34495e', fg: '#ffffff'}, // Style V7 pour moyenne

    // Styles pour le Bilan DEF (Phase 5)
    BILAN_HEADER_BG: "#1a73e8",           // Bleu foncé pour titre Bilan DEF
    BILAN_HEADER_FG: "#ffffff",           // Texte blanc pour titre Bilan DEF
    BILAN_SECTION_BG: "#E0E0E0",          // Gris clair pour titres de section
    BILAN_TABLE_HEADER_BG: "#F8F9FA",     // Gris très clair pour en-têtes de tableau
    BILAN_TABLE_BORDER_COLOR: "#cccccc",  // Gris pour bordures de tableau
    BILAN_DEF_TAB_COLOR: "#1a73e8",       // Couleur onglet Bilan DEF

    // Buffers pour nettoyage de feuille
    CLEANUP_BUFFER_ROWS: 3,
    CLEANUP_BUFFER_COLS: 1,

    // Textes / Libellés (si vous voulez les centraliser)
    TEXT_LABELS: {
        FILLES: "Filles",
        GARCONS: "Garçons"
        // ... autres libellés ...
    }
  },

  // Section FORMAT existante (pourrait être fusionnée dans STYLE à terme)
  FORMAT: {
    COLORS: { GARCON: "#aed6f1", FILLE: "#f5eef8", ESP: "#eb984e" } // Ces couleurs sont-elles encore utilisées ou remplacées par STYLE?
  },

  // ========= SECTION POUR NIRVANA V2 - VALEURS PAR DÉFAUT (INCHANGÉE) =========
  V2_SHEET_NAME_STRUCTURE: "_STRUCTURE",
  V2_STRUCTURE_HEADER_CLASSE: "CLASSE DEST",
  V2_STRUCTURE_HEADER_COM_S4: "Cap_COM_S4",
  V2_STRUCTURE_HEADER_COM_S1: "Cap_COM_S1",
  V2_STRUCTURE_HEADER_TRA_S4: "Cap_TRA_S4",
  V2_STRUCTURE_HEADER_TRA_S1: "Cap_TRA_S1",
  V2_STRUCTURE_HEADER_PART_S4: "Cap_PART_S4",
  V2_STRUCTURE_HEADER_PART_S1: "Cap_PART_S1",
  V2_STRUCTURE_HEADER_ABS_S1: "Cap_ABS_S1",
  V2_DEFAULT_CAPACITY_IF_UNDEFINED: 0,
  V2_MOBILITES_CONSIDEREES_FIXES: ['FIXE', 'SPEC'],
  V2_TOLERANCE_DELTA_EXTREME: 0,
  V2_MAX_ITER_PASSE_EXTREME: 200,
  V2_MAX_ITER_PASSE_MOYENS: 100,
  V2_MAX_ITER_PASSE_ABS: 100,
  V2_OPTI_STRUCTURE_VISIBLE: true,
  V2_SHEET_NAME_OPTI_STRUCTURE: "_NirvanaV2_Calculs",
  V2_SHEET_NAME_BILAN: "_Bilan_Nirvana_V2",
  V2_CRITERES_A_EQUILIBRER: ['COM', 'TRA', 'PART', 'ABS'],
  V2_DIMENSIONS_PRIO_EXTREMES: ["COM", "TRA", "PART"],
  V2_SCORES_PRIO_EXTREMES: [4, 1]
  // =====================================================================
};

// =================================================
// 2. CODES D'ERREUR STANDARDISÉS
// =================================================
const ERROR_CODES = {
  UNCAUGHT_EXCEPTION          : "ERR_UNCAUGHT_EXCEPTION",
  MISSING_STRUCTURE_SHEET     : "ERR_MISSING_STRUCTURE_SHEET",
  INVALID_STRUCTURE           : "ERR_INVALID_STRUCTURE",
  LESS_THAN_TWO_CLASSES       : "ERR_LESS_THAN_TWO_CLASSES",
  NO_TEST_SHEETS              : "ERR_NO_TEST_SHEETS",
  NO_STUDENTS_FOUND           : "ERR_NO_STUDENTS_FOUND",
  MISSING_CRITICAL_DATA_COLUMN: "ERR_MISSING_DATA_COLUMN",
  SWAP_WRITE_FAILED           : "ERR_SWAP_WRITE_FAILED",
  JOURNAL_SAVE_FAILED         : "ERR_JOURNAL_SAVE_FAILED",
  MISSING_CONFIG_SHEET        : "ERR_MISSING_CONFIG_SHEET",
  VB_WRITE_FAILED             : "ERR_VB_WRITE_FAILED",
  VB_BLOC_UNPLACED            : "ERR_VB_BLOC_UNPLACED",

  // AJOUTS POUR PHASE 5
  LOCK_TIMEOUT: "ERR_LOCK_TIMEOUT",
  ACTION_NOT_RECOGNIZED: "ERR_ACTION_NOT_RECOGNIZED",
  NO_CLASSES_SELECTED: "ERR_NO_CLASSES_SELECTED",
  PDF_EXPORT_FAILED: "ERR_PDF_EXPORT_FAILED",
  EXCEL_EXPORT_FAILED: "ERR_EXCEL_EXPORT_FAILED",
  RESET_FAILED: "ERR_RESET_FAILED",

  // Codes spécifiques V2
  V2_PREPARATION_FAILED       : "ERR_V2_PREPARATION_FAILED",
  V2_CAPACITY_LOAD_FAILED     : "ERR_V2_CAPACITY_LOAD_FAILED"
};

// =================================================
// 3. RÈGLES DE VALIDATION DES DONNÉES (CHECKS) - (INCHANGÉ)
// =================================================
const CHECKS = {
  ID_ELEVE:{required:true, validator:v=> v && ((typeof v === 'string' && v.startsWith("AUTO_")) || (!isNaN(Number(v)) && Number(v)>0)), message:'ID_ELEVE manquant ou invalide (non numérique ou <= 0)', unique:true},
  CLASSE  :{required:true, validator:()=>true, message:'Validation dynamique dans sanitizeStudents'},
  NOM     :{required:false, validator:v=> v === null || v === undefined || v === '' || (typeof v === 'string' && v.trim().length>0), message:'NOM vide si renseigné'},
  COM     :{required:false, validator:v=>v===''||v===null||v===undefined||(!isNaN(Number(v))&&Number(v)>=0&&Number(v)<=4), message:'Note COM invalide (doit être 0-4 ou vide)'},
  TRA     :{required:false, validator:v=>v===''||v===null||v===undefined||(!isNaN(Number(v))&&Number(v)>=0&&Number(v)<=4), message:'Note TRA invalide (doit être 0-4 ou vide)'},
  PART    :{required:false, validator:v=>v===''||v===null||v===undefined||(!isNaN(Number(v))&&Number(v)>=0&&Number(v)<=4), message:'Note PART invalide (doit être 0-4 ou vide)'},
  OPT     :{required:false, validator:()=>true, message:''},
  ASSO    :{required:false, validator:()=>true, message:''},
  DISSO   :{required:false, validator:()=>true, message:''},
  FIXE    :{required:false, validator:()=>true, message:''},
  LV2     :{required:false, validator:()=>true, message:''},
  MOBILITE:{required:false, validator:v=> v===''||v===null||v===undefined||['FIXE','PERMUT','CONDI','SPEC','LIBRE'].includes(String(v).toUpperCase()), message:'Valeur MOBILITE invalide'}
};

// =================================================
// 4. FONCTION DE LECTURE DE CONFIGURATION (Adaptée pour lire les nouvelles sections)
// =================================================
// =================================================
// 4. FONCTION DE LECTURE DE CONFIGURATION (CORRIGÉE pour OPT/LV2)
// =================================================
function getConfig() {
  Logger.log(">>> DEBUT getConfig() DANS Config.gs - VÉRIFICATION ERROR_CODES (Version 12/05 - Complète) <<<");
  
  let baseConfigCopy;
  try {
      baseConfigCopy = JSON.parse(JSON.stringify(CONFIG)); 
  } catch(e) {
      Logger.log("ERREUR getConfig: Copie Profonde de CONFIG a échoué: " + e + ". Utilisation d'un objet vide pour baseConfigCopy.");
      baseConfigCopy = { 
          SHEETS:{}, OPTIONS:{}, CRITERES:{}, FORMAT:{}, 
          COLUMN_NAMES:{}, COLUMN_ALIASES:{}, 
          STYLE:{ LV2_COLORS:{}, OPTION_FORMATS:[], SCORE_COLORS:{}, SCORE_FONT_COLORS:{}, TEXT_LABELS:{} }
      }; 
  }

  // Assurer que les sous-objets principaux existent dans baseConfigCopy
  baseConfigCopy.SHEETS = baseConfigCopy.SHEETS || (CONFIG.SHEETS || {});
  baseConfigCopy.OPTIONS = baseConfigCopy.OPTIONS || (CONFIG.OPTIONS || {});
  baseConfigCopy.CRITERES = baseConfigCopy.CRITERES || (CONFIG.CRITERES || {});
  baseConfigCopy.FORMAT = baseConfigCopy.FORMAT || (CONFIG.FORMAT || {}); 
  baseConfigCopy.COLUMN_NAMES = baseConfigCopy.COLUMN_NAMES || (CONFIG.COLUMN_NAMES || {});
  baseConfigCopy.COLUMN_ALIASES = baseConfigCopy.COLUMN_ALIASES || (CONFIG.COLUMN_ALIASES || {});
  baseConfigCopy.STYLE = baseConfigCopy.STYLE || (CONFIG.STYLE || {});
    
  if (baseConfigCopy.STYLE) { 
      baseConfigCopy.STYLE.LV2_COLORS = baseConfigCopy.STYLE.LV2_COLORS || ((CONFIG.STYLE||{}).LV2_COLORS || {});
      baseConfigCopy.STYLE.OPTION_FORMATS = baseConfigCopy.STYLE.OPTION_FORMATS || ((CONFIG.STYLE||{}).OPTION_FORMATS || []);
      baseConfigCopy.STYLE.SCORE_COLORS = baseConfigCopy.STYLE.SCORE_COLORS || ((CONFIG.STYLE||{}).SCORE_COLORS || {});
      baseConfigCopy.STYLE.SCORE_FONT_COLORS = baseConfigCopy.STYLE.SCORE_FONT_COLORS || ((CONFIG.STYLE||{}).SCORE_FONT_COLORS || {});
      baseConfigCopy.STYLE.TEXT_LABELS = baseConfigCopy.STYLE.TEXT_LABELS || ((CONFIG.STYLE||{}).TEXT_LABELS || {});
  }

  let finalConfig = { ...baseConfigCopy }; 
  
  try {
      finalConfig.ERROR_CODES = JSON.parse(JSON.stringify(ERROR_CODES)); 
      Logger.log("getConfig: finalConfig.ERROR_CODES initialisé depuis constante globale ERROR_CODES.");
  } catch (e) {
      Logger.log("ERREUR getConfig: Copie Profonde de ERROR_CODES globale a échoué: " + e);
      finalConfig.ERROR_CODES = {}; 
  }

  finalConfig._CONFIG_OK_SHEET_READ = false; 
  finalConfig._CONFIG_ERROR_SHEET_READ = null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheetName = (finalConfig.SHEETS && finalConfig.SHEETS.CONFIG) ? finalConfig.SHEETS.CONFIG 
                           : (CONFIG.SHEETS && CONFIG.SHEETS.CONFIG) ? CONFIG.SHEETS.CONFIG 
                           : "_CONFIG";
    const configSheet = ss.getSheetByName(configSheetName);

    if (configSheet) {
      Logger.log(`getConfig: Lecture feuille config: '${configSheetName}'...`);
      const data = configSheet.getDataRange().getValues();
      let headers = [];
      let headerRowIndex = -1;
      
      for (let i=0; i < Math.min(3, data.length); i++) { 
          const potentialHeaders = data[i].map(h => String(h || '').trim().toUpperCase());
          // ✅ Accepter PARAMETRE/VALEUR (sans accent) OU PARAMÈTRE/VALEUR (avec accent)
          const hasParam = potentialHeaders.includes("PARAMETRE") || potentialHeaders.includes("PARAMÈTRE");
          const hasValeur = potentialHeaders.includes("VALEUR");
          if (hasParam && hasValeur) {
              headers = potentialHeaders; headerRowIndex = i; break;
          }
      }
      
      if (headerRowIndex === -1) {
           Logger.log(`WARN getConfig: Colonnes PARAMETRE/VALEUR non trouvées dans '${configSheetName}'.`);
      } else {
          // ✅ Chercher PARAMETRE (sans accent) ou PARAMÈTRE (avec accent)
          const paramIndex = headers.indexOf("PARAMETRE") !== -1 ? headers.indexOf("PARAMETRE") : headers.indexOf("PARAMÈTRE");
          const valueIndex = headers.indexOf("VALEUR");

          // 🔍 DEBUG: Compter les lignes totales et afficher les 10 premières
          Logger.log(`📊 Total de lignes dans _CONFIG: ${data.length - headerRowIndex - 1}`);
          
          for (let i = headerRowIndex + 1; i < data.length; i++) { 
            const paramKeyRaw = String(data[i][paramIndex] || '').trim();
            const paramKeyUpper = paramKeyRaw.toUpperCase(); 
            let value = data[i][valueIndex];

            // 🔍 DEBUG: Afficher les 10 premières lignes pour diagnostic
            if (i <= headerRowIndex + 10) {
              Logger.log(`🔍 _CONFIG ligne ${i+1}: param="${paramKeyRaw}" (upper: "${paramKeyUpper}") | valeur="${value}"`);
            }

            if (!paramKeyRaw) continue;
            
            // 🔍 DEBUG: Tracer spécifiquement NB_SOURCES et NB_DEST
            if (paramKeyUpper === "NB_SOURCES" || paramKeyUpper === "NB_DEST") {
              Logger.log(`✨ TROUVÉ ligne ${i+1}: param="${paramKeyRaw}" | valeur BRUTE="${value}" | type=${typeof value}`);
            } 

            if (typeof value === 'string') {
                const valueLower = value.toLowerCase().trim();
                if (valueLower === 'true') value = true;
                else if (valueLower === 'false') value = false;
                else if (value.trim() !== '' && !isNaN(Number(value))) value = Number(value);
            }
            
            finalConfig.SHEETS = finalConfig.SHEETS || {};
            finalConfig.OPTIONS = finalConfig.OPTIONS || {};
            finalConfig.STYLE = finalConfig.STYLE || {};
            finalConfig.COLUMN_NAMES = finalConfig.COLUMN_NAMES || {};
            finalConfig.COLUMN_ALIASES = finalConfig.COLUMN_ALIASES || {};
            finalConfig.ERROR_CODES = finalConfig.ERROR_CODES || {};

            // =====================================
            // CORRECTION PRINCIPALE : Traiter OPT et LV2
            // =====================================
            if (paramKeyUpper === "OPT" && value) {
                // Récupérer le niveau pour mettre à jour les bonnes options
                const niveauActuel = finalConfig.NIVEAU;
                if (niveauActuel) {
                    const niveauKey = niveauActuel.toLowerCase().replace('°', 'e'); // "4°" → "4e"
                    finalConfig.OPTIONS[niveauKey] = String(value).split(',')
                        .map(item => item.trim().toUpperCase())
                        .filter(Boolean);
                    Logger.log(`getConfig: Options personnalisées mises à jour pour ${niveauKey}: ${finalConfig.OPTIONS[niveauKey].join(', ')}`);
                }
            } else if (paramKeyUpper === "LV2" && value) {
                // Traiter les LV2 personnalisées
                finalConfig.LV2_OPTIONS = String(value).split(',')
                    .map(item => item.trim().toUpperCase())
                    .filter(Boolean);
                Logger.log(`getConfig: LV2 personnalisées mises à jour: ${finalConfig.LV2_OPTIONS.join(', ')}`);
            }
            // =====================================
            // FIN CORRECTION
            // =====================================
            else if (paramKeyUpper.startsWith("SHEETS_")) {
                 const sheetKeySimple = paramKeyUpper.replace("SHEETS_", "").trim();
                 if(sheetKeySimple && value !== undefined && value !== null) finalConfig.SHEETS[sheetKeySimple] = String(value).trim();
            } else if (paramKeyUpper.startsWith("OPTIONS_") && paramKeyUpper !== "OPTIONS_TO_TRACK_IN_BILAN") {
                 const niveauOpt = paramKeyUpper.replace("OPTIONS_", "").trim().toLowerCase();
                 if (niveauOpt && value !== undefined && value !== null) {
                     finalConfig.OPTIONS[niveauOpt] = String(value).split(',').map(item => item.trim().toUpperCase()).filter(Boolean);
                 }
            } else if (paramKeyUpper.startsWith("STYLE_")) {
                const styleKey = paramKeyRaw.substring("STYLE_".length); 
                if (styleKey && finalConfig.STYLE.hasOwnProperty(styleKey)) { 
                    if (typeof finalConfig.STYLE[styleKey] === 'object' && finalConfig.STYLE[styleKey] !== null && typeof value === 'string') {
                        try { finalConfig.STYLE[styleKey] = JSON.parse(value); }
                        catch (e) { Logger.log(`WARN getConfig: STYLE_${styleKey} - Impossible de parser JSON: ${value}`); }
                    } else { finalConfig.STYLE[styleKey] = value; }
                } else if (styleKey) { finalConfig.STYLE[styleKey] = value; }
            } else if (paramKeyUpper.startsWith("COLUMN_NAMES_")) {
                 const colKey = paramKeyUpper.replace("COLUMN_NAMES_", "").trim();
                 if(colKey && value !== undefined && value !== null) finalConfig.COLUMN_NAMES[colKey] = String(value).trim();
            } else if (paramKeyUpper.startsWith("COLUMN_ALIASES_")) {
                 const colKey = paramKeyUpper.replace("COLUMN_ALIASES_", "").trim();
                  if(colKey && value !== undefined && value !== null) {
                     finalConfig.COLUMN_ALIASES[colKey] = String(value).split(',').map(item => item.trim()).filter(Boolean);
                 }
            } else if (paramKeyUpper.startsWith("ERROR_CODES_")) {
                const errKey = paramKeyUpper.replace("ERROR_CODES_", "").trim();
                 if(errKey && value !== undefined && value !== null) { 
                     finalConfig.ERROR_CODES[errKey] = String(value).trim(); 
                 }
            } else if (paramKeyUpper.startsWith("V2_")) {
                 if (["V2_MOBILITES_CONSIDEREES_FIXES", "V2_CRITERES_A_EQUILIBRER", "V2_DIMENSIONS_PRIO_EXTREMES", "V2_SCORES_PRIO_EXTREMES"].includes(paramKeyUpper)) {
                    if (typeof value === 'string') {
                         finalConfig[paramKeyRaw] = value.split(',').map(item => item.trim()).filter(Boolean);
                         if (paramKeyUpper === "V2_SCORES_PRIO_EXTREMES") finalConfig[paramKeyRaw] = finalConfig[paramKeyRaw].map(Number).filter(n => !isNaN(n));
                    } else if (Array.isArray(value)) { finalConfig[paramKeyRaw] = value.map(item => String(item).trim());
                    } else { finalConfig[paramKeyRaw] = value; }
                } else { finalConfig[paramKeyRaw] = value; }
            } else { 
                 if (finalConfig.hasOwnProperty(paramKeyRaw)) { finalConfig[paramKeyRaw] = value;
                 } else if (finalConfig.hasOwnProperty(paramKeyUpper) && paramKeyUpper === paramKeyRaw.toUpperCase()) { finalConfig[paramKeyUpper] = value;
                 } else { finalConfig[paramKeyRaw] = value; }
            }
          }
          finalConfig._CONFIG_OK_SHEET_READ = true;
          Logger.log(`getConfig: Configuration fusionnée avec les valeurs de la feuille '${configSheetName}'.`);
      }
    } else {
      Logger.log(`INFO getConfig: Feuille de configuration '${configSheetName}' absente.`);
    }
  } catch (e) {
    Logger.log(`⚠️ ERREUR SEVERE getConfig lors de la lecture/fusion de la feuille _CONFIG : ${e.message}\n${e.stack}`);
    finalConfig._CONFIG_ERROR_SHEET_READ = `Erreur lecture _CONFIG: ${e.message}`;
  }

  // Fonction de consolidation des valeurs par défaut
  const ensureDeepDefaults = (target, defaultsFromConst) => {
    if (!defaultsFromConst || typeof defaultsFromConst !== 'object') return target;
    target = target || (Array.isArray(defaultsFromConst) ? [] : {});

    for (const key in defaultsFromConst) {
      if (defaultsFromConst.hasOwnProperty(key)) {
        if (target[key] === undefined) {
          target[key] = JSON.parse(JSON.stringify(defaultsFromConst[key]));
        } else if (typeof defaultsFromConst[key] === 'object' && defaultsFromConst[key] !== null && !Array.isArray(defaultsFromConst[key])) {
          target[key] = target[key] || {};
          ensureDeepDefaults(target[key], defaultsFromConst[key]);
        }
      }
    }
    
    if (target === finalConfig.STYLE && finalConfig.STYLE) {
        finalConfig.STYLE.LV2_COLORS = finalConfig.STYLE.LV2_COLORS || ((CONFIG.STYLE||{}).LV2_COLORS || {});
        finalConfig.STYLE.OPTION_FORMATS = finalConfig.STYLE.OPTION_FORMATS || ((CONFIG.STYLE||{}).OPTION_FORMATS || []);
        finalConfig.STYLE.SCORE_COLORS = finalConfig.STYLE.SCORE_COLORS || ((CONFIG.STYLE||{}).SCORE_COLORS || {});
        finalConfig.STYLE.SCORE_FONT_COLORS = finalConfig.STYLE.SCORE_FONT_COLORS || ((CONFIG.STYLE||{}).SCORE_FONT_COLORS || {});
        finalConfig.STYLE.TEXT_LABELS = finalConfig.STYLE.TEXT_LABELS || ((CONFIG.STYLE||{}).TEXT_LABELS || {});
    }
    return target;
  };

  // Consolidation finale
  finalConfig.NIVEAU = finalConfig.NIVEAU !== undefined ? finalConfig.NIVEAU : (CONFIG.NIVEAU || "NiveauIndéfini");
  finalConfig.TEST_SUFFIX = finalConfig.TEST_SUFFIX !== undefined ? finalConfig.TEST_SUFFIX : (CONFIG.TEST_SUFFIX || "TEST");
  finalConfig.DEF_SUFFIX = finalConfig.DEF_SUFFIX !== undefined ? finalConfig.DEF_SUFFIX : (CONFIG.DEF_SUFFIX || "DEF");

  finalConfig.SHEETS = ensureDeepDefaults(finalConfig.SHEETS, CONFIG.SHEETS);
  finalConfig.OPTIONS = ensureDeepDefaults(finalConfig.OPTIONS, CONFIG.OPTIONS);
  finalConfig.CRITERES = ensureDeepDefaults(finalConfig.CRITERES, CONFIG.CRITERES);
  finalConfig.COLUMN_NAMES = ensureDeepDefaults(finalConfig.COLUMN_NAMES, CONFIG.COLUMN_NAMES);
  finalConfig.COLUMN_ALIASES = ensureDeepDefaults(finalConfig.COLUMN_ALIASES, CONFIG.COLUMN_ALIASES);
  finalConfig.STYLE = ensureDeepDefaults(finalConfig.STYLE, CONFIG.STYLE);
  
  let mergedErrorCodes = JSON.parse(JSON.stringify(ERROR_CODES)); 
  if (finalConfig.ERROR_CODES && typeof finalConfig.ERROR_CODES === 'object') {
      for (const key in finalConfig.ERROR_CODES) { 
          if (finalConfig.ERROR_CODES.hasOwnProperty(key)) {
              mergedErrorCodes[key] = finalConfig.ERROR_CODES[key];
          }
      }
  }
  finalConfig.ERROR_CODES = mergedErrorCodes;
  
  // ============================================================
  // COMPATIBILITÉ : Exposer LV2 et OPT comme propriétés directes (ARRAYS)
  // ============================================================
  // Certains modules (ListesDeroulantes.gs) s'attendent à ce que
  // config.LV2 et config.OPT soient des ARRAYS, pas des strings
  // On crée donc des alias pointant vers les vrais arrays
  if (finalConfig.LV2_OPTIONS && Array.isArray(finalConfig.LV2_OPTIONS)) {
    finalConfig.LV2 = finalConfig.LV2_OPTIONS; // Alias array (pas string!)
  }
  
  if (finalConfig.OPTIONS && finalConfig.NIVEAU) {
    const niveauKey = String(finalConfig.NIVEAU).toLowerCase().replace('°', 'e');
    if (finalConfig.OPTIONS[niveauKey] && Array.isArray(finalConfig.OPTIONS[niveauKey])) {
      finalConfig.OPT = finalConfig.OPTIONS[niveauKey]; // Alias array (pas string!)
    }
  }
  
  Logger.log(`✅ Compatibilité: LV2=${JSON.stringify(finalConfig.LV2 || [])} | OPT=${JSON.stringify(finalConfig.OPT || [])}`);
  // ============================================================

  Logger.log(">>> getConfig: Configuration finale prête. UNCAUGHT_EXCEPTION: " + (finalConfig.ERROR_CODES ? finalConfig.ERROR_CODES.UNCAUGHT_EXCEPTION : "ERROR_CODES non défini"));
  Logger.log(">>> FIN getConfig() DANS Config.gs <<<");
  return finalConfig;
}



// =================================================
// 5. UTILITAIRES DE CONFIGURATION
// =================================================
function createDefaultConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheetName = CONFIG.SHEETS?.CONFIG || "_CONFIG"; 
  let configSheet = ss.getSheetByName(configSheetName);
  if (!configSheet) configSheet = ss.insertSheet(configSheetName);
  else configSheet.clearContents().clearFormats(); // Efface aussi les formats

  const configData = [
    ["PARAMETRE", "VALEUR", "DESCRIPTION (Optionnelle)", "TYPE ATTENDU (pour info)"]
  ];
  
  const addConfigRow = (key, value, description = "", typeHint = "") => {
      let displayValue = value;
      if (Array.isArray(value)) {
          displayValue = value.join(',');
          typeHint = typeHint || "Liste (séparée par virgules)";
      } else if (typeof value === 'object' && value !== null) {
          try { displayValue = JSON.stringify(value, null, 2); typeHint = typeHint || "JSON"; } // Indent for readability
          catch(e) { displayValue = "Erreur de sérialisation JSON"; }
      } else {
          typeHint = typeHint || (typeof value);
      }
      configData.push([key, displayValue, description, typeHint]);
  };

  // Sections de CONFIG à écrire
  addConfigRow("VERSION", CONFIG.VERSION, "Version de l'application", "Texte");
  addConfigRow("NIVEAU", CONFIG.NIVEAU, "Niveau scolaire traité (ex: 5e)", "Texte");
  addConfigRow("TEST_SUFFIX", CONFIG.TEST_SUFFIX, "Suffixe pour les onglets de travail", "Texte");
  addConfigRow("DEF_SUFFIX", CONFIG.DEF_SUFFIX, "Suffixe pour les onglets finalisés", "Texte");
  addConfigRow("HEADER_MOBILITY", CONFIG.HEADER_MOBILITY, "Nom exact de la colonne Mobilité", "Texte");
  addConfigRow("ADMIN_PASSWORD_DEFAULT", CONFIG.ADMIN_PASSWORD_DEFAULT, "Mot de passe admin par défaut (sera écrasé si défini)", "Texte");
  addConfigRow("MAX_SWAPS", CONFIG.MAX_SWAPS, "Nombre max de swaps appliqués", "Nombre");
  addConfigRow("MAX_SWAPS_EVAL", CONFIG.MAX_SWAPS_EVAL, "Nombre max de swaps évalués", "Nombre");
  addConfigRow("PARITY_TOLERANCE", CONFIG.PARITY_TOLERANCE, "Tolérance pour l'équilibre de parité", "Nombre");
  addConfigRow("AUTO_RENAME", CONFIG.AUTO_RENAME, "Renommage auto (true/false)", "Booléen");
  addConfigRow("OPTIONS_TO_TRACK_IN_BILAN", CONFIG.OPTIONS_TO_TRACK_IN_BILAN, "Options suivies dans le Bilan DEF", "Liste");
  addConfigRow("LV2_OPTIONS", CONFIG.LV2_OPTIONS, "LV2 considérées comme options distinctes", "Liste");
  addConfigRow("LOCK_TIMEOUT_PDF", CONFIG.LOCK_TIMEOUT_PDF, "Timeout (ms) pour PDF", "Nombre");
  addConfigRow("LOCK_TIMEOUT_EXCEL", CONFIG.LOCK_TIMEOUT_EXCEL, "Timeout (ms) pour Excel", "Nombre");
  // ... autres clés simples ...

  Object.keys(CONFIG.SHEETS).forEach(key => addConfigRow(`SHEETS_${key.toUpperCase()}`, CONFIG.SHEETS[key], `Nom onglet: ${key}`));
  Object.keys(CONFIG.CRITERES).forEach(key => addConfigRow(`CRITERES_${key.toUpperCase()}_NOM`, CONFIG.CRITERES[key].nom, `Nom du critère ${key}`));
  Object.keys(CONFIG.OPTIONS).forEach(key => addConfigRow(`OPTIONS_${key}`, CONFIG.OPTIONS[key], `Options pour niveau ${key}`));
  addConfigRow("PROTECTED_SHEETS", CONFIG.PROTECTED_SHEETS, "Onglets protégés contre suppression/réinitialisation");

  Object.keys(CONFIG.COLUMN_NAMES).forEach(key => addConfigRow(`COLUMN_NAMES_${key.toUpperCase()}`, CONFIG.COLUMN_NAMES[key], `Nom canonique colonne: ${key}`));
  Object.keys(CONFIG.COLUMN_ALIASES).forEach(key => addConfigRow(`COLUMN_ALIASES_${key.toUpperCase()}`, CONFIG.COLUMN_ALIASES[key], `Alias pour colonne: ${key}`));
  
  // Écrire STYLE (chaque sous-clé de STYLE devient une ligne dans _CONFIG)
  Object.keys(CONFIG.STYLE).forEach(styleKey => {
      const fullKey = `STYLE_${styleKey.toUpperCase()}`;
      addConfigRow(fullKey, CONFIG.STYLE[styleKey], `Configuration de style: ${styleKey}`);
  });
  
  // Écrire FORMAT (similaire à STYLE)
  if (CONFIG.FORMAT && CONFIG.FORMAT.COLORS) {
     Object.keys(CONFIG.FORMAT.COLORS).forEach(colorKey => {
         addConfigRow(`FORMAT_COLORS_${colorKey.toUpperCase()}`, CONFIG.FORMAT.COLORS[colorKey], `Couleur format: ${colorKey}`);
     });
  }

  // Écrire les clés V2 (celles commençant par V2_)
  Object.keys(CONFIG).forEach(key => {
    if (key.startsWith("V2_")) {
       addConfigRow(key, CONFIG[key], `Paramètre Nirvana V2: ${key}`);
    }
  });
  
  // Écrire les ERROR_CODES (pour info, moins susceptibles d'être modifiés par l'utilisateur)
  // Object.keys(CONFIG.ERROR_CODES).forEach(key => addConfigRow(`ERROR_CODES_${key.toUpperCase()}`, CONFIG.ERROR_CODES[key], `Code d'erreur: ${key}`));


  configSheet.getRange(1, 1, configData.length, configData[0].length).setValues(configData);
  configSheet.getRange(1, 1, 1, configData[0].length).setFontWeight("bold").setBackground("#f3f3f3");
  try {configSheet.autoResizeColumns(1, 4); } catch(e) {Logger.log("Erreur autoResize: " + e);}
  configSheet.setFrozenRows(1);

  Logger.log(`Feuille de configuration '${configSheetName}' créée/réinitialisée avec les valeurs par défaut (incluant Phase5 styles/cols).`);
  SpreadsheetApp.getUi().alert("Configuration", `La feuille '${configSheetName}' a été (ré)initialisée avec les valeurs par défaut.`);
  return JSON.parse(JSON.stringify(CONFIG));
}

/** Met à jour un paramètre de configuration dans la feuille _CONFIG (INCHANGÉE) */
function updateConfig(param, value) {
  // ... (Votre code existant ici, il fonctionne pour les clés simples)
  // Pour mettre à jour des clés complexes (STYLE_OPTION_FORMATS par ex.),
  // 'value' devrait être une chaîne JSON valide.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Tente de lire la config actuelle pour le nom, sinon utilise la constante, puis fallback
  let currentConfigForSheetName;
  try { currentConfigForSheetName = getConfig(); } catch(e) { currentConfigForSheetName = CONFIG; }
  const configSheetName = (currentConfigForSheetName.SHEETS?.CONFIG || CONFIG.SHEETS?.CONFIG || "_CONFIG");
  const configSheet = ss.getSheetByName(configSheetName);

  if (!configSheet) {
    Logger.log(`Feuille '${configSheetName}' non trouvée pour updateConfig. Exécutez createDefaultConfig() d'abord.`);
    SpreadsheetApp.getUi().alert("Erreur", `Feuille de configuration '${configSheetName}' non trouvée.`);
    return;
  }

  const data = configSheet.getRange("A:A").getValues().flat(); // Recherche en colonne A
  let rowIndex = -1;
  const paramTrimmedUpper = String(param).trim().toUpperCase();

  // Chercher l'en-tête pour déterminer la première ligne de données
  let headerRowIndex = -1;
  const rawSheetData = configSheet.getDataRange().getValues();
   for (let i=0; i < Math.min(3, rawSheetData.length); i++) { 
      const potentialHeaders = rawSheetData[i].map(h => String(h || '').trim().toUpperCase());
      if (potentialHeaders.includes("PARAMETRE")) {
          headerRowIndex = i;
          break;
      }
  }
  const startDataRowSearch = headerRowIndex === -1 ? 0 : headerRowIndex + 1;


  for (let i = startDataRowSearch; i < data.length; i++) { // Commence après l'en-tête potentiel
    if (String(data[i] || '').trim().toUpperCase() === paramTrimmedUpper) {
      rowIndex = i + 1; // +1 car getValues est 0-basé, getRange est 1-basé
      break;
    }
  }

  let valueToWrite = value;
  if (Array.isArray(value)) {
      valueToWrite = value.join(',');
  } else if (typeof value === 'object' && value !== null) {
      try { valueToWrite = JSON.stringify(value); }
      catch (e) { Logger.log(`Erreur stringify pour ${param}: ${e}`); /* garde la valeur originale */ }
  }


  if (rowIndex > 0) {
    Logger.log(`Mise à jour du paramètre '${param}' (ligne ${rowIndex}) avec la valeur '${valueToWrite}' dans '${configSheetName}'`);
    configSheet.getRange(rowIndex, 2).setValue(valueToWrite); // Colonne B pour VALEUR
  } else {
    const newRow = configSheet.getLastRow() + 1;
    Logger.log(`Ajout du nouveau paramètre '${param}' (ligne ${newRow}) avec la valeur '${valueToWrite}' dans '${configSheetName}'`);
    configSheet.getRange(newRow, 1, 1, 2).setValues([[param, valueToWrite]]);
  }
   SpreadsheetApp.getActiveSpreadsheet().toast(`Config: '${param}' mis à jour.`, "Configuration");
}

// --- FIN DE Config.gs ---