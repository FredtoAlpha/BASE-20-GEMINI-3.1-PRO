/**
 * ===================================================================
 * SYSTÈME BASEOPTI - Pool centralisé pour optimisation incrémentale
 * ===================================================================
 *
 * Architecture :
 * - Crée _BASEOPTI depuis le mode choisi (TEST/CACHE/FIN)
 * - Phases 1→3 piochent exclusivement dans _BASEOPTI
 * - Tracking : _PLACED ("", "P1", "P2", "P3")
 * - Phase 4 travaille sur …CACHE avec mini-gardien
 */

// ===================================================================
// UTILITAIRES DE BASE
// ===================================================================

// logLine : définition canonique dans App.Core.js (pas de fallback local)
// getOrCreateSheetByName_ → supprimée (doublon de getOrCreateSheet_ dans App.SheetsData.js)

/**
 * Vide un onglet et écrit les en-têtes
 */
function clearAndSetHeader_(sh, headers) {
  sh.clearContents().clearFormats();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formater l'en-tête
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4A90E2')
    .setFontColor('#FFFFFF');

  return sh;
}

/**
 * Lit un onglet et retourne un tableau d'objets
 * ⚠️ FILTRE : Ne garde que les lignes avec un ID valide en colonne A
 */
function readRowsAsObjects_(sh) {
  const rng = sh.getDataRange();
  const values = rng.getValues();
  if (values.length < 2) return [];

  const head = values[0];
  const results = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    // ⚠️ IMPÉRATIF : Ignorer les lignes sans ID en colonne A (lignes de stats)
    const idColA = row[0];
    if (!idColA || String(idColA).trim() === '') continue;

    const o = {};
    for (let i = 0; i < head.length; i++) {
      o[head[i]] = row[i];
    }
    results.push(o);
  }

  return results;
}

/**
 * Écrit un tableau d'objets dans un onglet (à partir de la ligne 2)
 */
function writeObjects_(sh, headers, rows) {
  if (!rows || rows.length === 0) return;

  const out = rows.map(function(r) {
    return headers.map(function(h) {
      return r[h] !== undefined ? r[h] : '';
    });
  });

  sh.getRange(2, 1, out.length, headers.length).setValues(out);
}

// ===================================================================
// CONSTRUCTION DE _BASEOPTI
// ===================================================================

/**
 * ============================================================
 *  BASEOPTI — pool unique d'élèves pour le pipeline V2 (UI)
 * ============================================================
 * Crée un onglet `_BASEOPTI` (caché) en fusionnant les élèves
 * des onglets source choisis (ctx.srcSheets) avec un SCHÉMA FIXE
 * garantissant la présence de toutes les colonnes essentielles.
 */

// ✅ SCHÉMA FIXE _BASEOPTI (ordre standardisé)
const BASE_SCHEMA = [
  "ID_ELEVE", "NOM", "PRENOM", "NOM & PRENOM", "SEXE", "LV2", "OPT",
  "COM", "TRA", "PART", "ABS", "DISPO", "ASSO", "DISSO",
  "SOURCE", "FIXE", "CLASSE_FINAL", "CLASSE_DEF", 
  "",  // ← COLONNE VIDE (colonne S) pour alignement avec schéma legacy
  "MOBILITE",  // ← Maintenant en colonne T (au lieu de S)
  "SCORE F", "SCORE M", "GROUP", 
  "_ID", "_PLACED", "_TARGET_CLASS"  // ✅ Colonnes legacy pour compatibilité
];

// ✅ COUCHE DE COMPATIBILITÉ : Alias pour ancien schéma
const LEGACY_ALIASES = {
  // ID (clé primaire)
  "ID": ["ID_ELEVE", "ID", "_ID"],
  "ID_ELEVE": ["ID_ELEVE", "ID", "_ID"],
  "_ID": ["_ID", "ID_ELEVE", "ID"],
  
  // Classe finale / déf
  "CLASSE_FINAL": ["CLASSE_FINAL", "CLASSE FINAL", "LASSE_FINAL", "CLASSE", "_TARGET_CLASS"],
  "CLASSE DEF": ["CLASSE_DEF", "CLASSE DEF"],
  "CLASSE_DEF": ["CLASSE_DEF", "CLASSE DEF"],
  
  // Scores (essentiels pour P4)
  "COM": ["COM"],
  "TRA": ["TRA"],
  "PART": ["PART"],
  "ABS": ["ABS"],
  
  // Groupes ASSO/DISSO
  "A": ["ASSO", "A", "CODE_A"],
  "ASSO": ["ASSO", "A", "CODE_A"],
  "D": ["DISSO", "D", "CODE_D"],
  "DISSO": ["DISSO", "D", "CODE_D"],
  
  // Divers
  "SOURCE": ["SOURCE", "_SOURCE_CLASS", "_SOURCE_CLA"],
  "_SOURCE_CLASS": ["SOURCE", "_SOURCE_CLASS", "_SOURCE_CLA"],
  "CLASSE_FINAL": ["CLASSE_FINAL", "CLASSE FINAL", "LASSE_FINAL", "CLASSE", "_TARGET_CLASS", "_TARGET_CLA"],
  "_TARGET_CLASS": ["_TARGET_CLASS", "_TARGET_CLA", "CLASSE_FINAL", "CLASSE FINAL"],
  "_PLACED": ["_PLACED", "PLACED", "MOBILITE"],
  "SEXE": ["SEXE", "Sexe", "Genre", "GENRE"],
  "NOM": ["NOM", "Nom"],
  "PRENOM": ["PRENOM", "Prenom", "Prénom"]
};

/**
 * Résout un nom de colonne logique vers le nom physique présent dans les en-têtes
 * @param {string} logicalName - Nom logique de la colonne
 * @param {Array} headers - En-têtes physiques de la feuille
 * @returns {Object|null} {name: string, idx: number} ou null si non trouvé
 */
function resolveHeader_(logicalName, headers) {
  const candidates = LEGACY_ALIASES[logicalName] || [logicalName];
  for (const name of candidates) {
    const idx = headers.indexOf(name);
    if (idx !== -1) {
      return { name: name, idx: idx };
    }
  }
  return null;
}

/**
 * Getters robustes pour accès aux données (compatibilité ancien/nouveau schéma)
 */

function getId_(row, headers) {
  const h = resolveHeader_("ID_ELEVE", headers) || resolveHeader_("ID", headers) || resolveHeader_("_ID", headers);
  return h ? String(row[h.idx] || "").trim() : "";
}

function getScore_(row, headers, scoreKey) {
  const h = resolveHeader_(scoreKey, headers);
  return h ? Number(row[h.idx] || 0) : 0;
}

function getClasseFinal_(row, headers) {
  const h = resolveHeader_("CLASSE_FINAL", headers);
  return h ? String(row[h.idx] || "").trim() : "";
}

function getPlaced_(row, headers) {
  const h = resolveHeader_("_PLACED", headers);
  return h ? String(row[h.idx] || "").trim() : "";
}

function getAsso_(row, headers) {
  const h = resolveHeader_("ASSO", headers);
  return h ? String(row[h.idx] || "").trim() : "";
}

function getDisso_(row, headers) {
  const h = resolveHeader_("DISSO", headers);
  return h ? String(row[h.idx] || "").trim() : "";
}

/**
 * Extrait un ID stable depuis un objet (fallback sur variantes)
 */
function pickStableId_(obj) {
  return String(obj.ID_ELEVE || obj.ID || obj._ID || "").trim();
}

/**
 * Mapping d'une ligne source vers le schéma _BASEOPTI
 */
function mapWorkRowToBaseOpti_(work, srcName, rowIdx) {
  // ✅ Backfill des scores depuis les colonnes sources
  const scores = backfillScores_(work);
  
  // ✅ ID stable (idempotent : conserve _ID existant)
  const stableId = ensureStableId_(work, srcName, rowIdx);
  
  return {
    "ID_ELEVE": work.ID_ELEVE || work.ID || stableId,
    "NOM": work.NOM || "",
    "PRENOM": work.PRENOM || "",
    "NOM & PRENOM": work["NOM & PRENOM"] || (work.NOM + " " + work.PRENOM).trim(),
    "SEXE": work.SEXE || work.Sexe || work.Genre || "",
    "LV2": work.LV2 || "",
    "OPT": work.OPT || "",
    "COM": scores.COM,
    "TRA": scores.TRA,
    "PART": scores.PART,
    "ABS": scores.ABS,
    "DISPO": work.DISPO || "",
    "ASSO": work.ASSO || work.A || work.CODE_A || "",
    "DISSO": work.DISSO || work.D || work.CODE_D || "",
    "SOURCE": work.SOURCE || srcName.replace(/(TEST|CACHE|FIN)$/,'').replace(/'+/g,'').trim(),
    "FIXE": "",  // ✅ MODIF : Ne plus copier, sera calculé par computeMobilityFlags_()
    "CLASSE_FINAL": work.CLASSE_FINAL || work.CLASSE || work._TARGET_CLASS || "",
    "CLASSE_DEF": work.CLASSE_DEF || work["CLASSE DEF"] || "",
    "": "",  // ✅ COLONNE VIDE (colonne S) pour alignement avec BASE_SCHEMA
    "MOBILITE": "",  // ✅ MODIF : Ne plus copier, sera calculé par computeMobilityFlags_()
    "SCORE F": work["SCORE F"] || work.SCORE_F || "",
    "SCORE M": work["SCORE M"] || work.SCORE_M || "",
    "GROUP": work.GROUP || work.A || work.D || "",
    "_ID": work._ID || work.ID_ELEVE || work.ID || buildStableId_(work, srcName, rowIdx),
    "_PLACED": work._PLACED || "",
    "_TARGET_CLASS": work._TARGET_CLASS || work.CLASSE_FINAL || work.CLASSE || ""
  };
}

/**
 * Crée _BASEOPTI avec schéma fixe
 */
function createBaseOpti_(ctx) {
  const sh = getBaseOptiSheet_();
  sh.clear();

  // Écrire les en-têtes (schéma fixe)
  sh.getRange(1, 1, 1, BASE_SCHEMA.length).setValues([BASE_SCHEMA]);

  const allRows = [];
  
  // Lire toutes les sources
  (ctx.srcSheets || []).forEach(function(srcName) {
    const src = getSheetByNameSafe_(srcName);
    if (!src) return;
    const values = src.getDataRange().getValues();
    if (values.length < 2) return;
    const head = values[0].map(String);
    const h = indexer_(head);
    
    // Lire les élèves
    for (let i=1; i<values.length; i++) {
      const r = values[i];
      // Filtrage lignes vides
      if (!r[h["NOM"]] && !r[h["PRENOM"]]) continue;
      
      // Créer l'objet élève depuis la ligne source
      const work = {};
      head.forEach(function(col, idx) {
        if (col && String(col).trim() !== '') {
          work[String(col).trim()] = r[idx];
        }
      });
      
      // Mapper vers le schéma _BASEOPTI
      const mapped = mapWorkRowToBaseOpti_(work, srcName, i+1);
      allRows.push(mapped);
    }
  });
  
  // Écrire les données
  if (allRows.length > 0) {
    const rows = allRows.map(function(obj) {
      return BASE_SCHEMA.map(function(col) {
        return obj[col] !== undefined ? obj[col] : "";
      });
    });
    sh.getRange(2, 1, rows.length, BASE_SCHEMA.length).setValues(rows);
  }
  
  // Cache la feuille pour ne pas gêner l'utilisateur
  try { sh.hideSheet(); } catch(e) {}

  logLine('INFO', '✅ _BASEOPTI créé : ' + allRows.length + ' élèves, ' + BASE_SCHEMA.length + ' colonnes (schéma fixe)');

  // ✅ Audit post-création
  auditBaseoptiPostCreation_();

  return { ok:true, total: allRows.length };
}

// ===================================================================
// SÉLECTEURS BASEOPTI
// ===================================================================

/**
 * Retourne tous les élèves non encore placés (_PLACED="")
 */
function baseGetFree_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('_BASEOPTI');
  if (!sh) {
    throw new Error('_BASEOPTI introuvable');
  }

  const all = readRowsAsObjects_(sh);
  return all.filter(function(r) {
    return !r._PLACED || r._PLACED === '';
  });
}

/**
 * Retourne tous les élèves déjà placés (_PLACED non vide)
 */
function baseGetPlaced_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('_BASEOPTI');
  if (!sh) {
    throw new Error('_BASEOPTI introuvable');
  }

  const all = readRowsAsObjects_(sh);
  return all.filter(function(r) {
    return r._PLACED && r._PLACED !== '';
  });
}

/**
 * Marque des élèves comme placés
 *
 * @param {Array<string>} ids - Liste des _ID à marquer
 * @param {string} phase - "P1", "P2", "P3"
 * @param {string} targetClass - Classe de destination (ex: "6°1")
 */
function baseMarkPlaced_(ids, phase, targetClass) {
  const sh = SpreadsheetApp.getActive().getSheetByName('_BASEOPTI');
  if (!sh) {
    throw new Error('_BASEOPTI introuvable');
  }

  const rng = sh.getDataRange();
  const values = rng.getValues();
  const head = values[0];

  // ✅ Utiliser resolveHeader_ pour compatibilité
  const hId = resolveHeader_("ID_ELEVE", head) || resolveHeader_("_ID", head);
  const hPlaced = resolveHeader_("_PLACED", head);
  const hTarget = resolveHeader_("CLASSE_FINAL", head) || resolveHeader_("_TARGET_CLASS", head);

  if (!hId || !hPlaced || !hTarget) {
    logLine('ERROR', '❌ Colonnes introuvables dans _BASEOPTI');
    logLine('ERROR', '   En-têtes disponibles: ' + head.join(', '));
    throw new Error('Colonnes ID/PLACED/TARGET introuvables dans _BASEOPTI');
  }

  const idxId = hId.idx;
  const idxPlaced = hPlaced.idx;
  const idxTarget = hTarget.idx;
  
  // ✅ Trouver aussi _TARGET_CLASS pour compatibilité
  const hTargetLegacy = resolveHeader_("_TARGET_CLASS", head);
  const idxTargetLegacy = hTargetLegacy ? hTargetLegacy.idx : -1;

  const set = {};
  ids.forEach(function(id) { set[id] = true; });
  const ts = new Date();
  let marked = 0;

  for (let r = 1; r < values.length; r++) {
    const rowId = String(values[r][idxId] || '').trim();
    if (set[rowId]) {
      values[r][idxPlaced] = phase;
      values[r][idxTarget] = targetClass || values[r][idxTarget] || '';
      
      // ✅ Écrire aussi dans _TARGET_CLASS si présente
      if (idxTargetLegacy >= 0) {
        values[r][idxTargetLegacy] = targetClass || values[r][idxTargetLegacy] || '';
      }
      
      marked++;
    }
  }

  rng.setValues(values);
  logLine('INFO', '  ✅ ' + marked + ' élèves marqués ' + phase + ' → ' + targetClass);
}

// ===================================================================
// HELPERS POUR ÉCRITURE DANS CACHE
// ===================================================================

/**
 * Écrit un groupe d'élèves dans un onglet CACHE avec UPSERT par ID_ELEVE
 * Élimine les doublons en remplaçant les lignes existantes
 *
 * @param {Object} ctx - Contexte
 * @param {string} targetClass - Classe cible (ex: "6°1")
 * @param {Array} students - Tableau d'objets élèves
 */
function writeBatchToCache_(ctx, targetClass, students) {
  if (!students || students.length === 0) return;

  const ss = ctx.ss || SpreadsheetApp.getActive();
  const cacheName = targetClass + 'CACHE';
  const sh = ss.getSheetByName(cacheName);

  if (!sh) {
    logLine('WARN', '⚠️ Onglet ' + cacheName + ' introuvable');
    return;
  }

  // ✅ CORRECTIF : Si le CACHE est vide, créer les en-têtes depuis le premier élève
  const lastRow = sh.getLastRow();
  let headers;
  
  if (lastRow === 0 || sh.getLastColumn() === 0) {
    // CACHE vide : créer les en-têtes depuis le premier élève
    headers = Object.keys(students[0]);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    logLine('INFO', '  📝 ' + cacheName + ' : En-têtes créées (' + headers.length + ' colonnes)');
  } else {
    // CACHE existant : lire les en-têtes
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  }
  
  // Trouver l'index de la colonne ID_ELEVE (clé primaire)
  let idColIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    const hNorm = String(headers[i]).trim();
    if (hNorm === 'ID_ELEVE' || hNorm === 'ID' || hNorm === '_ID') {
      idColIdx = i;
      break;
    }
  }
  
  if (idColIdx === -1) {
    logLine('ERROR', '❌ Colonne ID_ELEVE introuvable dans ' + cacheName + ' (colonnes: ' + headers.join(', ') + ')');
    return;
  }

  // Lire les données existantes pour construire un index ID → ligne
  const existingData = sh.getDataRange().getValues();
  const idToRowMap = {}; // ID_ELEVE → numéro de ligne (1-indexed)
  
  for (let i = 1; i < existingData.length; i++) {
    const id = String(existingData[i][idColIdx] || '').trim();
    if (id) {
      idToRowMap[id] = i + 1; // +1 car getRange est 1-indexed
    }
  }

  // Construire les nouvelles lignes
  const newRows = [];
  const updateRows = []; // {row: numéro, data: [...]}
  
  students.forEach(function(stu) {
    const rowData = headers.map(function(h) {
      const hUpper = String(h).toUpperCase().trim();

      // Colonnes spéciales
      if (hUpper === 'CLASSE' || hUpper === 'CLASSE_FINALE' || hUpper === 'CLASSE DEF') {
        return targetClass;
      }
      if (h === 'ID_ELEVE' || h === 'ID' || h === '_ID') {
        return stu.ID_ELEVE || stu.ID || stu._ID || '';
      }
      if (h === 'ASSO' || h === 'A') {
        return stu.ASSO || stu.A || stu.CODE_A || '';
      }
      if (h === 'DISSO' || h === 'D') {
        return stu.DISSO || stu.D || stu.CODE_D || '';
      }

      // Colonnes standards (mapping direct)
      if (stu[h] !== undefined) return stu[h];

      // Fallback : vide
      return '';
    });
    
    const stuId = pickStableId_(stu);
    if (!stuId) {
      logLine('WARN', '⚠️ Élève sans ID ignoré dans ' + cacheName);
      return;
    }
    
    // UPSERT : si l'ID existe déjà, on met à jour, sinon on ajoute
    if (idToRowMap[stuId]) {
      updateRows.push({ row: idToRowMap[stuId], data: rowData });
    } else {
      newRows.push(rowData);
    }
  });

  // Appliquer les mises à jour
  updateRows.forEach(function(upd) {
    sh.getRange(upd.row, 1, 1, headers.length).setValues([upd.data]);
  });

  // Ajouter les nouvelles lignes
  if (newRows.length > 0) {
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  // Vérification post-écriture (utiliser pickStableId_)
  const totalWritten = updateRows.length + newRows.length;
  const uniqueIds = new Set(students.map(pickStableId_).filter(Boolean));
  
  if (uniqueIds.size !== totalWritten) {
    logLine('WARN', '⚠️ ' + cacheName + ' : ' + totalWritten + ' lignes écrites mais ' + uniqueIds.size + ' IDs uniques');
  }

  logLine('INFO', '  ✅ ' + cacheName + ' : ' + updateRows.length + ' màj + ' + newRows.length + ' ajouts (total=' + totalWritten + ')');
}

/**
 * Récupère les IDs d'un groupe d'élèves
 */
function grpsIds_(arr) {
  return arr.map(function(s) { return s._ID; });
}

// ===================================================================
// AUDIT & STATS
// ===================================================================

/**
 * Compte le nombre d'élèves dans chaque classe CACHE
 */
function getCacheStats_(ctx) {
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const stats = {};

  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const sh = ss.getSheetByName(cacheName);
    if (!sh) return;

    const classe = cacheName.replace('CACHE', '');
    const count = sh.getLastRow() - 1; // -1 pour l'en-tête

    stats[classe] = count;
  });

  return stats;
}

/**
 * Calcule, pour chaque classe cible, l'effectif courant (dans ...CACHE),
 * la cible (targets), le besoin (target - current) et l'écart de parité.
 * Retourne un objet :
 *   { "6°1": { current: 24, target: 25, need: 1, F:12, M:12, parityDelta:0 }, ... }
 */
function getClassNeedsFromCache_(ctx) {
  const res = {};
  const targets = resolveTargets_(ctx); // hiérarchie : STRUCTURE → _OPTI_CONFIG → fallback 25
  const offers = (ctx.offersByClass || ctx.offers || {}); // info LV2/OPT par classe (si utile ailleurs)

  // Compter dans les feuilles CACHE
  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const clazz = cacheName.replace(/CACHE$/,'').replace(/'+/g,'').trim();
    const sh = getSheetByNameSafe_(cacheName);
    if (!sh) return;
    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      // ✅ CORRECTION : Même si vide, définir target et need !
      const target = Number(targets[clazz] || 25);
      res[clazz] = {
        current: 0,
        target: target,
        need: target,  // besoin = target - 0 = target
        F: 0,
        M: 0,
        parityDelta: 0,
        offers: offers[clazz] || { LV2:[], OPT:[] }
      };
      return;
    }
    const head = values[0].map(String);
    const h = {};
    head.forEach(function(k,i) { h[k]=i; });
    let F=0, M=0, current=0;
    for (let i=1;i<values.length;i++) {
      const r = values[i];
      const nom = String(r[h["NOM"]]||"").trim();
      const prenom = String(r[h["PRENOM"]]||"").trim();
      if (!nom && !prenom) continue;
      current++;
      const sexe = String(r[h["SEXE"]]||"").toUpperCase();
      if (sexe==="F") F++; else if (sexe==="M") M++;
    }
    const target = Number(targets[clazz] || 25);
    res[clazz] = {
      current: current, target: target,
      need: Math.max(0, target - current),
      F: F, M: M,
      parityDelta: Math.abs(F - M),
      offers: offers[clazz] || { LV2:[], OPT:[] }
    };
  });
  return res;
}

function resolveTargets_(ctx) {
  // ✅ MICRO-PATCH 1 : Hiérarchie renforcée pour éliminer les undefined
  // 1) override V2 (_OPTI_CONFIG)
  // 2) STRUCTURE capacity
  // 3) Fallback égalitaire si nécessaire
  
  const out = {};
  const rules = getStructureRules(); // {_class:{capacity, quotas}}
  const override = (ctx && ctx.targets) || {};
  const classes = ctx.levels || ctx.niveaux || [];

  // 1) Override V2 (_OPTI_CONFIG)
  classes.forEach(function(c) {
    if (Number.isFinite(override[c])) {
      out[c] = override[c];
    }
  });

  // 2) STRUCTURE capacity
  classes.forEach(function(c) {
    if (!Number.isFinite(out[c]) && rules[c] && Number.isFinite(rules[c].capacity)) {
      out[c] = rules[c].capacity;
    }
  });

  // 3) Fallback égalitaire si nécessaire
  const needFallback = classes.filter(function(c) { return !Number.isFinite(out[c]); }).length > 0;
  if (needFallback) {
    const baseCount = countStudentsFromBaseopti_(); // total élèves _BASEOPTI
    const per = Math.ceil(baseCount / (classes.length || 1));
    classes.forEach(function(c) {
      if (!Number.isFinite(out[c])) {
        out[c] = per;
      }
    });
  }

  return out;
}

/**
 * Calcule si on a besoin de plus de F ou M dans une classe (pour parité)
 */
function needParityF_(ctx, targetClass) {
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const cacheName = targetClass + 'CACHE';
  const sh = ss.getSheetByName(cacheName);

  if (!sh) return true; // Par défaut, alterner

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return true;

  const headers = data[0];
  const idxSexe = headers.indexOf('SEXE');
  if (idxSexe === -1) return true;

  let countF = 0, countM = 0;
  for (let i = 1; i < data.length; i++) {
    const sexe = String(data[i][idxSexe] || '').toUpperCase();
    if (sexe === 'F') countF++;
    else if (sexe === 'M') countM++;
  }

  // Retourner true si on a besoin de plus de F
  return countF <= countM;
}

// ===================================================================
// UTILITAIRES BASEOPTI (PATCH V2)
// ===================================================================

function getBaseOptiSheet_() {
  const ss = SpreadsheetApp.getActive();
  const name = "_BASEOPTI";
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function readBaseOpti_(filters) {
  const sh = getBaseOptiSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(String);
  const h = indexer_(head);
  const out = [];
  for (let i=1;i<values.length;i++) {
    const r = values[i];
    const obj = {};
    head.forEach(function(k,idx) { obj[k] = r[idx]; });
    if (filters) {
      let pass = true;
      for (const k in filters) {
        if ((obj[k]||"") != filters[k]) { pass=false; break; }
      }
      if (!pass) continue;
    }
    out.push(obj);
  }
  return out;
}

function upsertBaseOpti_(rows) {
  // rows: tableau d'objets avec clé _ID obligatoire
  const sh = getBaseOptiSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 1) throw new Error("_BASEOPTI headers missing");
  const head = values[0].map(String);
  const h = indexer_(head);
  const id2row = new Map();
  for (let i=1;i<values.length;i++) {
    const id = String(values[i][h["_ID"]]||"");
    if (id) id2row.set(id, i+1); // 1-based row index
  }
  rows.forEach(function(obj) {
    const id = String(obj._ID||"");
    if (!id) return;
    const arr = head.map(function(k) { return obj[k] !== undefined ? obj[k] : ""; });
    const r = id2row.get(id);
    if (r) {
      sh.getRange(r,1,1,head.length).setValues([arr]);
    } else {
      sh.appendRow(arr);
    }
  });
  return { ok:true, upserted: rows.length };
}

function normalizeStudentRow_(row, h) {
  const pick = function(k) { return (h[k] != null ? row[h[k]] : ""); };
  const sexo = String(pick("SEXE")||"").toUpperCase();
  return {
    NOM: String(pick("NOM")||"").trim(),
    PRENOM: String(pick("PRENOM")||pick("PRÉNOM")||"").trim(),
    SEXE: (sexo==="F"||sexo==="M") ? sexo : "",
    LV2: String(pick("LV2")||"").toUpperCase().trim(),   // ex: ITA / ESP / ALL / PT / ""
    OPT: String(pick("OPT")||"").toUpperCase().trim(),   // ex: CHAV / ""
    A:   String(pick("A")||pick("ASSO")||"").toUpperCase().trim(),     // code(s) groupe A
    D:   String(pick("D")||pick("DISSO")||"").toUpperCase().trim()      // code(s) D
  };
}

function buildStableId_(stu, srcName, rowIdx) {
  // ID stable : NOM|PRENOM|SEXE|SRC|ROW
  return [
    (stu.NOM||"").toUpperCase(),
    (stu.PRENOM||"").toUpperCase(),
    (stu.SEXE||""),
    srcName,
    rowIdx
  ].join("|");
}

function indexer_(headerArr) {
  const map = {};
  headerArr.forEach(function(k,i) { map[String(k)] = i; });
  return map;
}

function getSheetByNameSafe_(name) {
  if (!name) return null;
  try { return SpreadsheetApp.getActive().getSheetByName(name); } catch(e) { return null; }
}

// ===================================================================
// AUDIT DE COHÉRENCE "DUR" - Invariants globaux
// ===================================================================

/**
 * Lit tous les élèves depuis les onglets CACHE
 * @returns {Object} { "6°1": [...rows], "6°2": [...rows], ... }
 */
function readAllCache_(ctx) {
  const cache = {};
  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const clazz = cacheName.replace(/CACHE$/,'').replace(/'+/g,'').trim();
    const sh = getSheetByNameSafe_(cacheName);
    if (!sh) {
      cache[clazz] = [];
      return;
    }
    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      cache[clazz] = [];
      return;
    }
    const head = values[0].map(String);
    const h = indexer_(head);
    const rows = [];
    for (let i=1; i<values.length; i++) {
      const r = values[i];
      const nom = String(r[h["NOM"]]||"").trim();
      const prenom = String(r[h["PRENOM"]]||"").trim();
      if (!nom && !prenom) continue;
      const obj = {};
      head.forEach(function(k,idx) { obj[k] = r[idx]; });
      rows.push(obj);
    }
    cache[clazz] = rows;
  });
  return cache;
}

/**
 * Vérifie les invariants globaux (conservation, unicité, exhaustivité, cibles définies)
 * @param {Object} ctx - Contexte
 * @param {string} label - Label pour les logs (ex: "POST P3", "AUDIT")
 */
function _assertInvariants_(ctx, label) {
  const base = readBaseOpti_();           // [{_ID, _PLACED, ...}]
  const cache = readAllCache_(ctx);       // { "6°1":[...], ..., "6°5":[...] }

  const totalBase = base.length;
  const totalPlaced = base.filter(function(r) { return r._PLACED; }).length;
  
  let totalCache = 0;
  for (const cls in cache) {
    totalCache += (cache[cls] ? cache[cls].length : 0);
  }

  // 1. Conservation
  if (totalPlaced !== totalCache) {
    logLine('ERROR', '❌ ' + label + ' – Conservation brisée: placed=' + totalPlaced + ' vs cache=' + totalCache);
  }

  // 2. Unicité (utiliser pickStableId_ pour compatibilité)
  const idsCache = [];
  for (const cls in cache) {
    (cache[cls] || []).forEach(function(r) {
      const id = pickStableId_(r);
      if (id) idsCache.push(id);
    });
  }
  const idsSet = {};
  idsCache.forEach(function(id) { idsSet[id] = (idsSet[id] || 0) + 1; });
  const uniqueCount = Object.keys(idsSet).length;
  if (uniqueCount !== totalCache) {
    logLine('ERROR', '❌ ' + label + ' – Doublons dans CACHE (ids uniques=' + uniqueCount + ' / rows=' + totalCache + ')');
  }

  // 3. Exhaustivité
  const notPlaced = totalBase - totalPlaced;
  if (notPlaced < 0) {
    logLine('ERROR', '❌ ' + label + ' – Trop d\'élèves placés: base=' + totalBase + ', placed=' + totalPlaced);
  } else if (notPlaced > 0) {
    logLine('WARN', '⚠️ ' + label + ' – ' + notPlaced + ' élèves non placés en fin de phase');
  }

  // 4. Cibles définies pour TOUTES les classes
  const targets = resolveTargets_(ctx);
  ctx.levels.forEach(function(cls) {
    if (!Number.isFinite(targets[cls])) {
      logLine('ERROR', '❌ ' + label + ' – target undefined pour ' + cls);
    }
  });
}

/**
 * Audit strict par classe (cibles, parité, quotas)
 * @param {Object} ctx - Contexte
 * @param {string} label - Label pour les logs
 */
function _auditStrictByClass_(ctx, label) {
  const cache = readAllCache_(ctx);
  const rules = getStructureRules();     // capacities & quotas depuis _STRUCTURE
  const targets = resolveTargets_(ctx);  // hiérarchie: _OPTI_CONFIG → _STRUCTURE → fallback

  ctx.levels.forEach(function(cls) {
    const rows = cache[cls] || [];
    let F = 0, M = 0;
    rows.forEach(function(r) {
      const sexe = String(r.SEXE || '').toUpperCase();
      if (sexe === 'F') F++;
      else if (sexe === 'M') M++;
    });
    const target = targets[cls];

    if (!Number.isFinite(target)) {
      logLine('ERROR', '❌ ' + label + ' – ' + cls + ' target=undefined (rows=' + rows.length + ')');
    } else if (rows.length !== target) {
      logLine('WARN', '⚠️ ' + label + ' – ' + cls + ' ' + rows.length + '/' + target + ' (écart=' + (target - rows.length) + ')');
    }

    // Quotas OPT/LV2
    const q = (rules[cls] && rules[cls].quotas) || {};
    for (const key in q) {
      let realized = 0;
      rows.forEach(function(r) {
        const opt = String(r.OPT || '').toUpperCase();
        const lv2 = String(r.LV2 || '').toUpperCase();
        if (opt === key || lv2 === key) realized++;
      });
      if (realized !== q[key]) {
        logLine('WARN', '⚠️ ' + label + ' – ' + cls + ' quota ' + key + ': attendu=' + q[key] + ', réalisé=' + realized);
      }
    }
  });
}
