/**
 * ===================================================================
 * ORCHESTRATION V14I - MODE STREAMING PAR PHASES
 * ===================================================================
 * 
 * Permet d'afficher les onglets CACHE en "live" pendant l'optimisation
 * Chaque phase est un appel séparé qui flush et ouvre les onglets
 */

// ===================================================================
// UTILITAIRES STREAMING
// ===================================================================

/**
 * Petit utilitaire : force l'UI à refléter les écritures (flush + micro pause)
 */
function _flushUi_(ms) {
  SpreadsheetApp.flush();
  if (ms && ms > 0) {
    try { Utilities.sleep(ms); } catch(e) {}
  }
}

/**
 * Construit le contexte depuis les options (version stream courte)
 */
function buildCtx_(opts) {
  const ss = SpreadsheetApp.getActive();
  const mode = (opts && opts.sourceFamily) || (opts && opts.mode) || 'TEST';
  const classes = (opts && opts.classes) || [];
  
  // Utiliser makeCtxFromUI_ si disponible, sinon construire manuellement
  const ctx = makeCtxFromUI_({ sourceFamily: mode, targetFamily: 'CACHE' });
  
  // Override avec les classes spécifiques si fournies
  if (classes && classes.length > 0) {
    ctx.niveaux = classes;
    ctx.cacheSheets = classes.map(function(c) { return c + 'CACHE'; });
    ctx.sourceSheets = classes.map(function(c) { return c + mode; });
  }
  
  return ctx;
}

/**
 * Alias pour compatibilité avec l'ancien code
 */
function buildContext_(options) {
  return buildCtx_(options);
}

/**
 * Crée/replace les …CACHE avec en-têtes, vide les lignes élèves
 * Retourne { opened: ["6°1CACHE",...], active: "6°1CACHE" }
 */
/**
 * Initialise les onglets CACHE (vide uniquement les lignes de données, garde les en-têtes)
 * ⚠️ SÉCURITÉ : Ne vide PAS tout, seulement les lignes élèves (>= ligne 2)
 */
function initEmptyCacheTabs_(ctx) {
  logLine('INFO', '📋 Initialisation onglets CACHE (vidage doux)...');

  const opened = [];

  (ctx.cacheSheets || []).forEach(function(name) {
    let sh = ctx.ss.getSheetByName(name);
    if (!sh) {
      sh = ctx.ss.insertSheet(name);
    }

    // ✅ VIDAGE DOUX : Garde les en-têtes (ligne 1), vide seulement les données
    if (sh.getLastRow() > 1) {
      const numRows = sh.getLastRow() - 1; // Nombre de lignes de données
      const numCols = Math.max(1, sh.getLastColumn());
      sh.getRange(2, 1, numRows, numCols).clearContent();
      logLine('INFO', '  🧹 ' + name + ' : ' + numRows + ' lignes vidées (en-tête conservé)');
    } else {
      // Pas d'en-têtes → créer
      writeCacheHeaders_(ctx, sh, name);
      logLine('INFO', '  ✨ ' + name + ' : en-têtes créés');
    }

    // S'assurer que les en-têtes existent
    if (sh.getLastRow() === 0) {
      writeCacheHeaders_(ctx, sh, name);
    }

    opened.push(name);
  });

  SpreadsheetApp.flush();

  const active = opened[0] || null;
  logLine('INFO', '✅ Onglets CACHE initialisés : ' + opened.join(', '));

  return { opened: opened, active: active };
}

/**
 * Écrit les en-têtes dans un onglet CACHE
 */
function writeCacheHeaders_(ctx, targetSheet, cacheName) {
  // Trouver l'onglet source pour copier les en-têtes
  let srcName;

  // ✅ FIX: Mode LEGACY avec mapping inverse
  if (ctx.destToSourceMapping && ctx.writeTarget === 'TEST') {
    // En mode LEGACY, cacheName est comme "6°1TEST"
    // On doit trouver l'onglet source correspondant (ex: PREVERT°1)
    const destName = cacheName.replace(ctx.writeTarget, ''); // "6°1"
    srcName = ctx.destToSourceMapping[destName]; // Accès direct !
  } else if (ctx.sourceToDestMapping && ctx.writeTarget === 'TEST') {
    // Fallback ancien code (si destToSourceMapping n'existe pas)
    const destName = cacheName.replace(ctx.writeTarget, ''); // "6°1"

    // Chercher la source qui mappe vers cette destination
    for (const [source, dest] of Object.entries(ctx.sourceToDestMapping)) {
      if (dest === destName) {
        srcName = source;
        break;
      }
    }
  } else {
    // Mode normal
    srcName = cacheName.replace('CACHE', ctx.modeSrc || '');
  }

  const srcSheet = ctx.ss.getSheetByName(srcName);

  if (srcSheet && srcSheet.getLastRow() > 0) {
    const headers = srcSheet.getRange(1, 1, 1, srcSheet.getLastColumn()).getValues()[0];
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Formater l'en-tête
    targetSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#C6E0B4');

    logLine('INFO', '    📋 En-têtes copiés de ' + srcName + ' vers ' + cacheName);
  } else {
    logLine('WARN', '⚠️ Impossible de trouver l\'onglet source pour ' + cacheName + ' (cherché: ' + srcName + ')');

    // ✅ Créer des en-têtes par défaut
    const defaultHeaders = ['NOM', 'PRENOM', 'SEXE', 'LV2', 'OPT', 'COM', 'TRA', 'PART', 'ABS', 'CODE_ASSO', 'CODE_DISSO', '_CLASS_ASSIGNED'];
    targetSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    targetSheet.getRange(1, 1, 1, defaultHeaders.length)
      .setFontWeight('bold')
      .setBackground('#C6E0B4');
    logLine('INFO', '    ✨ En-têtes par défaut créés pour ' + cacheName);
  }
}

/**
 * Nettoie exclusivement les colonnes LV2/OPT dans …CACHE (lignes élèves)
 * Évite de traîner des LV2/OPT invalides copiées de TEST
 */
function clearLv2OptColumnsInCache_(ctx) {
  logLine('INFO', '🧹 Nettoyage colonnes LV2/OPT dans CACHE...');
  
  (ctx.cacheSheets || []).forEach(function(name) {
    const sh = ctx.ss.getSheetByName(name);
    if (!sh) return;
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    
    const headers = data[0];
    const lastRow = sh.getLastRow();
    
    // Trouver les colonnes LV2 et OPT
    const idxLV2 = headers.indexOf('LV2');
    const idxOPT = headers.indexOf('OPT');
    
    // Nettoyer LV2
    if (idxLV2 >= 0 && lastRow > 1) {
      sh.getRange(2, idxLV2 + 1, lastRow - 1, 1).clearContent();
    }
    
    // Nettoyer OPT
    if (idxOPT >= 0 && lastRow > 1) {
      sh.getRange(2, idxOPT + 1, lastRow - 1, 1).clearContent();
    }
  });
  
  logLine('INFO', '✅ Colonnes LV2/OPT nettoyées');
}

/**
 * Place le focus sur la première feuille CACHE pour que l'UI "voit" les changements
 */
function focusFirstCacheTab_(ctx) {
  const first = (ctx.cacheSheets && ctx.cacheSheets[0]) || null;
  if (!first) return;
  
  const sh = ctx.ss.getSheetByName(first);
  if (sh) {
    ctx.ss.setActiveSheet(sh);
    sh.getRange('A1').activate();
  }
}

// openCacheTabs_() → supprimée (définition canonique dans App.CacheManager.js)

/**
 * Empaquette les infos de contexte pour le front
 */
function packCtxInfo_(ctx) {
  return {
    modeSrc: ctx.modeSrc,
    writeTarget: ctx.writeTarget,
    niveaux: ctx.niveaux,
    cacheSheets: ctx.cacheSheets,
    quotas: ctx.quotas
  };
}

// ===================================================================
// GARDE-FOUS OFFRE & QUOTAS
// ===================================================================

/**
 * Trouve une classe qui offre 'key' (LV2/OPT) et a du quota restant
 */
function findClassWithQuota_(quotas, offer, kind, key) {
  const k = String(key).toUpperCase();
  for (const cls in offer) {
    if (!offer.hasOwnProperty(cls)) continue;
    const arr = kind === 'LV2' ? (offer[cls].LV2 || []) : (offer[cls].OPT || []);
    if (arr.indexOf(k) >= 0 && quotas[cls] && quotas[cls][k] > 0) {
      return cls;
    }
  }
  return null;
}

/**
 * Applique l'offre + quotas après Phase 1 : corrige toute affectation LV2/OPT hors offre
 */
function enforceOfferOnCache_(ctx) {
  logLine('INFO', '🔒 Enforcement offre/quotas sur CACHE...');
  
  const offer = buildOfferWithQuotas_(ctx);
  const quotas = JSON.parse(JSON.stringify(ctx.quotas || {}));
  
  let corrected = 0;
  let purged = 0;
  
  (ctx.cacheSheets || []).forEach(function(name) {
    const cls = name.replace(/CACHE$/, '').trim();
    const sh = ctx.ss.getSheetByName(name);
    if (!sh) return;
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    
    const headers = data[0];
    const idxLV2 = headers.indexOf('LV2');
    const idxOPT = headers.indexOf('OPT');
    const idxClasse = headers.indexOf('Classe') >= 0 ? headers.indexOf('Classe') : 
                      (headers.indexOf('CLASSE') >= 0 ? headers.indexOf('CLASSE') : -1);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const lv2 = idxLV2 >= 0 ? String(row[idxLV2] || '').trim().toUpperCase() : '';
      const opt = idxOPT >= 0 ? String(row[idxOPT] || '').trim().toUpperCase() : '';
      const currentClass = idxClasse >= 0 ? String(row[idxClasse] || cls) : cls;
      
      // LV2 non offerte ? tenter un relogement
      if (lv2 && lv2 !== 'ANG') {
        const offeredLV2 = (offer[currentClass] && offer[currentClass].LV2) || [];
        if (offeredLV2.indexOf(lv2) === -1) {
          const dest = findClassWithQuota_(quotas, offer, 'LV2', lv2);
          if (dest) {
            quotas[dest][lv2]--;
            if (idxClasse >= 0) row[idxClasse] = dest;
            corrected++;
            logLine('INFO', '  ✅ LV2=' + lv2 + ' relogé: ' + currentClass + ' → ' + dest);
          } else {
            if (idxLV2 >= 0) row[idxLV2] = '';
            purged++;
            logLine('WARN', '  ⚠️ LV2=' + lv2 + ' sans place disponible → ligne ' + (i + 1) + ' purgée');
          }
        }
      }
      
      // OPT non offerte ?
      if (opt) {
        const offeredOPT = (offer[currentClass] && offer[currentClass].OPT) || [];
        if (offeredOPT.indexOf(opt) === -1) {
          const dest = findClassWithQuota_(quotas, offer, 'OPT', opt);
          if (dest) {
            quotas[dest][opt]--;
            if (idxClasse >= 0) row[idxClasse] = dest;
            corrected++;
            logLine('INFO', '  ✅ OPT=' + opt + ' relogé: ' + currentClass + ' → ' + dest);
          } else {
            if (idxOPT >= 0) row[idxOPT] = '';
            purged++;
            logLine('WARN', '  ⚠️ OPT=' + opt + ' sans place disponible → ligne ' + (i + 1) + ' purgée');
          }
        }
      }
      
      data[i] = row;
    }
    
    sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  });
  
  SpreadsheetApp.flush();
  logLine('INFO', '✅ Enforcement terminé: ' + corrected + ' relogés, ' + purged + ' purgés');
}

// ===================================================================
// SNAPSHOT LIVE POUR VIEWER
// ===================================================================

/**
 * Helper interne pour construire un snapshot des classes CACHE
 */
function buildCacheSnapshot_(ctx) {
  const ss = ctx.ss || SpreadsheetApp.getActive();
  const names = ctx.cacheSheets || [];
  const out = { classes: {} };

  names.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) { 
      out.classes[name.replace(/CACHE$/,'')] = { total:0, F:0, M:0, LV2:{}, OPT:{}, rows:[] }; 
      return; 
    }

    const values = sh.getRange(1,1,lastRow,lastCol).getValues();
    const headers = values[0].map(function(h) { return String(h).trim().toUpperCase(); });
    const idx = function(h) { return headers.indexOf(h); };

    const iSEXE = idx('SEXE'), iLV2 = idx('LV2'), iOPT = idx('OPT');
    const iNOM = idx('NOM'), iPRENOM = idx('PRENOM');
    const iA = (idx('A')>-1 ? idx('A') : idx('ASSO'));
    const iD = (idx('D')>-1 ? idx('D') : idx('DISSO'));

    let F=0, M=0; 
    const LV2={}, OPT={}, rows=[];
    
    for (let r=1; r<values.length; r++) {
      const row = values[r];
      const sexe = (iSEXE>=0 ? String(row[iSEXE]||'') : '').toUpperCase();
      if (sexe==='F') F++; else if (sexe==='M') M++;
      
      const lv2 = (iLV2>=0 ? String(row[iLV2]||'') : '').toUpperCase();
      const opt = (iOPT>=0 ? String(row[iOPT]||'') : '').toUpperCase();
      if (lv2 && lv2 !== 'ANG') LV2[lv2]=(LV2[lv2]||0)+1;
      if (opt) OPT[opt]=(OPT[opt]||0)+1;

      // 10 premières lignes pour l'aperçu
      if (rows.length<10) {
        rows.push({
          nom: (iNOM>=0 ? row[iNOM] : ''),
          prenom: (iPRENOM>=0 ? row[iPRENOM] : ''),
          lv2: lv2,
          opt: opt,
          a: (iA>=0 ? row[iA] : ''),
          d: (iD>=0 ? row[iD] : '')
        });
      }
    }
    
    out.classes[name.replace(/CACHE$/,'')] = { total: F+M, F: F, M: M, LV2: LV2, OPT: OPT, rows: rows };
  });

  return out;
}

/**
 * Endpoint direct pour le snapshot (appelé par l'UI)
 */
function getCacheSnapshotStream() {
  const ctx = optStream_init();
  return buildCacheSnapshot_(ctx);
}

// ===================================================================
// STREAMING ENDPOINTS (appelés par l'UI)
// => Noms alignés avec OptimizationPanel_StreamingMinimal.html
// ===================================================================

/**
 * 0) Ouvrir/activer les onglets ...CACHE (vides si besoin) + créer _BASEOPTI
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG via optStream_init_V2()
 * ⚙️ Défensif : garantit un retour stable { ok, cache, base, logs }
 * 🔒 LOCKSERVICE : Empêche les doubles lancements
 */
function openCacheTabsStream() {
  // ✅ LOCKSERVICE : Acquérir le verrou
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      logLine('WARN', '⚠️ Optimisation déjà en cours - verrou actif');
      return {
        ok: false,
        error: 'Optimisation déjà en cours. Veuillez patienter.',
        locked: true
      };
    }
  } catch(e) {
    logLine('ERROR', '❌ Erreur acquisition verrou: ' + e);
    return { ok: false, error: 'Erreur de verrouillage: ' + e };
  }

  try {
    const ctx = optStream_init_V2();

  // 🔥 ÉTAPE 1 : VIDER complètement les onglets CACHE
  logLine('INFO', '🧹 VIDAGE COMPLET des onglets CACHE...');
  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const sh = (ctx.ss || SpreadsheetApp.getActive()).getSheetByName(cacheName);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
      logLine('INFO', '  ✅ ' + cacheName + ' vidé');
    }
  });

  // 🔥 ÉTAPE 2 : Créer _BASEOPTI depuis les sources
  logLine('INFO', '🎯 Création _BASEOPTI depuis ' + ctx.modeSrc + '...');
  let baseResult = { ok: false, totalEleves: 0, msg: '' };
  try {
    if (typeof createBaseOpti_ === 'function') {
      const r = createBaseOpti_(ctx);
      baseResult = {
        ok: !!(r && (r.ok === true || r.total >= 0)),
        totalEleves: (r && typeof r.total === 'number') ? r.total : (r && r.count) || 0,
        msg: (r && r.msg) || 'BASEOPTI construit'
      };
      logLine('INFO', '🧱 _BASEOPTI créé: ' + baseResult.totalEleves + ' élèves');
    } else {
      baseResult.msg = 'createBaseOpti_ absent';
      logLine('WARN', '⚠️ ' + baseResult.msg);
    }
  } catch(e) {
    baseResult = { ok: false, totalEleves: 0, msg: 'Erreur: ' + e };
    logLine('ERROR', '❌ ' + e);
  }

  // 🔥 ÉTAPE 3 : Ajouter colonne _CLASS_ASSIGNED dans _BASEOPTI
  logLine('INFO', '📋 Ajout colonne _CLASS_ASSIGNED dans _BASEOPTI...');
  try {
    const baseSheet = (ctx.ss || SpreadsheetApp.getActive()).getSheetByName('_BASEOPTI');
    if (baseSheet) {
      const data = baseSheet.getDataRange().getValues();
      const headers = data[0];
      let colIdx = headers.indexOf('_CLASS_ASSIGNED');

      if (colIdx === -1) {
        headers.push('_CLASS_ASSIGNED');
        baseSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        colIdx = headers.length - 1;
      }

      // Vider la colonne
      if (data.length > 1) {
        baseSheet.getRange(2, colIdx + 1, data.length - 1, 1).clearContent();
      }

      logLine('INFO', '  ✅ Colonne _CLASS_ASSIGNED prête');
    }
  } catch(e) {
    logLine('ERROR', '❌ Erreur ajout colonne: ' + e);
  }

  // 🔥 ÉTAPE 4 : Ouvrir les onglets CACHE pour affichage
  const opened = openCacheTabs_(ctx);

  // Forcer le rendu
  try {
    _flushUi_(200);
  } catch(e) {
    SpreadsheetApp.flush();
  }

    // ✅ AUDIT : Vérifier les invariants après INIT
    try {
      _assertInvariants_(ctx, 'POST INIT');
    } catch(e) {
      logLine('WARN', '⚠️ Audit INIT échoué: ' + e);
    }

    // ✅ Créer cacheInfo depuis le contexte
    const cacheInfo = {
      levels: ctx.levels || [],
      cacheSheets: ctx.cacheSheets || [],
      targets: ctx.targets || {},
      quotas: ctx.quotas || {}
    };

    return {
      ok: true,
      step: 'INIT',
      cache: cacheInfo,
      base: baseResult,
      opened: opened && opened.opened || [],
      active: opened && opened.active || null,
      stats: opened && opened.stats || null
    };
  } catch(e) {
    logLine('ERROR', '❌ Erreur openCacheTabsStream: ' + e);
    return { ok: false, error: String(e) };
  } finally {
    // ✅ LOCKSERVICE : Libérer le verrou
    try {
      lock.releaseLock();
      logLine('INFO', '🔓 Verrou libéré (INIT)');
    } catch(e) {
      logLine('WARN', '⚠️ Erreur libération verrou: ' + e);
    }
  }
}

/**
 * 1) Phase 1I — Dispatch Options & LV2 (VERSION BASEOPTI)
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG via optStream_init_V2()
 * 🔒 LOCKSERVICE : Empêche les doubles lancements
 */
function phase1Stream() {
  // ✅ LOCKSERVICE : Acquérir le verrou
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      logLine('WARN', '⚠️ Phase 1 déjà en cours - verrou actif');
      return { ok: false, error: 'Phase 1 déjà en cours', locked: true };
    }
  } catch(e) {
    return { ok: false, error: 'Erreur de verrouillage: ' + e };
  }

  try {
    const ctx = optStream_init_V2();

    // ✅ ARCH V3 : Utiliser la version qui lit depuis _BASEOPTI
    const p1 = (typeof Phase1I_dispatchOptionsLV2_BASEOPTI_V3 === 'function')
      ? Phase1I_dispatchOptionsLV2_BASEOPTI_V3(ctx)
      : Phase1I_dispatchOptionsLV2_BASEOPTI(ctx);

  // ✅ AUDIT : Vérifier les invariants après P1
  try {
    _assertInvariants_(ctx, 'POST P1');
  } catch(e) {
    logLine('WARN', '⚠️ Audit P1 échoué: ' + e);
  }

    // Activer le premier onglet CACHE pour effet visuel
    if (ctx.cacheSheets && ctx.cacheSheets.length > 0) {
      const sh = ctx.ss.getSheetByName(ctx.cacheSheets[0]);
      if (sh) ctx.ss.setActiveSheet(sh);
    }

    _flushUi_(250);
    return {
      ok: !!(p1 && p1.ok),
      counts: p1 && p1.counts || {},
      step: "P1"
    };
  } catch(e) {
    logLine('ERROR', '❌ Erreur phase1Stream: ' + e);
    return { ok: false, error: String(e) };
  } finally {
    try {
      lock.releaseLock();
      logLine('INFO', '🔓 Verrou libéré (P1)');
    } catch(e) {
      logLine('WARN', '⚠️ Erreur libération verrou: ' + e);
    }
  }
}

/**
 * 2) Phase 2I — Application DISSO / ASSO (VERSION BASEOPTI)
 */
/**
 * 2) Phase 2I — Codes DISSO/ASSO (VERSION BASEOPTI)
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG via optStream_init_V2()
 * 🔒 LOCKSERVICE : Empêche les doubles lancements
 */
function phase2Stream() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: 'Phase 2 déjà en cours', locked: true };
    }
  } catch(e) {
    return { ok: false, error: 'Erreur de verrouillage: ' + e };
  }

  try {
    const ctx = optStream_init_V2();

    // ✅ ARCH V3 : Utiliser la version qui lit depuis _BASEOPTI
    const p2 = (typeof Phase2I_applyDissoAsso_BASEOPTI_V3 === 'function')
      ? Phase2I_applyDissoAsso_BASEOPTI_V3(ctx)
      : Phase2I_applyDissoAsso_BASEOPTI(ctx);

  // ✅ AUDIT : Vérifier les invariants après P2
  try {
    _assertInvariants_(ctx, 'POST P2');
  } catch(e) {
    logLine('WARN', '⚠️ Audit P2 échoué: ' + e);
  }

  // Activer le premier onglet CACHE pour effet visuel
  if (ctx.cacheSheets && ctx.cacheSheets.length > 0) {
    const sh = ctx.ss.getSheetByName(ctx.cacheSheets[0]);
    if (sh) ctx.ss.setActiveSheet(sh);
  }

    _flushUi_(250);
    return {
      ok: true,
      disso: p2 && p2.disso || 0,
      asso:  p2 && p2.asso  || 0,
      step: "P2"
    };
  } catch(e) {
    return { ok: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ===================================================================
// PHASE 3 – SHIM & SÉLECTEUR ROBUSTE
// ===================================================================

/**
 * Sélecteur robuste : retourne la bonne implémentation de Phase 3
 * en testant plusieurs signatures possibles.
 * ⚠️ NE PAS inclure Phase3I_fillAndParity_ ici (c'est le shim, risque de récursion)
 */
function getPhase3Runner_() {
  const G = this; // global
  const candidates = [
    G.Phase3I_completeAndParity_,   // ✅ NOM RÉEL trouvé dans Orchestration_V14I.gs
    G.Phase3I_fillAndBalance_,      // variantes possibles
    G.fillAndParity_,
    G.equilibrerParite_,
    G.equilibrerParite,
    G.ParityAlgorithm && G.ParityAlgorithm.equilibrerParite,
    G.ParityAlgorithm && G.ParityAlgorithm.fillAndParity_
  ];
  for (const fn of candidates) {
    if (typeof fn === 'function') return fn;
  }
  throw new Error('Phase 3 introuvable : aucune implémentation détectée (equilibrerParite/fillAndParity).');
}

/**
 * Shim de compatibilité : si Phase3I_fillAndParity_ n'existe pas,
 * on crée un alias vers l'implémentation détectée.
 */
if (typeof Phase3I_fillAndParity_ !== 'function') {
  globalThis.Phase3I_fillAndParity_ = function(ctx) {
    const runner = getPhase3Runner_();
    // Certaines implémentations prennent (ctx), d'autres rien.
    try { return runner.length >= 1 ? runner(ctx) : runner(); }
    catch (e) { throw e; }
  };
}

/**
 * 3) Phase 3I — Compléter effectifs & équilibrer parité (VERSION BASEOPTI)
 * Endpoint streaming Phase 3 (appelé par l'UI sans paramètre)
 */
/**
 * 3) Phase 3I — Effectifs & Parité (VERSION BASEOPTI)
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG via optStream_init_V2()
 * 🔒 LOCKSERVICE : Empêche les doubles lancements
 */
function phase3Stream() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: 'Phase 3 déjà en cours', locked: true };
    }
  } catch(e) {
    return { ok: false, error: 'Erreur de verrouillage: ' + e };
  }

  try {
    const ctx = optStream_init_V2();
    const res = (typeof Phase3I_completeAndParity_BASEOPTI_V3 === 'function')
      ? Phase3I_completeAndParity_BASEOPTI_V3(ctx)
      : Phase3I_completeAndParity_BASEOPTI(ctx);

  // ✅ AUDIT : Vérifier les invariants après P3
  try {
    _assertInvariants_(ctx, 'POST P3');
    
    // ✅ MICRO-POINT 1 : Alerte si élèves non placés après P3
    const base = readBaseOpti_();
    const notPlaced = base.filter(function(r) { return !r._PLACED || r._PLACED === ''; }).length;
    if (notPlaced > 0) {
      logLine('WARN', '⚠️ POST P3 – ' + notPlaced + ' élèves non placés. Si P4 fait 0 swap, vérifier la logique P3.');
    }
  } catch(e) {
    logLine('WARN', '⚠️ Audit P3 échoué: ' + e);
  }

  _flushUi_(250);                        // forcer le rendu côté Sheets

  // Activer le premier onglet CACHE pour effet visuel
  if (ctx.cacheSheets && ctx.cacheSheets.length > 0) {
    const sh = ctx.ss.getSheetByName(ctx.cacheSheets[0]);
    if (sh) ctx.ss.setActiveSheet(sh);
  }

  // Audit rapide (optionnel mais pratique pour le "live")
  const audit = (typeof auditCacheAgainstStructure_ === 'function')
    ? auditCacheAgainstStructure_(ctx) : null;

    return {
      ok: true,
      step: 'P3',
      details: res && res.stats ? res.stats : null,
      audit: audit
    };
  } catch(e) {
    return { ok: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ===================================================================
// PHASE 4 – SHIM & SÉLECTEUR ROBUSTE
// ===================================================================

/**
 * Sélecteur robuste : retourne la bonne implémentation de Phase 4
 * en testant plusieurs signatures possibles.
 * ⚠️ NE PAS inclure Phase4_optimizeSwaps_ ici (c'est le shim, risque de récursion)
 */
function getPhase4Runner_() {
  const G = this; // global
  const candidates = [
    G.Phase4_balanceScoresSwaps_,   // ✅ NOM RÉEL trouvé dans Orchestration_V14I.gs:1415
    G.Phase4_optimize_,
    G.Phase4_swaps_,
    G.optimizeSwaps_,
    G.balanceSwaps_,
    G.OptimizationAlgorithm && G.OptimizationAlgorithm.optimize,
    G.OptimizationAlgorithm && G.OptimizationAlgorithm.balanceSwaps
  ];
  for (const fn of candidates) {
    if (typeof fn === 'function') return fn;
  }
  throw new Error('Phase 4 introuvable : aucune implémentation détectée (Phase4_balanceScoresSwaps_/optimizeSwaps).');
}

/**
 * Shim de compatibilité : si Phase4_optimizeSwaps_ n'existe pas,
 * on crée un alias vers l'implémentation détectée.
 */
if (typeof Phase4_optimizeSwaps_ !== 'function') {
  globalThis.Phase4_optimizeSwaps_ = function(ctx, options) {
    const runner = getPhase4Runner_();
    // La vraie fonction Phase4_balanceScoresSwaps_ prend seulement (ctx)
    // Les options (maxSwaps, priority) sont ignorées pour l'instant
    try { return runner(ctx); }
    catch (e) { throw e; }
  };
}

/**
 * 4) Phase 4 — Optimisation par swaps
 * Endpoint streaming Phase 4 (appelé par l'UI sans paramètre)
 */
/**
 * 4) Phase 4 — Swaps (optimisation)
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG via optStream_init_V2()
 * 🔒 LOCKSERVICE : Empêche les doubles lancements
 */
function phase4Stream() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: 'Phase 4 déjà en cours', locked: true };
    }
  } catch(e) {
    return { ok: false, error: 'Erreur de verrouillage: ' + e };
  }

  try {
    const ctx = optStream_init_V2();
    
    // ✅ GARDE-FOU : Valider le contexte avant P4
    if (!ctx) {
      logLine('ERROR', '❌ Contexte V2 undefined');
      return { ok: false, error: 'Contexte V2 undefined' };
    }
    if (!ctx.levels || ctx.levels.length === 0) {
      logLine('ERROR', '❌ ctx.levels manquant ou vide');
      return { ok: false, error: 'ctx.levels manquant' };
    }
    if (!ctx.cacheSheets || ctx.cacheSheets.length === 0) {
      logLine('ERROR', '❌ ctx.cacheSheets manquant ou vide');
      return { ok: false, error: 'ctx.cacheSheets manquant' };
    }
    
    // ✅ GARDE-FOU : Valider _BASEOPTI avant P4
    try {
      assertBaseoptiValid_();
    } catch(e) {
      logLine('ERROR', '❌ _BASEOPTI invalide: ' + e);
      return { ok: false, error: '_BASEOPTI invalide: ' + e };
    }
    
    let r = null;
  try {
    // ✅ ARCH V3 : Utiliser la version qui lit depuis _BASEOPTI
    if (typeof Phase4_balanceScoresSwaps_BASEOPTI_V3 === 'function') {
      r = Phase4_balanceScoresSwaps_BASEOPTI_V3(ctx);
      logLine('INFO', '✅ P4 V3 exécutée');
    } else if (typeof Phase4_balanceScoresSwaps_ === 'function') {
      r = Phase4_balanceScoresSwaps_(ctx);
      logLine('INFO', '✅ P4 exécutée (fallback)');
    } else {
      logLine('ERROR', '❌ Phase4 introuvable');
      r = { ok: false, swapsApplied: 0, error: 'Phase4 introuvable' };
    }
  } catch(e) {
    logLine('ERROR', '❌ Erreur P4: ' + e);
    r = { ok: false, swapsApplied: 0, error: String(e) };
  }

  // ✅ AUDIT : Vérifier les invariants après P4
  try {
    _assertInvariants_(ctx, 'POST P4');
    
    // ✅ MICRO-POINT 1 : Log unique si notPlaced > 0 et P4 a fait 0 swap
    const base = readBaseOpti_();
    const notPlaced = base.filter(function(r) { return !r._PLACED || r._PLACED === ''; }).length;
    const swapsApplied = (r && r.swapsApplied) || 0;
    
    if (notPlaced > 0 && swapsApplied === 0) {
      logLine('ERROR', '❌ POST P4 – ' + notPlaced + ' élèves non placés ET 0 swap appliqué → Problème en Phase 3 !');
    }
  } catch(e) {
    logLine('WARN', '⚠️ Audit P4 échoué: ' + e);
  }

  _flushUi_(300);                        // forcer le rendu côté Sheets

  // Activer le premier onglet CACHE pour effet visuel
  if (ctx.cacheSheets && ctx.cacheSheets.length > 0) {
    const sh = ctx.ss.getSheetByName(ctx.cacheSheets[0]);
    if (sh) ctx.ss.setActiveSheet(sh);
  }

  // Audit rapide (optionnel mais pratique pour le "live")
  const audit = (typeof auditCacheAgainstStructure_ === 'function')
    ? auditCacheAgainstStructure_(ctx) : null;

    // ✅ CONTRAT DE SORTIE : Garantir toutes les propriétés (même vides)
    return {
      ok: r && r.ok !== false,
      step: 'P4',
      swaps: (r && r.swapsApplied) || (r && r.total) || 0,
      details: r && r.stats ? r.stats : {},
      audit: audit || {},
      summary: r && r.summary ? r.summary : {},
      scores: r && r.scores ? r.scores : { byClass: {}, totals: {} },
      metrics: r && r.metrics ? r.metrics : { com: 0, tra: 0, part: 0, abs: 0 },
      parity: r && r.parity ? r.parity : { byClass: {}, outOfTol: [] },
      swapsLog: r && r.swapsLog ? r.swapsLog : [],
      warnings: r && r.warnings ? r.warnings : [],
      weights: ctx.weights || {},
      error: r && r.error ? r.error : null
    };
  } catch(e) {
    return { ok: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ===================================================================
// CALCUL DES MOYENNES DE SCORES PAR CLASSE
// ===================================================================

/**
 * Calcule les moyennes COM/TRA/PART/ABS par classe depuis CACHE
 * ✅ NOUVEAU : Audit des scores comportementaux
 */
function computeScoreAveragesByClass_(ctx) {
  logLine('INFO', '📊 Calcul moyennes scores par classe...');

  const scoresByClass = {};

  (ctx.cacheSheets || []).forEach(function(cacheName) {
    const cls = cacheName.replace(/CACHE$/, '').replace(/TEST$/, '').trim();
    const sh = (ctx.ss || SpreadsheetApp.getActive()).getSheetByName(cacheName);

    if (!sh || sh.getLastRow() < 2) {
      scoresByClass[cls] = { COM: 0, TRA: 0, PART: 0, ABS: 0, count: 0 };
      return;
    }

    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const idxCOM = headers.indexOf('COM');
    const idxTRA = headers.indexOf('TRA');
    const idxPART = headers.indexOf('PART');
    const idxABS = headers.indexOf('ABS');

    let sumCOM = 0, sumTRA = 0, sumPART = 0, sumABS = 0;
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip lignes vides

      count++;

      if (idxCOM >= 0) {
        const com = Number(row[idxCOM] || 0);
        sumCOM += com;
      }
      if (idxTRA >= 0) {
        const tra = Number(row[idxTRA] || 0);
        sumTRA += tra;
      }
      if (idxPART >= 0) {
        const part = Number(row[idxPART] || 0);
        sumPART += part;
      }
      if (idxABS >= 0) {
        const abs = Number(row[idxABS] || 0);
        sumABS += abs;
      }
    }

    scoresByClass[cls] = {
      COM: count > 0 ? (sumCOM / count).toFixed(2) : 0,
      TRA: count > 0 ? (sumTRA / count).toFixed(2) : 0,
      PART: count > 0 ? (sumPART / count).toFixed(2) : 0,
      ABS: count > 0 ? (sumABS / count).toFixed(2) : 0,
      count: count
    };
  });

  // Log résumé
  logLine('INFO', '📊 Moyennes scores par classe :');
  for (const cls in scoresByClass) {
    const s = scoresByClass[cls];
    logLine('INFO', '  ' + cls + ' : COM=' + s.COM + ', TRA=' + s.TRA + ', PART=' + s.PART + ', ABS=' + s.ABS + ' (' + s.count + ' élèves)');
  }

  return scoresByClass;
}

/**
 * ✅ FINALISATION COMPLÈTE : Applique la MÊME mise en forme que les onglets FIN
 * - Formatage scores (rouge/jaune/vert clair/vert foncé)
 * - Formatage LV2 (ITA=vert, ESP=orange, CHAV=violet)
 * - Moyennes scores avec formules AVERAGE
 * - Décomptes LV2 et OPT
 * - Noms en gras, colonnes cachées
 */
function finalizeTestSheets_(ctx) {
  logLine('INFO', '📊 Finalisation onglets TEST : mise en forme COMPLÈTE comme onglets FIN...');

  (ctx.cacheSheets || []).forEach(function(sheetName) {
    try {
      const sh = (ctx.ss || SpreadsheetApp.getActive()).getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;

      const data = sh.getDataRange().getValues();
      const headers = data[0];
      const numRows = data.length - 1; // Sans l'en-tête

      if (numRows === 0) return;

      // ===== INDICES DES COLONNES =====
      const idxNom = Math.max(headers.indexOf('NOM'), 0);
      const idxPrenom = Math.max(headers.indexOf('PRENOM'), headers.indexOf('PRÉNOM'), 1);
      const idxNomPrenom = headers.indexOf('NOM & PRENOM');
      const idxSexe = headers.indexOf('SEXE');
      const idxLV2 = headers.indexOf('LV2');
      const idxOPT = Math.max(headers.indexOf('OPT'), headers.indexOf('OPTION'));
      const idxCOM = headers.indexOf('COM');
      const idxTRA = headers.indexOf('TRA');
      const idxPART = headers.indexOf('PART');
      const idxABS = headers.indexOf('ABS');
      const idxASSO = Math.max(headers.indexOf('CODE_ASSO'), headers.indexOf('ASSO'), headers.indexOf('A'));
      const idxDISSO = Math.max(headers.indexOf('CODE_DISSO'), headers.indexOf('DISSO'), headers.indexOf('D'));

      // ========== 1. LARGEURS DES COLONNES ==========
      if (idxNomPrenom >= 0) sh.setColumnWidth(idxNomPrenom + 1, 240);
      if (idxSexe >= 0) sh.setColumnWidth(idxSexe + 1, 70);
      if (idxLV2 >= 0) sh.setColumnWidth(idxLV2 + 1, 90);
      if (idxOPT >= 0) sh.setColumnWidth(idxOPT + 1, 110);
      if (idxCOM >= 0) sh.setColumnWidth(idxCOM + 1, 75);
      if (idxTRA >= 0) sh.setColumnWidth(idxTRA + 1, 75);
      if (idxPART >= 0) sh.setColumnWidth(idxPART + 1, 75);
      if (idxABS >= 0) sh.setColumnWidth(idxABS + 1, 75);
      if (idxASSO >= 0) sh.setColumnWidth(idxASSO + 1, 80);
      if (idxDISSO >= 0) sh.setColumnWidth(idxDISSO + 1, 80);

      // ========== 2. EN-TÊTE (ligne 1) - Violet foncé + Blanc gras + Capitales ==========
      const headerRange = sh.getRange(1, 1, 1, headers.length);
      headerRange
        .setBackground('#5b21b6')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      // Mettre en capitales
      const headerValues = headers.map(function(h) { return String(h).toUpperCase(); });
      sh.getRange(1, 1, 1, headers.length).setValues([headerValues]);

      // ========== 3. FORMATAGE COLONNE NOM & PRENOM - Gras 14pt ==========
      if (idxNomPrenom >= 0 && numRows > 0) {
        const namesRange = sh.getRange(2, idxNomPrenom + 1, numRows, 1);
        namesRange.setFontWeight('bold').setFontSize(14);
      }

      // ========== 4. FORMATAGE COLONNE SEXE - F=rose, M=bleu ==========
      if (idxSexe >= 0 && numRows > 0) {
        for (let r = 1; r <= numRows; r++) {
          const sexeValue = String(data[r][idxSexe] || '').toUpperCase().trim();
          const sexeCell = sh.getRange(r + 1, idxSexe + 1);
          sexeCell.setFontWeight('bold').setFontColor('#000000');

          if (sexeValue === 'F') {
            sexeCell.setBackground('#fce7f3'); // Rose
          } else if (sexeValue === 'M') {
            sexeCell.setBackground('#dbeafe'); // Bleu
          }
        }
      }

      // ========== 5. FORMATAGE COLONNE LV2 - ITA=vert, ESP=orange, CHAV=violet ==========
      if (idxLV2 >= 0 && numRows > 0) {
        for (let r = 1; r <= numRows; r++) {
          const lv2Value = String(data[r][idxLV2] || '').toUpperCase().trim();
          const lv2Cell = sh.getRange(r + 1, idxLV2 + 1);

          if (lv2Value === 'ITA' || lv2Value === 'ITALIEN') {
            lv2Cell.setBackground('#86efac').setFontWeight('bold').setFontColor('#000000'); // Vert
          } else if (lv2Value === 'ESP' || lv2Value === 'ESPAGNOL') {
            lv2Cell.setBackground('#fb923c').setFontWeight('bold').setFontColor('#000000'); // Orange
          } else if (lv2Value === 'CHAV') {
            lv2Cell.setBackground('#5b21b6').setFontWeight('bold').setFontColor('#ffffff'); // Violet
          } else if (lv2Value) {
            lv2Cell.setBackground('#000000').setFontWeight('bold').setFontColor('#ffffff'); // Noir
          }
        }
      }

      // ========== 6. FORMATAGE COLONNE OPT - CHAV=violet, autres=gris ==========
      if (idxOPT >= 0 && numRows > 0) {
        for (let r = 1; r <= numRows; r++) {
          const optValue = String(data[r][idxOPT] || '').toUpperCase().trim();
          const optCell = sh.getRange(r + 1, idxOPT + 1);

          if (optValue === 'CHAV') {
            optCell.setBackground('#5b21b6').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
          } else if (optValue) {
            optCell.setBackground('#d1d5db').setFontColor('#000000').setFontWeight('bold').setHorizontalAlignment('center');
          }
        }
      }

      // ========== 7. FORMATAGE COLONNE ASSO - Gras noir + Fond bleu SI REMPLI ==========
      if (idxASSO >= 0 && numRows > 0) {
        for (let r = 1; r <= numRows; r++) {
          const assoValue = String(data[r][idxASSO] || '').trim();
          const assoCell = sh.getRange(r + 1, idxASSO + 1);
          assoCell.setFontWeight('bold').setFontColor('#000000');

          if (assoValue) {
            assoCell.setBackground('#3b82f6').setFontColor('#ffffff');
          }
        }
      }

      // ========== 8. FORMATAGE COLONNE DISSO - Blanc gras sur bleu SI REMPLI ==========
      if (idxDISSO >= 0 && numRows > 0) {
        for (let r = 1; r <= numRows; r++) {
          const dissoValue = String(data[r][idxDISSO] || '').trim();
          const dissoCell = sh.getRange(r + 1, idxDISSO + 1);

          if (dissoValue) {
            dissoCell.setBackground('#3b82f6').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
          }
        }
      }

      // ========== 9. FORMATAGE CONDITIONNEL DES SCORES (COM, TRA, PART, ABS) ==========
      const scoreIndices = [idxCOM, idxTRA, idxPART, idxABS].filter(function(idx) { return idx >= 0; });

      if (numRows > 0) {
        scoreIndices.forEach(function(colIdx) {
          for (let r = 1; r <= numRows; r++) {
            const scoreValue = data[r][colIdx];
            const scoreCell = sh.getRange(r + 1, colIdx + 1);
            scoreCell.setHorizontalAlignment('center').setFontWeight('bold');

            if (scoreValue === 1 || scoreValue === '1') {
              // Score 1 = Rouge + Blanc gras
              scoreCell.setBackground('#dc2626').setFontColor('#ffffff');
            } else if (scoreValue === 2 || scoreValue === '2') {
              // Score 2 = Jaune CLAIR + Noir gras
              scoreCell.setBackground('#fde047').setFontColor('#000000');
            } else if (scoreValue === 3 || scoreValue === '3') {
              // Score 3 = Vert clair + Noir gras
              scoreCell.setBackground('#86efac').setFontColor('#000000');
            } else if (scoreValue === 4 || scoreValue === '4') {
              // Score 4 = Vert foncé + Blanc gras
              scoreCell.setBackground('#16a34a').setFontColor('#ffffff');
            }
          }
        });
      }

      // ========== 10. BORDURES POUR TOUTES LES DONNÉES ==========
      if (numRows > 0) {
        const dataRange = sh.getRange(2, 1, numRows, headers.length);
        dataRange.setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
      }

      // ========== 11. MOYENNES DYNAMIQUES AVEC FORMULES ==========
      const avgRow = numRows + 4; // +2 pour en-tête, +2 lignes vides

      // Écrire "MOYENNE" dans la première colonne
      sh.getRange(avgRow, 1).setValue('MOYENNE').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');

      const firstDataRow = 2;
      const lastDataRow = numRows + 1;

      // Créer formules AVERAGE pour chaque score
      if (idxCOM >= 0) {
        const colLetter = String.fromCharCode(65 + idxCOM); // A=65, B=66, etc.
        sh.getRange(avgRow, idxCOM + 1).setFormula('=AVERAGE(' + colLetter + firstDataRow + ':' + colLetter + lastDataRow + ')');
      }
      if (idxTRA >= 0) {
        const colLetter = String.fromCharCode(65 + idxTRA);
        sh.getRange(avgRow, idxTRA + 1).setFormula('=AVERAGE(' + colLetter + firstDataRow + ':' + colLetter + lastDataRow + ')');
      }
      if (idxPART >= 0) {
        const colLetter = String.fromCharCode(65 + idxPART);
        sh.getRange(avgRow, idxPART + 1).setFormula('=AVERAGE(' + colLetter + firstDataRow + ':' + colLetter + lastDataRow + ')');
      }
      if (idxABS >= 0) {
        const colLetter = String.fromCharCode(65 + idxABS);
        sh.getRange(avgRow, idxABS + 1).setFormula('=AVERAGE(' + colLetter + firstDataRow + ':' + colLetter + lastDataRow + ')');
      }

      // Formater la ligne des moyennes
      const avgRange = sh.getRange(avgRow, 1, 1, headers.length);
      avgRange
        .setBackground('#f3f4f6')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setFontSize(11)
        .setNumberFormat('0.00')
        .setBorder(true, true, true, true, false, false, '#9ca3af', SpreadsheetApp.BorderStyle.SOLID);

      // ========== 12. DÉCOMPTES LV2 ==========
      if (idxLV2 >= 0) {
        const lv2Row = avgRow + 2;

        // Compter les occurrences
        const lv2Counts = {};
        for (let r = 1; r <= numRows; r++) {
          const lv2Val = String(data[r][idxLV2] || '').toUpperCase().trim();
          if (lv2Val) {
            lv2Counts[lv2Val] = (lv2Counts[lv2Val] || 0) + 1;
          }
        }

        sh.getRange(lv2Row, 1).setValue('DÉCOMPTE LV2').setFontWeight('bold').setFontSize(11);

        let lv2Summary = [];
        for (const lang in lv2Counts) {
          lv2Summary.push(lang + ': ' + lv2Counts[lang]);
        }

        if (idxLV2 + 1 <= headers.length) {
          sh.getRange(lv2Row, idxLV2 + 1).setValue(lv2Summary.join(' | ')).setFontWeight('bold');
        }
      }

      // ========== 13. DÉCOMPTES OPT ==========
      if (idxOPT >= 0) {
        const optRow = avgRow + 3;

        // Compter les occurrences
        const optCounts = {};
        for (let r = 1; r <= numRows; r++) {
          const optVal = String(data[r][idxOPT] || '').toUpperCase().trim();
          if (optVal) {
            optCounts[optVal] = (optCounts[optVal] || 0) + 1;
          }
        }

        sh.getRange(optRow, 1).setValue('DÉCOMPTE OPT').setFontWeight('bold').setFontSize(11);

        let optSummary = [];
        for (const opt in optCounts) {
          optSummary.push(opt + ': ' + optCounts[opt]);
        }

        if (idxOPT + 1 <= headers.length) {
          sh.getRange(optRow, idxOPT + 1).setValue(optSummary.join(' | ')).setFontWeight('bold');
        }
      }

      // ========== 14. FIGER L'EN-TÊTE ==========
      sh.setFrozenRows(1);

      logLine('INFO', '  ✅ ' + sheetName + ' finalisé : ' + numRows + ' élèves, moyennes + décomptes + formatage COMPLET');

    } catch (e) {
      logLine('ERROR', '  ❌ Erreur finalisation ' + sheetName + ': ' + e.message);
    }
  });

  logLine('INFO', '✅ Finalisation onglets TEST terminée avec mise en forme FIN complète');
}

/**
 * 5) Audit final — conformité LV2/OPT/A/D + stats par classe + MOYENNES SCORES
 */
/**
 * 5) Audit final
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG via optStream_init_V2()
 * ✅ AJOUT : Moyennes scores COM/TRA/PART/ABS par classe
 * 🔒 LOCKSERVICE : Empêche les doubles lancements
 */
function auditStream() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: 'Audit déjà en cours', locked: true };
    }
  } catch(e) {
    return { ok: false, error: 'Erreur de verrouillage: ' + e };
  }

  try {
    const ctx = optStream_init_V2();

  // ✅ AUDIT : Vérifier les invariants globaux
  try {
    _assertInvariants_(ctx, 'AUDIT');
  } catch(e) {
    logLine('ERROR', '❌ Audit invariants échoué: ' + e);
  }

  // ✅ AUDIT : Vérifier les quotas et cibles par classe
  try {
    _auditStrictByClass_(ctx, 'AUDIT');
  } catch(e) {
    logLine('ERROR', '❌ Audit strict échoué: ' + e);
  }

  var audit = {};
  try {
    if (typeof auditCacheAgainstStructure_ === "function") {
      audit = auditCacheAgainstStructure_(ctx) || {};
    } else {
      // File d'attente : audit non intégré dans cette build
      audit = { _warning: "auditCacheAgainstStructure_ indisponible dans ce déploiement." };
    }
  } catch (e) {
    audit = { _error: String(e && e.message || e) };
  }

  // ✅ NOUVEAU : Calculer moyennes scores
  var scores = {};
  try {
    scores = computeScoreAveragesByClass_(ctx);
  } catch (e) {
    logLine('ERROR', '❌ Erreur calcul moyennes scores: ' + e);
    scores = { _error: String(e) };
  }

    return {
      ok: true,
      audit: audit,
      scores: scores,
      step: "AUDIT"
    };
  } catch(e) {
    return { ok: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * Pipeline complet streaming (utilisé par startStreaming côté UI)
 */
function runOptimizationStreaming() {
  const t0 = Date.now();
  const steps = [];
  steps.push({ name: "open",   res: openCacheTabsStream() });
  steps.push({ name: "phase1", res: phase1Stream() });
  steps.push({ name: "phase2", res: phase2Stream() });
  steps.push({ name: "phase3", res: phase3Stream() });
  steps.push({ name: "phase4", res: phase4Stream() });
  steps.push({ name: "audit",  res: auditStream() });

  // 🔥 NOUVEAUTÉ : Rouvrir les onglets CACHE à la fin pour affichage auto
  const ctx = optStream_init();
  openCacheTabs_(ctx);
  logLine('INFO', '✅ Onglets CACHE ouverts automatiquement');

  const ms = Date.now() - t0;
  return { ok: true, ms: ms, steps: steps };
}

// ===================================================================
// ANCIENNES FONCTIONS (COMPATIBILITÉ)
// ===================================================================

/**
 * INIT : Construit et retourne le contexte pour les phases streaming
 * Utilisé par les endpoints directs (phase1Stream, phase2Stream, etc.)
 *
 * ⚠️ SYSTÈME LEGACY : Utilise _STRUCTURE
 */
function optStream_init(args) {
  const mode = (args && args.sourceFamily) || 'TEST';
  const ctx = buildContext_({ sourceFamily: mode, targetFamily: 'CACHE' });

  // --- NORMALISATION UNIVERSELLE ---
  normalizeSheetsInCtx_(ctx);

  // Log utile en streaming
  logLine('INFO', '🔧 STREAM CTX (LEGACY): levels=' + JSON.stringify(ctx.levels) +
    ' | src=' + JSON.stringify(ctx.srcSheets) +
    ' | cache=' + JSON.stringify(ctx.cacheSheets));

  return ctx;  // ✅ Retourne le contexte normalisé
}

/**
 * INIT V2 : Construit le contexte depuis _OPTI_CONFIG
 * Utilisé par le SYSTÈME NOUVEAU (UI Optimisation + BASEOPTI)
 *
 * ✅ SYSTÈME NOUVEAU : Utilise _OPTI_CONFIG
 *
 * @param {Object} args - Options (sourceFamily, targetFamily)
 * @returns {Object} Contexte pour les phases BASEOPTI
 */
function optStream_init_V2(args) {
  logLine('INFO', '🆕 Initialisation contexte V2 (depuis _OPTI_CONFIG)...');

  // Construire le contexte depuis _OPTI_CONFIG
  const ctx = buildCtx_V2(args);

  // Normalisation (si nécessaire pour compatibilité)
  if (!ctx.levels) ctx.levels = ctx.niveaux;
  if (!ctx.mode) ctx.mode = ctx.modeSrc;

  // Log utile en streaming
  logLine('INFO', '🔧 STREAM CTX (V2): levels=' + JSON.stringify(ctx.levels) +
    ' | src=' + JSON.stringify(ctx.srcSheets) +
    ' | cache=' + JSON.stringify(ctx.cacheSheets));
  logLine('INFO', '  📊 Effectifs cibles: ' + JSON.stringify(ctx.targets));
  logLine('INFO', '  📌 Quotas: ' + JSON.stringify(ctx.quotas));
  logLine('INFO', '  ⚖️ Poids: ' + JSON.stringify(ctx.weights));

  return ctx;  // ✅ Retourne le contexte normalisé V2
}

/**
 * Lit l'onglet _STRUCTURE et retourne les données formatées
 * Utilisé par normalizeSheetsInCtx_ pour détecter les classes
 */
function readStructureSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('_STRUCTURE');

  if (!sheet) {
    throw new Error('Onglet _STRUCTURE introuvable');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { rows: [] };
  }

  const headers = data[0];
  const classeCol = headers.indexOf('CLASSE_DEST');
  const classeColAlt = headers.indexOf('Classe');
  const idxClasse = classeCol >= 0 ? classeCol : classeColAlt;

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const classe = idxClasse >= 0 ? String(row[idxClasse] || '').trim() : '';

    if (classe) {
      rows.push({
        classe: classe,
        Classe: classe,
        0: classe,  // Pour compatibilité avec r[0]
        1: classe   // Pour compatibilité avec r[1]
      });
    }
  }

  return { rows: rows };
}

/**
 * Helpers de normalisation pour streaming
 */
function normalizeSheetsInCtx_(ctx) {
  const srcTag   = (ctx.mode || ctx.sourceTag || 'TEST').toString().trim();   // ex: TEST
  const cacheTag = (ctx.target || ctx.cacheTag  || 'CACHE').toString().trim(); // ex: CACHE

  // 1) Niveaux/classes de base
  let levels = ctx.levels || ctx.niveaux || ctx.niveauxStr;
  if (Array.isArray(levels)) {
    // ok
  } else if (typeof levels === 'string') {
    levels = levels.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  } else {
    // Fallback : détecter depuis STRUCTURE si dispo
    try {
      const st = readStructureSheet_(); // doit exposer les classes en colonne 0/1
      const tmp = new Set();
      (st.rows || []).forEach(function(r) {
        const c = String(r.classe || r.Classe || r[0] || r[1] || '').trim();
        if (c) tmp.add(c.replace(/(CACHE|TEST|FIN)$/,''));
      });
      levels = Array.from(tmp);
    } catch(e) {
      levels = [];
    }
  }
  ctx.levels = levels;
  ctx.niveaux = levels; // Alias pour compatibilité

  // 2) Construire src/cache sheets de façon sûre (Arrays)
  if (!Array.isArray(ctx.srcSheets)) {
    if (typeof ctx.srcSheets === 'string') {
      ctx.srcSheets = ctx.srcSheets.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    } else {
      ctx.srcSheets = (levels || []).map(function(l) { return l + srcTag; });
    }
  }
  if (!Array.isArray(ctx.cacheSheets)) {
    if (typeof ctx.cacheSheets === 'string') {
      ctx.cacheSheets = ctx.cacheSheets.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    } else {
      ctx.cacheSheets = (levels || []).map(function(l) { return l + cacheTag; });
    }
  }
}

/**
 * PHASE 1 : Options & LV2
 */
function optStream_phase1(args) {
  try {
    logLine('INFO', '='.repeat(80));
    logLine('INFO', '📌 STREAMING PHASE 1 - Options & LV2');
    logLine('INFO', '='.repeat(80));
    
    const mode = (args && args.sourceFamily) || 'TEST';
    const ctx = buildContext_({ sourceFamily: mode, targetFamily: 'CACHE' });
    
    // ⚠️ Facultatif mais recommandé : nettoyer les colonnes LV2/OPT avant peuplement
    clearLv2OptColumnsInCache_(ctx);
    
    // Exécuter Phase 1
    const p1 = Phase1I_dispatchOptionsLV2_(ctx);
    
    // Calculer mobilité après Phase 1
    computeMobilityFlags_(ctx);
    
    SpreadsheetApp.flush();
    
    // Afficher en direct la 1ère classe cible
    focusFirstCacheTab_(ctx);
    
    logLine('INFO', '✅ PHASE 1 terminée: ' + JSON.stringify(p1.counts));
    
    return {
      ok: true,
      step: 'P1',
      counts: p1.counts || {},
      warnings: p1.warnings || [],
      message: 'Phase 1 terminée'
    };
  } catch (e) {
    logLine('ERROR', 'Erreur Phase 1: ' + e.message);
    return { ok: false, error: String(e) };
  }
}

/**
 * PHASE 2 : Codes DISSO/ASSO
 */
function optStream_phase2(args) {
  try {
    logLine('INFO', '='.repeat(80));
    logLine('INFO', '📌 STREAMING PHASE 2 - Codes DISSO/ASSO');
    logLine('INFO', '='.repeat(80));
    
    const mode = (args && args.sourceFamily) || 'TEST';
    const ctx = buildContext_({ sourceFamily: mode, targetFamily: 'CACHE' });
    
    // Exécuter Phase 2
    const p2 = Phase2I_applyDissoAsso_(ctx);
    
    // Recalculer mobilité après Phase 2
    computeMobilityFlags_(ctx);
    
    SpreadsheetApp.flush();
    
    // Afficher en direct
    focusFirstCacheTab_(ctx);
    
    logLine('INFO', '✅ PHASE 2 terminée: ' + JSON.stringify(p2.counts));
    
    return {
      ok: true,
      step: 'P2',
      counts: p2.counts || {},
      warnings: p2.warnings || [],
      message: 'Phase 2 terminée'
    };
  } catch (e) {
    logLine('ERROR', 'Erreur Phase 2: ' + e.message);
    return { ok: false, error: String(e) };
  }
}

/**
 * PHASE 3 : Effectifs & Parité
 */
function optStream_phase3(args) {
  try {
    logLine('INFO', '='.repeat(80));
    logLine('INFO', '📌 STREAMING PHASE 3 - Effectifs & Parité');
    logLine('INFO', '='.repeat(80));
    
    const mode = (args && args.sourceFamily) || 'TEST';
    const ctx = buildContext_({ sourceFamily: mode, targetFamily: 'CACHE' });
    
    // Exécuter Phase 3
    const p3 = Phase3I_completeAndParity_(ctx);
    
    // Recalculer mobilité après Phase 3
    computeMobilityFlags_(ctx);
    
    SpreadsheetApp.flush();
    
    // Afficher en direct
    focusFirstCacheTab_(ctx);
    
    logLine('INFO', '✅ PHASE 3 terminée');
    
    return {
      ok: true,
      step: 'P3',
      counts: p3.counts || {},
      warnings: p3.warnings || [],
      message: 'Phase 3 terminée'
    };
  } catch (e) {
    logLine('ERROR', 'Erreur Phase 3: ' + e.message);
    return { ok: false, error: String(e) };
  }
}

/**
 * PHASE 4 : Swaps + Audit final
 */
function optStream_phase4(args) {
  try {
    logLine('INFO', '='.repeat(80));
    logLine('INFO', '📌 STREAMING PHASE 4 - Swaps & Audit');
    logLine('INFO', '='.repeat(80));
    
    const mode = (args && args.sourceFamily) || 'TEST';
    const ctx = buildContext_({ sourceFamily: mode, targetFamily: 'CACHE' });
    
    // Exécuter Phase 4 avec respect FIXE/PERMUT
    const p4 = Phase4_balanceScoresSwaps_(ctx);
    
    SpreadsheetApp.flush();
    
    // Afficher en direct
    focusFirstCacheTab_(ctx);
    
    // Audit final
    const cacheAudit = (typeof auditCacheAgainstStructure_ === 'function')
      ? auditCacheAgainstStructure_(ctx) || {}
      : {};
    
    logLine('INFO', '✅ PHASE 4 terminée: ' + (p4.swapsApplied || 0) + ' swaps appliqués');
    logLine('INFO', '='.repeat(80));
    
    return {
      ok: true,
      step: 'P4',
      swaps: p4.swapsApplied || 0,
      warnings: p4.warnings || [],
      cacheAudit: cacheAudit,
      message: 'Phase 4 terminée'
    };
  } catch (e) {
    logLine('ERROR', 'Erreur Phase 4: ' + e.message);
    return { ok: false, error: String(e) };
  }
}

// ===================================================================
// WRAPPER POUR COMPATIBILITÉ
// ===================================================================

/**
 * Wrapper pour appeler le streaming complet depuis l'ancienne interface
 * Exécute toutes les phases en séquence (mode non-streaming)
 */
function runOptimizationV14FullI_Sequential(options) {
  const results = [];
  
  try {
    // INIT
    const init = optStream_init(options);
    results.push(init);
    
    // Phase 1
    const p1 = optStream_phase1(options);
    results.push(p1);
    
    // Phase 2
    const p2 = optStream_phase2(options);
    results.push(p2);
    
    // Phase 3
    const p3 = optStream_phase3(options);
    results.push(p3);
    
    // Phase 4
    const p4 = optStream_phase4(options);
    results.push(p4);
    
    // ===================================================================
    // 📊 ANALYTICS : SAUVEGARDER UN SNAPSHOT
    // ===================================================================
    // Après optimisation réussie, générer et sauvegarder un snapshot analytique
    if (typeof buildAnalyticsSnapshot === 'function' && typeof saveAnalyticsSnapshot === 'function') {
      logLine('INFO', '📊 Génération du snapshot analytique (V14I)...');
      
      try {
        const snapshotResult = buildAnalyticsSnapshot({ pipeline: 'V14I' });
        
        if (snapshotResult.success) {
          const optimizationResult = {
            swaps: p4.swaps,
            duration: 0, // À calculer si nécessaire
            improvement: 0 // À calculer si nécessaire
          };
          
          const saveResult = saveAnalyticsSnapshot(snapshotResult.snapshot, optimizationResult);
          
          if (saveResult.success) {
            logLine('INFO', '✅ Snapshot analytique sauvegardé (V14I)');
          } else {
            logLine('WARN', '⚠️ Échec sauvegarde snapshot: ' + saveResult.error);
          }
        } else {
          logLine('WARN', '⚠️ Échec génération snapshot: ' + snapshotResult.error);
        }
      } catch (e) {
        logLine('ERROR', '❌ Erreur analytics (V14I): ' + e.message);
      }
    } else {
      logLine('WARN', '⚠️ Module Analytics non disponible (V14I)');
    }
    
    // Résultat final
    return {
      success: true,
      ok: true,
      nbSwaps: p4.swaps,
      swaps: p4.swaps,
      cacheAudit: p4.cacheAudit,
      cacheSheets: init.ctxInfo.cacheSheets,
      quotasLus: init.ctxInfo.quotas,
      phases: results,
      message: 'Optimisation réussie (mode séquentiel)'
    };
    
  } catch (e) {
    logLine('ERROR', 'Erreur optimisation séquentielle: ' + e.message);
    return {
      success: false,
      ok: false,
      error: e.message,
      phases: results
    };
  }
}

// ===================================================================
// MINI-GARDIEN PHASE 4 — anti-swaps destructeurs (PATCH V2)
// ===================================================================

/**
 * ============================================================
 *      MINI-GARDIEN PHASE 4 — anti-swaps destructeurs
 * ============================================================
 * Empêche les permutations qui cassent :
 *  - L'offre de la classe cible (LV2/OPT non offerte)
 *  - Les quotas réalisés < quotas attendus
 *  - Les groupes A consolidés (si _GROUP_A_LOCK est actif)
 */
function Phase4_optimizeSwaps_Guarded_(ctx, opts) {
  // On s'appuie sur l'implémentation existante si elle expose un générateur de swaps,
  // sinon on applique le garde juste avant chaque application.
  const base = typeof Phase4_optimizeSwaps_ === "function" ? Phase4_optimizeSwaps_ : null;
  if (!base) throw new Error("Phase4_optimizeSwaps_ manquante");
  // Monkey-patch d'un hook "applySwap" si supporté, sinon garde locale
  if (typeof setSwapApplyHook_ === "function") {
    setSwapApplyHook_(function guardedApplySwap(swap) {
      // swap: { s1, s2, fromClass, toClass, ... }
      return isSwapAllowed_(ctx, swap.s1, swap.fromClass, swap.s2, swap.toClass);
    });
    return base(ctx, opts);
  }
  // Fallback : si pas de hook, on filtre à la volée via un wrapper
  const apply = function(swap) {
    return isSwapAllowed_(ctx, swap.s1, swap.fromClass, swap.s2, swap.toClass);
  };
  // Appelle l'algo avec un "apply" local si l'API le permet,
  // sinon on n'a pas le point d'injection : on applique un budget 0.
  if (typeof base === "function" && base.length >= 2) {
    // On suppose que l'algo lit/appelle une fonction globale tryApplySwap_ si dispo.
    if (typeof tryApplySwap_ === "function") {
      const orig = tryApplySwap_;
      tryApplySwap_ = function(swap) {
        if (!apply(swap)) return false;
        return orig(swap);
      };
      try {
        return base(ctx, opts);
      } finally {
        tryApplySwap_ = orig;
      }
    } else {
      // Pas de point d'ancrage ⇒ éviter les swaps risqués
      logLine('WARN', '⚠️ Aucun hook de swap disponible – Phase4 exécutée en lecture seule');
      return { ok:true, swaps:0, guarded:true };
    }
  }
  return { ok:false, reason:"No integration point for guarded swaps" };
}

function isSwapAllowed_(ctx, s1, fromClass, s2, toClass) {
  // s1 et s2 sont attendus comme objets élève { LV2, OPT, A, _ID, ... }
  // fromClass / toClass : noms simples ("6°1", pas "...CACHE")
  if (!s1 || !s2) return false;

  const offers = (ctx.offersByClass || ctx.offers || {}); // { "6°1": {LV2:[...], OPT:[...]} }
  const quotas = (ctx.quotas || {});                      // { "6°1": {ITA:6, CHAV:0,...}, ...}

  // 1) Respect de l'offre : un élève avec LV2/OPT ne peut aller que dans une classe qui l'offre
  const okOffer1 = matchesClassOffer_(offers[toClass], s1);
  const okOffer2 = matchesClassOffer_(offers[fromClass], s2); // car s2 va partir DE toClass VERS fromClass (swap croisé)
  if (!okOffer1 || !okOffer2) return false;

  // 2) Quotas réalisés ≥ quotas attendus après swap ?
  if (!respectsQuotasAfterSwap_(ctx, quotas, s1, fromClass, toClass)) return false;
  if (!respectsQuotasAfterSwap_(ctx, quotas, s2, toClass, fromClass)) return false;

  // 3) Groupes A consolidés : si un lock est actif, refuser si s1 et s2 appartiennent à des groupes A
  //   qui seraient brisés en changeant de classe.
  if (ctx.lockGroupsA) {
    if (wouldBreakGroupA_(s1, fromClass, toClass)) return false;
    if (wouldBreakGroupA_(s2, toClass, fromClass)) return false;
  }
  return true;
}

function matchesClassOffer_(classOffer, stu) {
  if (!classOffer) return true; // pas d'info ⇒ on n'interdit pas
  const LV2 = (stu.LV2||"").trim().toUpperCase();
  const OPT = (stu.OPT||"").trim().toUpperCase();
  if (LV2 && Array.isArray(classOffer.LV2) && classOffer.LV2.length) {
    if (classOffer.LV2.indexOf(LV2) === -1) return false;
  }
  if (OPT && Array.isArray(classOffer.OPT) && classOffer.OPT.length) {
    if (classOffer.OPT.indexOf(OPT) === -1) return false;
  }
  return true;
}

function respectsQuotasAfterSwap_(ctx, quotas, stu, fromClass, toClass) {
  // Simplifié : si la classe FROM est à peine au quota d'une option/LV2 du student,
  // le retirer casserait le quota ⇒ refuser.
  const qFrom = quotas[fromClass] || {};
  const qTo   = quotas[toClass]   || {};

  const lv2 = (stu.LV2||"").toUpperCase();
  const opt = (stu.OPT||"").toUpperCase();

  // Compteurs actuels (réalisés) dans CACHE
  const realizedFrom = countRealized_(ctx, fromClass);
  const realizedTo   = countRealized_(ctx, toClass);

  // Test "ne pas descendre sous le quota"
  if (lv2 && qFrom[lv2] != null) {
    const current = (realizedFrom.LV2[lv2]||0);
    if (current - 1 < qFrom[lv2]) return false;
  }
  if (opt && qFrom[opt] != null) {
    const current = (realizedFrom.OPT[opt]||0);
    if (current - 1 < qFrom[opt]) return false;
  }

  // On peut raffiner avec un test symétrique côté toClass si tu veux imposer
  // de ne swapper QUE si cela rapproche toClass de son quota (optionnel).
  return true;
}

function countRealized_(ctx, clazz) {
  // Compte LV2/OPT réalisés dans la classe `clazz` depuis ...CACHE
  const sheetName = clazz+"CACHE";
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const out = { LV2:{}, OPT:{} };
  if (!sh) return out;
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return out;
  const head = values[0].map(String); const h = {};
  head.forEach(function(k,i) { h[k]=i; });
  for (let i=1;i<values.length;i++) {
    const r = values[i];
    const nom = String(r[h["NOM"]]||""); const prenom = String(r[h["PRENOM"]]||"");
    if (!nom && !prenom) continue;
    const lv2 = String(r[h["LV2"]]||"").toUpperCase();
    const opt = String(r[h["OPT"]]||"").toUpperCase();
    if (lv2) out.LV2[lv2] = (out.LV2[lv2]||0)+1;
    if (opt) out.OPT[opt] = (out.OPT[opt]||0)+1;
  }
  return out;
}

function wouldBreakGroupA_(stu, fromClass, toClass) {
  // Simplifié : si l'élève porte un code A (A1..A9), on suppose qu'il est "consolidé"
  // dans sa classe actuelle, on refuse de le déplacer si la politique lock est on.
  const a = String(stu.A||"").trim();
  if (!a) return false;
  // Si besoin : vérifier que d'autres membres du même groupe A sont encore dans fromClass.
  return true;
}
