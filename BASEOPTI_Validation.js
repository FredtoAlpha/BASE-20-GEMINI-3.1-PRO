/**
 * ===================================================================
 * VALIDATION _BASEOPTI — Garde-fous et audits
 * ===================================================================
 * 
 * Fonctions de validation pour garantir l'intégrité de _BASEOPTI
 * avant l'exécution des phases d'optimisation.
 */

/**
 * Valide que _BASEOPTI a le schéma attendu
 * @returns {Object} {ok: boolean, errors: Array<string>}
 */
function validateBaseoptiSchema_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('_BASEOPTI');
  const errors = [];
  
  if (!sh) {
    errors.push('Feuille _BASEOPTI introuvable');
    return { ok: false, errors: errors };
  }
  
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    errors.push('_BASEOPTI vide (aucun élève)');
    return { ok: false, errors: errors };
  }
  
  const headers = values[0].map(String);
  
  // Colonnes obligatoires
  const requiredCols = ['_ID', '_PLACED', '_TARGET_CLASS', 'COM', 'TRA', 'PART', 'ABS', 'SEXE', 'NOM', 'PRENOM'];
  const missing = [];
  
  requiredCols.forEach(function(col) {
    const h = resolveHeader_(col, headers);
    if (!h) {
      missing.push(col);
    }
  });
  
  if (missing.length > 0) {
    errors.push('Colonnes manquantes: ' + missing.join(', '));
    errors.push('En-têtes disponibles: ' + headers.join(', '));
  }
  
  // Vérifier que _ID est rempli
  let emptyIds = 0;
  const hId = resolveHeader_('_ID', headers);
  if (hId) {
    for (let i = 1; i < values.length; i++) {
      const id = String(values[i][hId.idx] || '').trim();
      if (!id) emptyIds++;
    }
    if (emptyIds > 0) {
      errors.push(emptyIds + ' élèves sans _ID');
    }
  }
  
  // Vérifier unicité des _ID
  if (hId) {
    const ids = [];
    for (let i = 1; i < values.length; i++) {
      const id = String(values[i][hId.idx] || '').trim();
      if (id) ids.push(id);
    }
    const uniqueIds = new Set(ids);
    if (uniqueIds.size !== ids.length) {
      errors.push('IDs dupliqués: ' + ids.length + ' IDs pour ' + uniqueIds.size + ' uniques');
    }
  }
  
  return {
    ok: errors.length === 0,
    errors: errors,
    totalRows: values.length - 1,
    headers: headers
  };
}

/**
 * Garde-fou : Refuse de lancer l'optimisation si _BASEOPTI invalide
 * @throws {Error} Si _BASEOPTI invalide
 */
function assertBaseoptiValid_() {
  const validation = validateBaseoptiSchema_();
  
  if (!validation.ok) {
    logLine('ERROR', '❌ _BASEOPTI invalide :');
    validation.errors.forEach(function(err) {
      logLine('ERROR', '   - ' + err);
    });
    logLine('ERROR', '');
    logLine('ERROR', '🔧 Solution : Reconstruire _BASEOPTI via le bouton UI ou createBaseOpti_()');
    throw new Error('_BASEOPTI invalide - Reconstruction nécessaire');
  }
  
  logLine('INFO', '✅ _BASEOPTI valide : ' + validation.totalRows + ' élèves, ' + validation.headers.length + ' colonnes');
}

/**
 * Audit rapide post-création de _BASEOPTI
 */
function auditBaseoptiPostCreation_() {
  logLine('INFO', '🔍 Audit post-création _BASEOPTI...');
  
  const validation = validateBaseoptiSchema_();
  
  if (validation.ok) {
    logLine('INFO', '  ✅ Schéma valide : ' + validation.totalRows + ' élèves');
    logLine('INFO', '  ✅ En-têtes : ' + validation.headers.length + ' colonnes');
    logLine('INFO', '  ✅ IDs uniques et remplis');
  } else {
    logLine('ERROR', '  ❌ Schéma invalide :');
    validation.errors.forEach(function(err) {
      logLine('ERROR', '     - ' + err);
    });
    throw new Error('Audit _BASEOPTI échoué');
  }
}

/**
 * Backfill des scores depuis les colonnes sources
 * @param {Object} work - Objet élève depuis la source
 * @returns {Object} Objet avec scores backfillés
 */
function backfillScores_(work) {
  return {
    COM: work.COM !== undefined ? work.COM : (work.SCORE_COM !== undefined ? work.SCORE_COM : 0),
    TRA: work.TRA !== undefined ? work.TRA : (work.SCORE_TRA !== undefined ? work.SCORE_TRA : 0),
    PART: work.PART !== undefined ? work.PART : (work.SCORE_PART !== undefined ? work.SCORE_PART : 0),
    ABS: work.ABS !== undefined ? work.ABS : (work.SCORE_ABS !== undefined ? work.SCORE_ABS : 0)
  };
}

/**
 * Idempotence : Ne réécrire _ID que s'il est vide
 * @param {Object} work - Objet élève
 * @param {string} srcName - Nom de la source
 * @param {number} rowIdx - Index de la ligne
 * @returns {string} ID stable
 */
function ensureStableId_(work, srcName, rowIdx) {
  // Si _ID déjà présent, le conserver
  if (work._ID && String(work._ID).trim() !== '') {
    return String(work._ID).trim();
  }
  
  // Sinon, chercher ID_ELEVE ou ID
  if (work.ID_ELEVE && String(work.ID_ELEVE).trim() !== '') {
    return String(work.ID_ELEVE).trim();
  }
  
  if (work.ID && String(work.ID).trim() !== '') {
    return String(work.ID).trim();
  }
  
  // Sinon, générer un ID stable
  return buildStableId_(work, srcName, rowIdx);
}
