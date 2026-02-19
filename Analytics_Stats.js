/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ANALYTICS_STATS.GS - Analyse complète des données consolidées
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Lit l'onglet CONSOLIDATION et calcule toutes les statistiques nécessaires
 * pour la Phase STATS de la Console V3.
 */

/**
 * Fonction principale : récupère toutes les statistiques depuis CONSOLIDATION
 * @returns {Object} Objet contenant toutes les métriques
 */
function getConsolidationStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const consoSheet = ss.getSheetByName('CONSOLIDATION');

    if (!consoSheet) {
      return {
        success: false,
        error: "L'onglet CONSOLIDATION n'existe pas. Veuillez d'abord consolider les données."
      };
    }

    const data = consoSheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: false,
        error: "L'onglet CONSOLIDATION est vide. Veuillez d'abord générer les IDs et consolider."
      };
    }

    const headers = data[0];
    const rows = data.slice(1).filter(row => row[0]); // Filtrer les lignes vides

    // Fonction helper pour trouver une colonne (insensible à la casse et aux espaces)
    const findColumn = (possibleNames) => {
      const normalized = headers.map(h => String(h).trim().toUpperCase());
      for (const name of possibleNames) {
        const idx = normalized.indexOf(name.toUpperCase());
        if (idx !== -1) return idx;
      }
      return -1;
    };

    // Indices des colonnes avec recherche flexible
    const idx = {
      SEXE: findColumn(['SEXE', 'SEX']),
      LV2: findColumn(['LV2', 'LANGUE']),
      OPT: findColumn(['OPT', 'OPTION', 'OPTIONS']),
      DISPOSITIF: findColumn(['DISPOSITIF', 'DISPO', 'DISPOSITIFS']),
      ASSO: findColumn(['ASSO', 'CODES_ASSO', 'CODE_ASSO']),
      DISSO: findColumn(['DISSO', 'CODES_DISSO', 'CODE_DISSO'])
    };

    Logger.log(`📋 Colonnes détectées dans CONSOLIDATION: ${headers.join(', ')}`);
    Logger.log(`📍 Indices: SEXE=${idx.SEXE}, LV2=${idx.LV2}, OPT=${idx.OPT}, DISPOSITIF=${idx.DISPOSITIF}, ASSO=${idx.ASSO}, DISSO=${idx.DISSO}`);

    // Calculer toutes les statistiques
    const stats = {
      success: true,
      effectifs: calculerEffectifs(rows),
      parite: calculerParite(rows, idx.SEXE),
      lv2: calculerLV2(rows, idx.LV2, idx.OPT),
      options: calculerOptions(rows, idx.OPT, idx.LV2),
      combos: calculerCombos(rows, idx.LV2, idx.OPT),
      global: calculerComptagesGlobaux(rows, idx.LV2, idx.OPT), // NOUVEAU : Comptage global unifié
      dispositifs: calculerDispositifs(rows, idx.DISPOSITIF),
      asso: calculerCodesAsso(rows, idx.ASSO),
      disso: calculerCodesDisso(rows, idx.DISSO)
    };

    return stats;

  } catch (e) {
    Logger.log(`Erreur getConsolidationStats: ${e.message}`);
    return {
      success: false,
      error: `Erreur technique: ${e.message}`
    };
  }
}

/**
 * Calcule les effectifs totaux
 */
function calculerEffectifs(rows) {
  return {
    total: rows.length
  };
}

/**
 * Calcule la parité F/M
 */
function calculerParite(rows, sexeIdx) {
  const counts = { F: 0, M: 0, inconnu: 0 };

  rows.forEach(row => {
    const sexe = String(row[sexeIdx] || '').trim().toUpperCase();
    if (sexe === 'F') counts.F++;
    else if (sexe === 'M') counts.M++;
    else counts.inconnu++;
  });

  const total = counts.F + counts.M;

  return {
    F: counts.F,
    M: counts.M,
    inconnu: counts.inconnu,
    ratioF: total > 0 ? ((counts.F / total) * 100).toFixed(2) : 0,
    ratioM: total > 0 ? ((counts.M / total) * 100).toFixed(2) : 0
  };
}

/**
 * Calcule les LV2 SEULES (sans option)
 */
function calculerLV2(rows, lv2Idx, optIdx) {
  const lv2Counts = {};

  rows.forEach(row => {
    const lv2 = String(row[lv2Idx] || '').trim().toUpperCase();
    const opt = String(row[optIdx] || '').trim();

    // Ne compter que si LV2 existe ET pas d'option
    if (lv2 && !opt) {
      lv2Counts[lv2] = (lv2Counts[lv2] || 0) + 1;
    }
  });

  return lv2Counts;
}

/**
 * Calcule les OPTIONS SEULES (sans LV2 spécifique, ou avec ESP par défaut)
 */
function calculerOptions(rows, optIdx, lv2Idx) {
  const optCounts = {};

  rows.forEach(row => {
    const opt = String(row[optIdx] || '').trim().toUpperCase();
    const lv2 = String(row[lv2Idx] || '').trim().toUpperCase();

    // Ne compter que si option existe ET pas de LV2 (car si LV2, c'est un combo)
    if (opt && !lv2) {
      optCounts[opt] = (optCounts[opt] || 0) + 1;
    }
  });

  return optCounts;
}

/**
 * Calcule les PROFILS DOUBLES (LV2 spécifique + Option)
 * Ex: ITA+GREC, ALL+LATIN, ESP+LATIN...
 */
function calculerCombos(rows, lv2Idx, optIdx) {
  const combos = {};
  const debugDetails = [];

  // Helper : découpe commune pour LV2 et options (gère "/", ",", "+"…)
  const splitList = (value) => {
    return String(value || '')
      .toUpperCase()
      .split(/[+,;/]|\s+\+\s+|\s*\/\s*/)
      .map(v => v.trim())
      .filter(Boolean);
  };

  // Si les colonnes n'existent pas, aucun calcul possible (évite de lire la dernière colonne avec l'index -1)
  if (lv2Idx === -1 || optIdx === -1) return combos;

  rows.forEach((row, index) => {
    const lv2List = splitList(row[lv2Idx]);
    const options = splitList(row[optIdx]);

    if (lv2List.length && options.length) {
      // 🔒 Sécurisation : ne compter chaque couple qu'une seule fois par élève,
      // même si l'option ou la LV2 est saisie en double ou avec des séparateurs multiples.
      const seenForRow = new Set();

      lv2List.forEach(lv2 => {
        options.forEach(opt => {
          if (!opt || !lv2 || opt === lv2) return; // Pas de combo si option vide ou identique à la LV2

          const combo = `${lv2} + ${opt}`;
          if (seenForRow.has(combo)) return; // évite de compter deux fois la même paire pour un élève

          combos[combo] = (combos[combo] || 0) + 1;
          seenForRow.add(combo);
        });
      });

      // Stocker une trace détaillée (index de ligne + listes normalisées + combos retenus)
      debugDetails.push({
        ligne: index + 2, // +2 pour compter l'en-tête + index 0-based
        lv2: lv2List,
        options,
        combos: Array.from(seenForRow)
      });
    }
  });

  // Log limité pour investiguer l'affichage des profils doubles
  const maxRows = 50;
  const preview = debugDetails.slice(0, maxRows);
  Logger.log(`🔎 Trace combos (premières ${preview.length} lignes contenant LV2+option, max ${maxRows}): ${JSON.stringify(preview)}`);
  Logger.log(`📊 Totaux combos calculés: ${JSON.stringify(combos)}`);

  return combos;
}

/**
 * Calcule les comptages GLOBAUX pour chaque matière (LV2 ou Option)
 * Permet de savoir combien d'élèves font "ITA" au total (LV2 + Option + Combos)
 */
function calculerComptagesGlobaux(rows, lv2Idx, optIdx) {
  const globalCounts = {};

  // Recycle la même logique de découpe que pour les combos pour éviter les écarts
  const splitList = (value) => {
    return String(value || '')
      .toUpperCase()
      .split(/[+,;/]|\s+\+\s+|\s*\/\s*/)
      .map(v => v.trim())
      .filter(Boolean);
  };

  rows.forEach(row => {
    const lv2List = splitList(row[lv2Idx]);
    const options = splitList(row[optIdx]);

    // Ajouter LV2 (déduplication intra-ligne pour éviter 2× ESP en double saisie)
    if (lv2List.length) {
      const seenLv2 = new Set();
      lv2List.forEach(lv2 => {
        if (!lv2 || seenLv2.has(lv2)) return;
        globalCounts[lv2] = (globalCounts[lv2] || 0) + 1;
        seenLv2.add(lv2);
      });
    }

    // Ajouter chaque option, dédupliquée au sein de la ligne et sans double compter LV2
    if (options.length) {
      const seenOpts = new Set();
      options.forEach(opt => {
        if (!opt || seenOpts.has(opt)) return;
        if (lv2List.includes(opt)) return; // évite de double compter une LV2 saisie aussi en option
        globalCounts[opt] = (globalCounts[opt] || 0) + 1;
        seenOpts.add(opt);
      });
    }
  });

  return globalCounts;
}

/**
 * Calcule les dispositifs (PAI, PPRE, ULIS...)
 */
function calculerDispositifs(rows, dispoIdx) {
  if (dispoIdx === -1) return {};

  const dispoCounts = {};

  rows.forEach(row => {
    const dispo = String(row[dispoIdx] || '').trim().toUpperCase();
    if (dispo) {
      dispoCounts[dispo] = (dispoCounts[dispo] || 0) + 1;
    }
  });

  return dispoCounts;
}

/**
 * Calcule les codes ASSO (A1, A2...)
 */
function calculerCodesAsso(rows, assoIdx) {
  if (assoIdx === -1) return { codes: 0, eleves: 0, details: {} };

  const codesSet = new Set();
  const codeDetails = {};
  let elevesTotal = 0;

  rows.forEach(row => {
    const asso = String(row[assoIdx] || '').trim().toUpperCase();
    if (asso) {
      codesSet.add(asso);
      codeDetails[asso] = (codeDetails[asso] || 0) + 1;
      elevesTotal++;
    }
  });

  return {
    codes: codesSet.size,
    eleves: elevesTotal,
    details: codeDetails
  };
}

/**
 * Calcule les codes DISSO (D1, D2...) avec détection de conflits
 */
function calculerCodesDisso(rows, dissoIdx) {
  if (dissoIdx === -1) return { codes: 0, eleves: 0, details: {}, conflicts: [] };

  const codesSet = new Set();
  const codeDetails = {};
  let elevesTotal = 0;

  rows.forEach(row => {
    const disso = String(row[dissoIdx] || '').trim().toUpperCase();
    if (disso) {
      codesSet.add(disso);
      codeDetails[disso] = (codeDetails[disso] || 0) + 1;
      elevesTotal++;
    }
  });

  return {
    codes: codesSet.size,
    eleves: elevesTotal,
    details: codeDetails,
    conflicts: [] // Sera calculé plus tard avec le nb de classes
  };
}

/**
 * Détecte les conflits DISSO par rapport au nombre de classes cibles
 */
function detecterConflitsDisso(statsDisso, nbClassesCibles) {
  const conflicts = [];

  Object.entries(statsDisso.details).forEach(([code, count]) => {
    if (count > nbClassesCibles) {
      conflicts.push({
        code: code,
        count: count,
        nbClasses: nbClassesCibles,
        message: `⚠️ IMPOSSIBLE: ${count} élèves avec code ${code} pour seulement ${nbClassesCibles} classes`
      });
    }
  });

  return conflicts;
}
