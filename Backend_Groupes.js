/**
 * ===================================================================
 * BACKEND_GROUPES.GS - GESTION DU MODULE GROUPES V4
 * ===================================================================
 * Module responsable de la logique des Groupes (Besoins, LV, etc)
 * Sauvegarde, chargement des scores, algorithme de groupage
 * Extrait du Code.gs originel (Lignes 418+)
 * ===================================================================
 */

// Configuration des groupes
const GROUPS_CONFIG = {
  maxGroupSize: 8,
  minGroupSize: 3,
  scoreColumns: {
    scoreF: 'U', // Colonne U : Score F
    scoreM: 'V'  // Colonne V : Score M
  },
  criteria: {
    needsSpecial: 'BESOINS_SPECIAUX',
    livingLanguage: 'LANGUE_VIVANTE',
    transport: 'TRANSPORT'
  }
};

// loadFINSheetsWithScores() → supprimée (définition canonique dans Code.js avec SHEET_PATTERNS.FIN + SCORE_COLUMNS)

/**
 * Sauvegarde une configuration de groupes
 * @param {Array} groups - Les groupes à sauvegarder
 * @param {string} configName - Nom de la configuration
 * @returns {boolean} Succès
 */
function saveGroupsConfiguration(groups, configName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Chercher ou créer l'onglet de configuration
    let configSheet = ss.getSheetByName('_GROUPES_CONFIG');
    if (!configSheet) {
      configSheet = ss.insertSheet('_GROUPES_CONFIG');
    } else {
      configSheet.clearContents();
    }

    // En-têtes
    const headers = ['GroupeName', 'Élèves', 'Taille', 'Score_Moyen', 'Timestamp'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Données
    let row = 2;
    groups.forEach((group, idx) => {
      const studentIds = group.students.map(s => s.id).join('; ');
      const avgScore = group.students.length > 0
        ? group.students.reduce((sum, s) => sum + (s.scores.female + s.scores.male) / 2, 0) / group.students.length
        : 0;

      configSheet.getRange(row, 1, 1, headers.length).setValues([[
        `${configName}_Groupe_${idx + 1}`,
        studentIds,
        group.students.length,
        avgScore.toFixed(2),
        new Date().toISOString()
      ]]);
      row++;
    });

    SpreadsheetApp.flush();
    Logger.log(`[SUCCESS] Configuration "${configName}" sauvegardée (${groups.length} groupes)`);
    return true;

  } catch (e) {
    Logger.log(`[ERROR] Erreur sauvegarde groupes : ${e.toString()}`);
    return false;
  }
}

/**
 * Charge une configuration de groupes existante
 * @param {string} configName - Nom de la configuration
 * @returns {Array} Les groupes
 */
function loadGroupsConfiguration(configName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('_GROUPES_CONFIG');

    if (!configSheet) {
      Logger.log('[WARN] Onglet _GROUPES_CONFIG introuvable');
      return [];
    }

    const data = configSheet.getDataRange().getValues();
    const groups = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]).includes(configName)) {
        const studentIds = String(row[1]).split('; ').filter(s => s.trim() !== '');
        groups.push({
          name: row[0],
          studentIds: studentIds,
          size: row[2],
          avgScore: row[3]
        });
      }
    }

    Logger.log(`[SUCCESS] Configuration "${configName}" chargée (${groups.length} groupes)`);
    return groups;

  } catch (e) {
    Logger.log(`[ERROR] Erreur chargement groupes : ${e.toString()}`);
    return [];
  }
}

/**
 * Calcule la qualité d'un groupe basée sur les scores
 * @param {Array} students - Élèves du groupe
 * @returns {number} Score de qualité (0-100)
 */
function calculateGroupQuality(students) {
  if (students.length === 0) return 0;

  // Score moyen du groupe
  const avgScore = students.reduce((sum, s) => sum + (s.scores.female + s.scores.male) / 2, 0) / students.length;

  // Variance (hétérogénéité = bon)
  let variance = 0;
  students.forEach(s => {
    const studentScore = (s.scores.female + s.scores.male) / 2;
    variance += Math.pow(studentScore - avgScore, 2);
  });
  variance = Math.sqrt(variance / students.length);

  // Score de qualité : plus les scores sont proches, moins c'est bon (on veut hétérogène)
  return avgScore * 50 + variance * 50;
}

/**
 * Crée des groupes optimisés basés sur les scores
 * @param {Array} students - Liste des élèves
 * @param {number} targetGroupSize - Taille cible des groupes
 * @returns {Array} Groupes créés
 */
function createOptimizedGroups(students, targetGroupSize = 4) {
  if (students.length === 0) return [];

  // Trier par score décroissant
  const sorted = [...students].sort((a, b) => {
    const scoreA = (a.scores.female + a.scores.male) / 2;
    const scoreB = (b.scores.female + b.scores.male) / 2;
    return scoreB - scoreA;
  });

  const groups = [];
  let currentGroup = [];

  sorted.forEach(student => {
    currentGroup.push(student);

    if (currentGroup.length >= targetGroupSize) {
      groups.push({
        name: `Groupe_${groups.length + 1}`,
        students: currentGroup,
        quality: calculateGroupQuality(currentGroup)
      });
      currentGroup = [];
    }
  });

  // Ajouter les élèves restants
  if (currentGroup.length > 0) {
    if (currentGroup.length >= GROUPS_CONFIG.minGroupSize) {
      groups.push({
        name: `Groupe_${groups.length + 1}`,
        students: currentGroup,
        quality: calculateGroupQuality(currentGroup)
      });
    } else if (groups.length > 0) {
      // Fusionner avec le dernier groupe
      groups[groups.length - 1].students.push(...currentGroup);
      groups[groups.length - 1].quality = calculateGroupQuality(groups[groups.length - 1].students);
    }
  }

  Logger.log(`[SUCCESS] ${groups.length} groupes créés`);
  return groups;
}

/**
 * Optimise les groupes en effectuant des swaps intelligents
 * @param {Array} groups - Les groupes à optimiser
 * @param {number} iterations - Nombre d'itérations
 * @returns {Array} Groupes optimisés
 */
function optimizeGroupsV4(groups, iterations = 100) {
  if (groups.length < 2) return groups;

  let optimized = JSON.parse(JSON.stringify(groups));
  let bestQuality = optimized.reduce((sum, g) => sum + g.quality, 0);

  for (let iter = 0; iter < iterations; iter++) {
    const gidx1 = Math.floor(Math.random() * optimized.length);
    const gidx2 = Math.floor(Math.random() * optimized.length);

    if (gidx1 === gidx2 || optimized[gidx1].students.length === 0 || optimized[gidx2].students.length === 0) {
      continue;
    }

    const sidx1 = Math.floor(Math.random() * optimized[gidx1].students.length);
    const sidx2 = Math.floor(Math.random() * optimized[gidx2].students.length);

    // Effectuer le swap
    const temp = optimized[gidx1].students[sidx1];
    optimized[gidx1].students[sidx1] = optimized[gidx2].students[sidx2];
    optimized[gidx2].students[sidx2] = temp;

    // Recalculer les qualités
    const newQuality1 = calculateGroupQuality(optimized[gidx1].students);
    const newQuality2 = calculateGroupQuality(optimized[gidx2].students);
    const newTotalQuality = optimized.reduce((sum, g) => sum + (
      g === optimized[gidx1] ? newQuality1 : g === optimized[gidx2] ? newQuality2 : g.quality
    ), 0);

    // Accepter ou rejeter le swap
    if (newTotalQuality > bestQuality) {
      optimized[gidx1].quality = newQuality1;
      optimized[gidx2].quality = newQuality2;
      bestQuality = newTotalQuality;
    } else {
      // Rejeter : swap inverse
      const temp = optimized[gidx1].students[sidx1];
      optimized[gidx1].students[sidx1] = optimized[gidx2].students[sidx2];
      optimized[gidx2].students[sidx2] = temp;
    }
  }

  Logger.log(`[SUCCESS] Optimisation V4 complétée (Qualité: ${bestQuality.toFixed(2)})`);
  return optimized;
}

/**
 * Récupère les statistiques globales des groupes
 * @param {Array} groups - Les groupes
 * @returns {Object} Statistiques
 */
function getGroupsStatistics(groups) {
  const totalStudents = groups.reduce((sum, g) => sum + g.students.length, 0);
  const avgGroupSize = groups.length > 0 ? totalStudents / groups.length : 0;
  const totalQuality = groups.reduce((sum, g) => sum + g.quality, 0);
  const avgQuality = groups.length > 0 ? totalQuality / groups.length : 0;

  return {
    groupCount: groups.length,
    totalStudents: totalStudents,
    avgGroupSize: avgGroupSize,
    avgQuality: avgQuality,
    groups: groups.map(g => ({
      name: g.name,
      size: g.students.length,
      quality: g.quality
    }))
  };
}

/**
 * Exporte les groupes en format texte/tableau
 * @param {Array} groups - Les groupes à exporter
 * @returns {string} Représentation textuelle
 */
function exportGroupsAsText(groups) {
  let text = '=== EXPORT GROUPES V4 ===\n\n';

  groups.forEach((group, idx) => {
    text += `${group.name} (${group.students.length} élèves, Qualité: ${group.quality.toFixed(2)})\n`;
    text += '─'.repeat(50) + '\n';

    group.students.forEach(student => {
      const score = (student.scores.female + student.scores.male) / 2;
      text += `  • ${student.prenom} ${student.nom} (Score: ${score.toFixed(2)})\n`;
    });

    text += '\n';
  });

  return text;
}

/**
 * Valide les groupes (taille, complétude)
 * @param {Array} groups - Les groupes
 * @returns {Object} Résultat de validation
 */
function validateGroups(groups) {
  const errors = [];
  const warnings = [];

  if (!Array.isArray(groups) || groups.length === 0) {
    errors.push('Aucun groupe défini');
    return { ok: false, errors, warnings };
  }

  groups.forEach((group, idx) => {
    if (group.students.length < GROUPS_CONFIG.minGroupSize) {
      warnings.push(`Groupe ${idx + 1} : Effectif trop réduit (${group.students.length}/${GROUPS_CONFIG.minGroupSize})`);
    }
    if (group.students.length > GROUPS_CONFIG.maxGroupSize) {
      warnings.push(`Groupe ${idx + 1} : Effectif trop élevé (${group.students.length}/${GROUPS_CONFIG.maxGroupSize})`);
    }
  });

  const totalStudents = groups.reduce((sum, g) => sum + g.students.length, 0);
  const uniqueIds = new Set(groups.flatMap(g => g.students.map(s => s.id)));

  if (uniqueIds.size !== totalStudents) {
    errors.push('Doublons détectés dans les groupes');
  }

  return {
    ok: errors.length === 0,
    errors: errors,
    warnings: warnings,
    summary: `${groups.length} groupes, ${totalStudents} élèves, ${uniqueIds.size} uniques`
  };
}
