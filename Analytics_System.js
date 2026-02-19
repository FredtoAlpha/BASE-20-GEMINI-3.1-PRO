/**
 * ===================================================================
 * ANALYTICS SYSTEM - Module d'analytique basé sur _OPTI_CONFIG
 * ===================================================================
 * 
 * Fournit une vision consolidée des répartitions, un suivi historique
 * des arbitrages et l'identification des zones de risque.
 * 
 * SOURCES DE DONNÉES :
 * - _OPTI_CONFIG : Configuration et paramètres d'optimisation
 * - ...CACHE : Résultats finaux des optimisations
 * - _BASEOPTI : Données d'optimisation en cours
 * - _ANALYTICS_LOG : Historique des snapshots
 * 
 * @version 1.0
 * @date 2025-10-21
 */

// ===================================================================
// EXTRACTEUR CENTRAL : buildAnalyticsSnapshot()
// ===================================================================

/**
 * Construit un snapshot analytique complet
 * 
 * Lit toutes les sources de données (_OPTI_CONFIG, _BASEOPTI, CACHE)
 * et retourne un objet JSON normalisé avec toutes les métriques.
 * 
 * @param {Object} options - Options de génération (optionnel)
 * @returns {Object} Snapshot analytique complet
 */
function buildAnalyticsSnapshot(options) {
  const startTime = new Date();
  logLine('INFO', '📊 Construction du snapshot analytique...');
  
  options = options || {};
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ===================================================================
    // 1. LIRE _OPTI_CONFIG
    // ===================================================================
    const config = {
      mode: kvGet_('mode.selected') || 'TEST',
      timestamp: kvGet_('mode.timestamp') || new Date().toISOString(),
      targets: {
        byClass: JSON.parse(kvGet_('targets.byClass') || '{}'),
        byOption: JSON.parse(kvGet_('targets.byOption') || '{}')
      },
      quotas: JSON.parse(kvGet_('offers.byClass') || '{}'),
      weights: JSON.parse(kvGet_('weights') || '{"parity":0.3,"com":0.4,"tra":0.1,"part":0.1,"abs":0.1}'),
      swaps: {
        max: parseInt(kvGet_('swaps.max') || '500'),
        runtime: parseInt(kvGet_('swaps.runtime') || '180'),
        history: JSON.parse(kvGet_('swaps.history') || '[]')
      },
      parity: {
        tolerance: parseInt(kvGet_('parity.tolerance') || '2')
      },
      profile: {
        current: kvGet_('profile.current') || 'DEFAULT',
        role: kvGet_('profile.role') || 'TEACHER'
      }
    };
    
    // ===================================================================
    // 2. LIRE LES ONGLETS CACHE
    // ===================================================================
    const cacheSheets = ss.getSheets().filter(sh => sh.getName().endsWith('CACHE'));
    
    const classes = {};
    let totalEleves = 0;
    let totalFemmes = 0;
    let totalHommes = 0;
    const lv2Global = {};
    const optGlobal = {};
    
    for (const sheet of cacheSheets) {
      const className = sheet.getName().replace('CACHE', '').trim();
      const data = sheet.getDataRange().getValues();
      
      if (data.length < 2) continue; // Pas de données
      
      const headers = data[0];
      
      // Trouver les colonnes
      const idxSexe = headers.indexOf('SEXE');
      const idxLV2 = headers.indexOf('LV2');
      const idxOPT = headers.indexOf('OPT');
      const idxCOM = headers.indexOf('COM');
      const idxTRA = headers.indexOf('TRA');
      const idxPART = headers.indexOf('PART');
      const idxABS = headers.indexOf('ABS');
      const idxFIXE = headers.indexOf('FIXE');
      const idxMOBILITE = headers.indexOf('MOBILITE');
      
      // Compter les élèves
      let effectif = 0;
      let femmes = 0;
      let hommes = 0;
      const lv2Count = {};
      const optCount = {};
      let sumCOM = 0, sumTRA = 0, sumPART = 0, sumABS = 0;
      let fixeCount = 0;
      const mobilityStatus = {};
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue; // Ligne vide
        
        effectif++;
        
        // Sexe
        const sexe = String(row[idxSexe] || '').toUpperCase();
        if (sexe === 'F') femmes++;
        if (sexe === 'M') hommes++;
        
        // LV2
        const lv2 = String(row[idxLV2] || '').trim().toUpperCase();
        if (lv2 && lv2 !== 'ESP' && lv2 !== 'ANG') {
          lv2Count[lv2] = (lv2Count[lv2] || 0) + 1;
          lv2Global[lv2] = (lv2Global[lv2] || 0) + 1;
        }
        
        // OPT
        const opt = String(row[idxOPT] || '').trim().toUpperCase();
        if (opt) {
          optCount[opt] = (optCount[opt] || 0) + 1;
          optGlobal[opt] = (optGlobal[opt] || 0) + 1;
        }
        
        // Scores
        if (idxCOM >= 0) sumCOM += parseFloat(row[idxCOM]) || 0;
        if (idxTRA >= 0) sumTRA += parseFloat(row[idxTRA]) || 0;
        if (idxPART >= 0) sumPART += parseFloat(row[idxPART]) || 0;
        if (idxABS >= 0) sumABS += parseFloat(row[idxABS]) || 0;
        
        // Mobilité
        if (idxFIXE >= 0 && row[idxFIXE]) fixeCount++;
        if (idxMOBILITE >= 0) {
          const mob = String(row[idxMOBILITE] || '').trim();
          if (mob) {
            mobilityStatus[mob] = (mobilityStatus[mob] || 0) + 1;
          }
        }
      }
      
      // Calculer les moyennes
      const avgCOM = effectif > 0 ? sumCOM / effectif : 0;
      const avgTRA = effectif > 0 ? sumTRA / effectif : 0;
      const avgPART = effectif > 0 ? sumPART / effectif : 0;
      const avgABS = effectif > 0 ? sumABS / effectif : 0;
      
      // Parité
      const parityRatio = effectif > 0 ? (femmes / effectif) * 100 : 0;
      const parityGap = Math.abs(femmes - hommes);
      const parityOK = parityGap <= config.parity.tolerance;
      
      // Cible
      const target = config.targets.byClass[className] || 0;
      const gapToTarget = effectif - target;
      
      // Quotas
      const quotasClass = config.quotas[className] || {};
      const quotasViolations = [];
      
      for (const opt in quotasClass) {
        const quota = quotasClass[opt] || 0;
        const actual = (lv2Count[opt] || 0) + (optCount[opt] || 0);
        if (actual > quota) {
          quotasViolations.push({
            option: opt,
            quota: quota,
            actual: actual,
            excess: actual - quota
          });
        }
      }
      
      const quotasOK = quotasViolations.length === 0;
      
      classes[className] = {
        effectif,
        femmes,
        hommes,
        parityRatio: parseFloat(parityRatio.toFixed(1)),
        parityGap,
        parityOK,
        target,
        gapToTarget,
        lv2: lv2Count,
        opt: optCount,
        quotasOK,
        quotasViolations,
        scores: {
          com: parseFloat(avgCOM.toFixed(2)),
          tra: parseFloat(avgTRA.toFixed(2)),
          part: parseFloat(avgPART.toFixed(2)),
          abs: parseFloat(avgABS.toFixed(2))
        },
        mobility: {
          fixe: fixeCount,
          status: mobilityStatus
        }
      };
      
      totalEleves += effectif;
      totalFemmes += femmes;
      totalHommes += hommes;
    }
    
    // ===================================================================
    // 3. CALCULER LES MÉTRIQUES GLOBALES
    // ===================================================================
    const nbClasses = Object.keys(classes).length;
    const effectifMoyen = nbClasses > 0 ? totalEleves / nbClasses : 0;
    const pariteGlobale = totalEleves > 0 ? (totalFemmes / totalEleves) * 100 : 0;
    
    // Variance des effectifs
    const effectifs = Object.values(classes).map(c => c.effectif);
    const varianceEffectifs = calculateVariance_(effectifs);
    
    // Variance des scores COM
    const scoresCOM = Object.values(classes).map(c => c.scores.com);
    const varianceScores = calculateVariance_(scoresCOM);
    
    // Quotas et parité respectés partout ?
    const quotasRespectesPartout = Object.values(classes).every(c => c.quotasOK);
    const pariteRespecteePartout = Object.values(classes).every(c => c.parityOK);
    
    // Compter les violations totales
    const totalQuotasViolations = Object.values(classes).reduce((sum, c) => sum + c.quotasViolations.length, 0);
    const totalParityViolations = Object.values(classes).filter(c => !c.parityOK).length;
    
    // ===================================================================
    // 4. IDENTIFIER LES ZONES DE RISQUE
    // ===================================================================
    const risks = identifyRisks_(classes, config);
    
    // ===================================================================
    // 5. GÉNÉRER LES RECOMMANDATIONS
    // ===================================================================
    const recommendations = generateRecommendations_(classes, config);
    
    // ===================================================================
    // 6. CONSTRUIRE LE SNAPSHOT
    // ===================================================================
    const snapshot = {
      timestamp: new Date().toISOString(),
      user: Session.getActiveUser().getEmail(),
      pipeline: options.pipeline || 'UNKNOWN',
      config: config,
      global: {
        totalEleves,
        totalFemmes,
        totalHommes,
        nbClasses,
        effectifMoyen: parseFloat(effectifMoyen.toFixed(1)),
        pariteGlobale: parseFloat(pariteGlobale.toFixed(1)),
        varianceEffectifs: parseFloat(varianceEffectifs.toFixed(2)),
        varianceScores: parseFloat(varianceScores.toFixed(2)),
        quotasRespectesPartout,
        pariteRespecteePartout,
        totalQuotasViolations,
        totalParityViolations,
        lv2Distribution: lv2Global,
        optDistribution: optGlobal
      },
      classes: classes,
      risks: risks,
      recommendations: recommendations,
      metadata: {
        generatedAt: new Date().toISOString(),
        generatedBy: Session.getActiveUser().getEmail(),
        version: '1.0'
      }
    };
    
    const endTime = new Date();
    const durationMs = endTime - startTime;
    
    logLine('INFO', '✅ Snapshot analytique construit en ' + durationMs + 'ms');
    logLine('INFO', '   → ' + totalEleves + ' élèves, ' + nbClasses + ' classes');
    logLine('INFO', '   → Parité globale : ' + pariteGlobale.toFixed(1) + '% F');
    logLine('INFO', '   → Quotas respectés : ' + (quotasRespectesPartout ? 'OUI' : 'NON (' + totalQuotasViolations + ' violations)'));
    logLine('INFO', '   → Parité respectée : ' + (pariteRespecteePartout ? 'OUI' : 'NON (' + totalParityViolations + ' classes)'));
    
    return {
      success: true,
      snapshot: snapshot,
      durationMs: durationMs
    };
    
  } catch (e) {
    logLine('ERROR', '❌ Erreur buildAnalyticsSnapshot : ' + e.message);
    logLine('ERROR', e.stack);
    
    return {
      success: false,
      error: e.message,
      stack: e.stack
    };
  }
}

// ===================================================================
// FONCTIONS UTILITAIRES
// ===================================================================

/**
 * Calcule la variance d'un tableau de nombres
 */
function calculateVariance_(values) {
  if (values.length === 0) return 0;
  
  const mean = values.reduce((sum, v) => sum + v, 0) / values.length;
  const squaredDiffs = values.map(v => Math.pow(v - mean, 2));
  const variance = squaredDiffs.reduce((sum, v) => sum + v, 0) / values.length;
  
  return variance;
}

/**
 * Identifie les zones de risque
 */
function identifyRisks_(classes, config) {
  const risks = [];
  
  for (const className in classes) {
    const cls = classes[className];
    
    // Risque : Effectif trop éloigné de la cible
    if (Math.abs(cls.gapToTarget) > 3) {
      risks.push({
        type: 'EFFECTIF',
        severity: Math.abs(cls.gapToTarget) > 5 ? 'HIGH' : 'MEDIUM',
        class: className,
        message: `Écart de ${cls.gapToTarget > 0 ? '+' : ''}${cls.gapToTarget} élèves par rapport à la cible (${cls.target})`,
        data: {
          actual: cls.effectif,
          target: cls.target,
          gap: cls.gapToTarget
        }
      });
    }
    
    // Risque : Parité non respectée
    if (!cls.parityOK) {
      risks.push({
        type: 'PARITE',
        severity: cls.parityGap > 4 ? 'HIGH' : 'MEDIUM',
        class: className,
        message: `Écart F/M de ${cls.parityGap} élèves (tolérance : ${config.parity.tolerance})`,
        data: {
          femmes: cls.femmes,
          hommes: cls.hommes,
          gap: cls.parityGap,
          tolerance: config.parity.tolerance
        }
      });
    }
    
    // Risque : Quotas dépassés
    if (!cls.quotasOK) {
      for (const violation of cls.quotasViolations) {
        risks.push({
          type: 'QUOTAS',
          severity: 'HIGH',
          class: className,
          message: `Quota ${violation.option} dépassé : ${violation.actual}/${violation.quota} (+${violation.excess})`,
          data: violation
        });
      }
    }
  }
  
  // Trier par sévérité
  risks.sort((a, b) => {
    const severityOrder = { HIGH: 0, MEDIUM: 1, LOW: 2 };
    return severityOrder[a.severity] - severityOrder[b.severity];
  });
  
  return risks;
}

/**
 * Génère des recommandations
 */
function generateRecommendations_(classes, config) {
  const recommendations = [];
  
  // Recommandation : Rééquilibrer les effectifs
  const effectifs = Object.values(classes).map(c => c.effectif);
  const variance = calculateVariance_(effectifs);
  
  if (variance > 5) {
    recommendations.push({
      type: 'REEQUILIBRAGE',
      priority: 'HIGH',
      message: `Variance des effectifs élevée (${variance.toFixed(2)}). Envisager un rééquilibrage entre classes.`,
      action: 'Relancer l\'optimisation avec des cibles ajustées'
    });
  }
  
  // Recommandation : Ajuster la tolérance de parité
  const classesAvecProblemeParite = Object.values(classes).filter(c => !c.parityOK).length;
  
  if (classesAvecProblemeParite > 0) {
    recommendations.push({
      type: 'PARITE',
      priority: 'MEDIUM',
      message: `${classesAvecProblemeParite} classe(s) avec problème de parité.`,
      action: 'Augmenter la tolérance de parité ou relancer l\'optimisation'
    });
  }
  
  // Recommandation : Ajuster les quotas
  const classesAvecQuotasDepasses = Object.values(classes).filter(c => !c.quotasOK).length;
  
  if (classesAvecQuotasDepasses > 0) {
    recommendations.push({
      type: 'QUOTAS',
      priority: 'HIGH',
      message: `${classesAvecQuotasDepasses} classe(s) avec quotas dépassés.`,
      action: 'Vérifier les quotas dans _OPTI_CONFIG et relancer l\'optimisation'
    });
  }
  
  // Recommandation : Optimiser les scores
  const varianceScores = calculateVariance_(Object.values(classes).map(c => c.scores.com));
  
  if (varianceScores > 1.5) {
    recommendations.push({
      type: 'SCORES',
      priority: 'LOW',
      message: `Variance des scores COM élevée (${varianceScores.toFixed(2)}).`,
      action: 'Augmenter le poids COM dans les paramètres d\'optimisation'
    });
  }
  
  return recommendations;
}

// ===================================================================
// PERSISTANCE : _ANALYTICS_LOG
// ===================================================================

/**
 * Sauvegarde un snapshot dans _ANALYTICS_LOG
 * 
 * @param {Object} snapshot - Snapshot à sauvegarder
 * @param {Object} optimizationResult - Résultat de l'optimisation (optionnel)
 * @returns {Object} Résultat de la sauvegarde
 */
function saveAnalyticsSnapshot(snapshot, optimizationResult) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('_ANALYTICS_LOG');
    
    // Créer la feuille si elle n'existe pas
    if (!logSheet) {
      logSheet = ss.insertSheet('_ANALYTICS_LOG');
      
      // En-têtes
      logSheet.getRange(1, 1, 1, 16).setValues([[
        'TIMESTAMP',
        'USER',
        'PIPELINE',
        'MODE',
        'TOTAL_ELEVES',
        'TOTAL_CLASSES',
        'EFFECTIF_MOYEN',
        'PARITE_GLOBALE',
        'VARIANCE_EFFECTIFS',
        'VARIANCE_SCORES',
        'NB_SWAPS',
        'DUREE_SEC',
        'SUCCESS',
        'QUOTAS_RESPECTES',
        'PARITE_RESPECTEE',
        'SNAPSHOT_JSON'
      ]]);
      
      // Formater les en-têtes
      logSheet.getRange(1, 1, 1, 16)
        .setFontWeight('bold')
        .setBackground('#4a5568')
        .setFontColor('#ffffff');
      
      logLine('INFO', '✅ Feuille _ANALYTICS_LOG créée');
    }
    
    // Extraire les données du résultat d'optimisation
    const nbSwaps = optimizationResult ? (optimizationResult.nbSwaps || optimizationResult.swaps || 0) : 0;
    const dureeSec = optimizationResult ? (optimizationResult.durationSec || 0) : 0;
    const success = optimizationResult ? (optimizationResult.success || optimizationResult.ok || false) : true;
    
    // Ajouter la ligne
    const row = [
      snapshot.timestamp,
      snapshot.user,
      snapshot.pipeline,
      snapshot.config.mode,
      snapshot.global.totalEleves,
      snapshot.global.nbClasses,
      snapshot.global.effectifMoyen,
      snapshot.global.pariteGlobale,
      snapshot.global.varianceEffectifs,
      snapshot.global.varianceScores,
      nbSwaps,
      dureeSec,
      success,
      snapshot.global.quotasRespectesPartout,
      snapshot.global.pariteRespecteePartout,
      JSON.stringify(snapshot)
    ];
    
    logSheet.appendRow(row);
    
    logLine('INFO', '✅ Snapshot sauvegardé dans _ANALYTICS_LOG');
    
    return {
      success: true,
      message: 'Snapshot sauvegardé'
    };
    
  } catch (e) {
    logLine('ERROR', '❌ Erreur saveAnalyticsSnapshot : ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Récupère l'historique des snapshots
 * 
 * @param {Object} options - Options de filtrage
 * @returns {Object} Liste des snapshots
 */
function getAnalyticsHistory(options) {
  try {
    options = options || {};
    const limit = options.limit || 10;
    const pipeline = options.pipeline || null;
    const mode = options.mode || null;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('_ANALYTICS_LOG');
    
    if (!logSheet) {
      return {
        success: true,
        snapshots: [],
        count: 0
      };
    }
    
    const data = logSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Trouver les colonnes
    const idxTimestamp = headers.indexOf('TIMESTAMP');
    const idxPipeline = headers.indexOf('PIPELINE');
    const idxMode = headers.indexOf('MODE');
    const idxSnapshot = headers.indexOf('SNAPSHOT_JSON');
    
    // Filtrer et extraire les snapshots
    const snapshots = [];
    
    for (let i = data.length - 1; i >= 1 && snapshots.length < limit; i--) {
      const row = data[i];
      
      // Filtrer par pipeline
      if (pipeline && row[idxPipeline] !== pipeline) continue;
      
      // Filtrer par mode
      if (mode && row[idxMode] !== mode) continue;
      
      try {
        const snapshot = JSON.parse(row[idxSnapshot]);
        snapshots.push(snapshot);
      } catch (e) {
        logLine('WARN', '⚠️ Snapshot invalide à la ligne ' + (i + 1));
      }
    }
    
    return {
      success: true,
      snapshots: snapshots,
      count: snapshots.length
    };
    
  } catch (e) {
    logLine('ERROR', '❌ Erreur getAnalyticsHistory : ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// ===================================================================
// API POUR L'INTERFACE
// ===================================================================

/**
 * Génère et retourne un snapshot pour l'interface
 * 
 * @returns {Object} Snapshot analytique
 */
function getAnalyticsSnapshotForUI() {
  const result = buildAnalyticsSnapshot({ pipeline: 'UI_REQUEST' });
  
  if (result.success) {
    return result.snapshot;
  } else {
    throw new Error(result.error);
  }
}

/**
 * Récupère l'historique pour l'interface
 * 
 * @param {Object} options - Options de filtrage
 * @returns {Array} Liste des snapshots
 */
function getAnalyticsHistoryForUI(options) {
  const result = getAnalyticsHistory(options);
  
  if (result.success) {
    return result.snapshots;
  } else {
    throw new Error(result.error);
  }
}
