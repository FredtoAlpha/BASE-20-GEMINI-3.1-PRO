// =============================================================================
// CONSOLE SCORES ELEVES -- Google Apps Script
// =============================================================================
// Ce script calcule 4 scores (ABS, COM, TRA, PART) a partir des exports Pronote
// et genere un tableau de synthese pour l'app de repartition.
//
// ARCHITECTURE DU GOOGLE SHEET :
// - Onglet "DATA_ABS"       -> Coller l'export Pronote des absences
// - Onglet "DATA_INCIDENTS"  -> Coller l'export Pronote des incidents/sanctions
// - Onglet "DATA_PUNITIONS"  -> Coller l'export Pronote des punitions
// - Onglet "DATA_NOTES"      -> Coller l'export Pronote des notes
// - Onglet "PARAMETRES"      -> Seuils modifiables pour chaque score
// - Onglet "SYNTHESE"        -> Tableau final avec les 4 scores
// =============================================================================

// =============================================================================
// MENU ET INITIALISATION
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Scores Élèves')
    .addItem('🏗️ Initialiser le classeur', 'initialiserClasseur')
    .addSeparator()
    .addItem('📋 Calculer ABS (Absences)', 'calculerABS')
    .addItem('🚨 Calculer COM (Comportement)', 'calculerCOM')
    .addItem('📚 Calculer TRA (Travail)', 'calculerTRA')
    .addItem('🗣️ Calculer PART (Participation)', 'calculerPART')
    .addSeparator()
    .addItem('🎯 Calculer TOUS les scores', 'calculerTous')
    .addItem('🔄 Générer SYNTHÈSE', 'genererSynthese')
    .addToUi();
}

// =============================================================================
// INITIALISATION -- Cree tous les onglets et les parametres par defaut
// =============================================================================

function initialiserClasseur() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var reponse = ui.alert(
    '🏗️ Initialiser le classeur',
    'Cela va créer les onglets nécessaires (DATA_ABS, DATA_INCIDENTS, DATA_PUNITIONS, DATA_NOTES, PARAMÈTRES, SYNTHÈSE).\n\nLes onglets existants ne seront pas écrasés.\n\nContinuer ?',
    ui.ButtonSet.YES_NO
  );
  if (reponse !== ui.Button.YES) return;

  // Créer les onglets s'ils n'existent pas
  var onglets = ['DATA_ABS', 'DATA_INCIDENTS', 'DATA_PUNITIONS', 'DATA_NOTES', 'PARAMÈTRES', 'SYNTHÈSE'];
  onglets.forEach(function(nom) {
    if (!ss.getSheetByName(nom)) {
      ss.insertSheet(nom);
    }
  });

  // Remplir l'onglet PARAMÈTRES
  creerParametres_(ss);

  // Mettre des instructions dans les onglets DATA
  var instructions = {
    'DATA_ABS': 'Collez ici l\'export Pronote des ABSENCES (avec les 2 lignes d\'en-tête)',
    'DATA_INCIDENTS': 'Collez ici l\'export Pronote des INCIDENTS (avec les 2 lignes d\'en-tête)',
    'DATA_PUNITIONS': 'Collez ici l\'export Pronote des PUNITIONS (avec la ligne d\'en-tête)',
    'DATA_NOTES': 'Collez ici l\'export Pronote des NOTES (avec les 2 lignes d\'en-tête)',
  };

  for (var nom in instructions) {
    var ws = ss.getSheetByName(nom);
    if (ws.getLastRow() === 0) {
      ws.getRange('A1').setValue(instructions[nom])
        .setFontStyle('italic').setFontColor('#808080');
    }
  }

  ui.alert('✅ Classeur initialisé !\n\nCollez vos exports Pronote dans les onglets DATA_*, puis utilisez le menu "Scores Élèves" pour calculer.');
}

function creerParametres_(ss) {
  var ws = ss.getSheetByName('PARAMÈTRES');
  ws.clear();

  // --- Titre ---
  ws.getRange('A1:F1').merge().setValue('⚙️ PARAMÈTRES DES SCORES')
    .setBackground('#2F5496').setFontColor('white').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center');
  ws.setRowHeight(1, 40);

  // --- ABS ---
  ws.getRange('A3').setValue('📋 SCORE ABS — Absences').setFontWeight('bold').setFontSize(12).setFontColor('#2F5496');
  ws.getRange('A4:F4').setValues([['Score', 'Libellé', 'DJ min', 'DJ max', 'NJ min', 'NJ max']]);
  ws.getRange('A4:F4').setBackground('#4472C4').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A5:F8').setValues([
    [4, 'Excellent', 0, 5, 0, 0],
    [3, 'Correct', 6, 13, 1, 2],
    [2, 'Fragile', 14, 25, 3, 5],
    [1, 'Critique', 26, 999, 6, 999]
  ]);
  ws.getRange('C5:F8').setBackground('#FFF2CC').setFontColor('#0000FF').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A5:A8').setHorizontalAlignment('center').setFontWeight('bold');

  ws.getRange('A9').setValue('Formule : Score DJ ×0.6 + Score NJ ×0.4, arrondi supérieur')
    .setFontStyle('italic').setFontColor('#808080');

  // --- COM ---
  ws.getRange('A11').setValue('🚨 SCORE COM — Comportement').setFontWeight('bold').setFontSize(12).setFontColor('#C00000');
  ws.getRange('A12:D12').setValues([['Score', 'Libellé', 'Points min', 'Points max']]);
  ws.getRange('A12:D12').setBackground('#C00000').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A13:D16').setValues([
    [4, 'Excellent', 0, 0],
    [3, 'Correct', 1, 5],
    [2, 'Fragile', 6, 20],
    [1, 'Critique', 21, 999]
  ]);
  ws.getRange('C13:D16').setBackground('#FFF2CC').setFontColor('#0000FF').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A13:A16').setHorizontalAlignment('center').setFontWeight('bold');

  ws.getRange('A17').setValue('Formule : Punitions ×1 + Incidents (pts gravité) ×3')
    .setFontStyle('italic').setFontColor('#808080');

  // --- TRA ---
  ws.getRange('A19').setValue('📚 SCORE TRA — Travail').setFontWeight('bold').setFontSize(12).setFontColor('#2F5496');
  ws.getRange('A20:D20').setValues([['Score', 'Libellé', 'Moy. min', 'Moy. max']]);
  ws.getRange('A20:D20').setBackground('#4472C4').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A21:D24').setValues([
    [4, 'Excellent', 15, 20],
    [3, 'Correct', 12, 14.99],
    [2, 'Fragile', 8, 11.99],
    [1, 'Critique', 0, 7.99]
  ]);
  ws.getRange('C21:D24').setBackground('#FFF2CC').setFontColor('#0000FF').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A21:A24').setHorizontalAlignment('center').setFontWeight('bold');

  // Pondérations
  ws.getRange('A26').setValue('⚖️ Pondérations par matière').setFontWeight('bold').setFontColor('#2F5496');
  ws.getRange('A27:C27').setValues([['Matière', 'Colonne CSV', 'Coeff']]);
  ws.getRange('A27:C27').setBackground('#4472C4').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');

  var matieres = [
    ['Français', 9, 4.5],
    ['Maths', 12, 3.5],
    ['Histoire-Géo', 10, 3],
    ['Anglais (moy.)', 3, 3],
    ['Espagnol/It. (moy.)', 13, 2.5],
    ['EPS', 7, 2],
    ['Phys.-Chimie', '20 ou 21', 1.5],
    ['SVT', '18 ou 19', 1.5],
    ['Technologie', '16 ou 17', 1.5],
    ['Arts Pla.', 6, 1],
    ['Musique', 8, 1],
    ['Latin', 22, 1]
  ];
  ws.getRange(28, 1, matieres.length, 3).setValues(matieres);
  ws.getRange(28, 3, matieres.length, 1).setBackground('#FFF2CC').setFontColor('#0000FF').setFontWeight('bold').setHorizontalAlignment('center');

  // --- PART ---
  ws.getRange('A41').setValue('🗣️ SCORE PART — Participation orale').setFontWeight('bold').setFontSize(12).setFontColor('#7030A0');
  ws.getRange('A42:D42').setValues([['Score', 'Libellé', 'Moy. min', 'Moy. max']]);
  ws.getRange('A42:D42').setBackground('#7030A0').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A43:D46').setValues([
    [4, 'Excellent', 15, 20],
    [3, 'Correct', 12, 14.99],
    [2, 'Fragile', 8, 11.99],
    [1, 'Critique', 0, 7.99]
  ]);
  ws.getRange('C43:D46').setBackground('#FFF2CC').setFontColor('#0000FF').setFontWeight('bold').setHorizontalAlignment('center');
  ws.getRange('A43:A46').setHorizontalAlignment('center').setFontWeight('bold');

  ws.getRange('A47').setValue('Oral Anglais (col 5 CSV) + Oral LV2 (col 15 CSV) → moyenne')
    .setFontStyle('italic').setFontColor('#808080');

  // Largeurs
  ws.setColumnWidth(1, 180);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 100);
  ws.setColumnWidth(5, 100);
  ws.setColumnWidth(6, 100);
}

// =============================================================================
// MODULE ABS -- Calcul du score d'assiduite
// =============================================================================

function calculerABS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsData = ss.getSheetByName('DATA_ABS');
  var wsParam = ss.getSheetByName('PARAMÈTRES');

  if (!wsData || wsData.getLastRow() < 3) {
    SpreadsheetApp.getUi().alert('❌ L\'onglet DATA_ABS est vide.\nCollez d\'abord l\'export Pronote des absences.');
    return;
  }

  // Lire les seuils
  var seuilsDJ = lireSeuilsABS_(wsParam, 'C', 'D', 5, 8);  // DJ min/max
  var seuilsNJ = lireSeuilsABS_(wsParam, 'E', 'F', 5, 8);  // NJ min/max

  // Lire les données brutes (on saute les 2 lignes d'en-tête)
  var data = wsData.getDataRange().getValues();
  var eleves = {};

  for (var i = 2; i < data.length; i++) {
    var nom = String(data[i][0]).trim();
    if (!nom) continue;
    var classe = String(data[i][1]).trim();
    var djBulletin = parseNote_(data[i][9]);  // Colonne J = demi-journées bulletin
    var justifiee = String(data[i][10]).trim();

    if (!eleves[nom]) {
      eleves[nom] = { classe: classe, djTotal: 0, nonJustifiees: 0 };
    }
    if (djBulletin !== null) eleves[nom].djTotal += djBulletin;
    if (justifiee === 'Non') eleves[nom].nonJustifiees++;
  }

  // Calculer les scores
  var resultats = [];
  for (var nom in eleves) {
    var e = eleves[nom];
    var scoreDJ = attribuerScore_(e.djTotal, seuilsDJ);
    var scoreNJ = attribuerScore_(e.nonJustifiees, seuilsNJ);
    var scoreABS = Math.ceil(scoreDJ * 0.6 + scoreNJ * 0.4);
    resultats.push({
      nom: nom, classe: e.classe, dj: Math.round(e.djTotal * 10) / 10,
      nj: e.nonJustifiees, scoreDJ: scoreDJ, scoreNJ: scoreNJ, scoreABS: scoreABS
    });
  }

  resultats.sort(function(a, b) {
    return a.classe.localeCompare(b.classe) || a.nom.localeCompare(b.nom);
  });

  // Écrire dans SYNTHÈSE
  ecrireScoreDansSynthese_(ss, resultats, 'ABS');

  SpreadsheetApp.getUi().alert('✅ Score ABS calculé pour ' + resultats.length + ' élèves !');
}

// =============================================================================
// MODULE COM -- Calcul du score de comportement
// =============================================================================

function calculerCOM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsInc = ss.getSheetByName('DATA_INCIDENTS');
  var wsPun = ss.getSheetByName('DATA_PUNITIONS');
  var wsParam = ss.getSheetByName('PARAMÈTRES');

  if ((!wsInc || wsInc.getLastRow() < 3) && (!wsPun || wsPun.getLastRow() < 2)) {
    SpreadsheetApp.getUi().alert('❌ Les onglets DATA_INCIDENTS et DATA_PUNITIONS sont vides.');
    return;
  }

  // Lire les seuils COM
  var seuilsCOM = lireSeuils_(wsParam, 'C', 'D', 13, 16);

  // Lire les punitions
  var punitions = {};
  if (wsPun && wsPun.getLastRow() >= 2) {
    var dataPun = wsPun.getDataRange().getValues();
    for (var i = 1; i < dataPun.length; i++) {
      var nom = String(dataPun[i][1]).trim();
      if (!nom) continue;
      var nb = parseInt(dataPun[i][0]) || 0;
      var classe = String(dataPun[i][2]).trim();
      punitions[nom] = { nb: nb, classe: classe };
    }
  }

  // Lire les incidents
  var incidents = {};
  if (wsInc && wsInc.getLastRow() >= 3) {
    var dataInc = wsInc.getDataRange().getValues();
    for (var i = 2; i < dataInc.length; i++) {
      var nom = String(dataInc[i][1]).trim();
      if (!nom) continue;
      var classe = String(dataInc[i][0]).trim();
      var gravStr = String(dataInc[i][7]).trim();
      var grav = 1;
      if (gravStr && gravStr.indexOf('/') > -1) {
        grav = parseInt(gravStr.split('/')[0]) || 1;
      }

      if (!incidents[nom]) {
        incidents[nom] = { classe: '', nbInc: 0, ptsGrav: 0 };
      }
      if (classe) incidents[nom].classe = classe;
      incidents[nom].nbInc++;
      incidents[nom].ptsGrav += grav;
    }
  }

  // Fusionner
  var tousNoms = {};
  for (var nom in punitions) tousNoms[nom] = true;
  for (var nom in incidents) tousNoms[nom] = true;

  var resultats = [];
  for (var nom in tousNoms) {
    var ptsPun = punitions[nom] ? punitions[nom].nb : 0;
    var ptsInc = incidents[nom] ? incidents[nom].ptsGrav * 3 : 0;
    var total = ptsPun + ptsInc;
    var classe = (punitions[nom] ? punitions[nom].classe : '') || (incidents[nom] ? incidents[nom].classe : '');
    var scoreCOM = attribuerScoreInverse_(total, seuilsCOM);

    resultats.push({
      nom: nom, classe: classe, total: total, scoreCOM: scoreCOM
    });
  }

  resultats.sort(function(a, b) {
    return a.classe.localeCompare(b.classe) || a.nom.localeCompare(b.nom);
  });

  ecrireScoreDansSynthese_(ss, resultats, 'COM');

  SpreadsheetApp.getUi().alert('✅ Score COM calculé pour ' + resultats.length + ' élèves !');
}

// =============================================================================
// MODULE TRA -- Calcul du score de travail (moyenne ponderee)
// =============================================================================

function calculerTRA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsData = ss.getSheetByName('DATA_NOTES');
  var wsParam = ss.getSheetByName('PARAMÈTRES');

  if (!wsData || wsData.getLastRow() < 3) {
    SpreadsheetApp.getUi().alert('❌ L\'onglet DATA_NOTES est vide.\nCollez d\'abord l\'export Pronote des notes.');
    return;
  }

  // Lire les seuils TRA
  var seuilsTRA = lireSeuils_(wsParam, 'C', 'D', 21, 24);

  // Lire les pondérations depuis PARAMÈTRES (lignes 28 à 39)
  // Structure fixe basée sur le CSV Pronote
  // Colonnes CSV (0-indexed) :
  //   2=Anglais moy, 4=Anglais oral, 5=ArtsPla, 6=EPS, 7=Musique
  //   8=Français, 9=HG, 11=Maths, 12=Esp moy, 14=Esp oral
  //   15ou16=Techno, 17ou18=SVT, 19ou20=PhCh, 21=Latin

  var matieresTRA = [
    { nom: 'Français', cols: [8], coeff: 4.5 },
    { nom: 'Maths', cols: [11], coeff: 3.5 },
    { nom: 'Histoire-Géo', cols: [9], coeff: 3.0 },
    { nom: 'Anglais', cols: [2], coeff: 3.0 },
    { nom: 'Espagnol/It.', cols: [12], coeff: 2.5 },
    { nom: 'EPS', cols: [6], coeff: 2.0 },
    { nom: 'Phys.-Chimie', cols: [19, 20], coeff: 1.5 },
    { nom: 'SVT', cols: [17, 18], coeff: 1.5 },
    { nom: 'Technologie', cols: [15, 16], coeff: 1.5 },
    { nom: 'Arts Pla.', cols: [5], coeff: 1.0 },
    { nom: 'Musique', cols: [7], coeff: 1.0 },
    { nom: 'Latin', cols: [21], coeff: 1.0 },
  ];

  // Essayer de lire les coefficients personnalisés depuis PARAMÈTRES
  var paramData = wsParam.getRange('C28:C39').getValues();
  for (var i = 0; i < matieresTRA.length && i < paramData.length; i++) {
    var customCoeff = parseFloat(paramData[i][0]);
    if (!isNaN(customCoeff) && customCoeff > 0) {
      matieresTRA[i].coeff = customCoeff;
    }
  }

  var data = wsData.getDataRange().getValues();
  var resultats = [];

  for (var i = 2; i < data.length; i++) {
    var nom = String(data[i][0]).trim();
    if (!nom) continue;
    var classe = String(data[i][1]).trim();

    var totalPts = 0;
    var totalCoeff = 0;

    for (var m = 0; m < matieresTRA.length; m++) {
      var note = null;
      for (var c = 0; c < matieresTRA[m].cols.length; c++) {
        var colIdx = matieresTRA[m].cols[c];
        if (colIdx < data[i].length) {
          var n = parseNote_(data[i][colIdx]);
          if (n !== null) { note = n; break; }
        }
      }
      if (note !== null) {
        totalPts += note * matieresTRA[m].coeff;
        totalCoeff += matieresTRA[m].coeff;
      }
    }

    var moyPond = totalCoeff > 0 ? Math.round(totalPts / totalCoeff * 100) / 100 : null;
    var scoreTRA = moyPond !== null ? attribuerScoreNote_(moyPond, seuilsTRA) : null;

    resultats.push({
      nom: nom, classe: classe, moyPond: moyPond, scoreTRA: scoreTRA
    });
  }

  resultats.sort(function(a, b) {
    return a.classe.localeCompare(b.classe) || a.nom.localeCompare(b.nom);
  });

  ecrireScoreDansSynthese_(ss, resultats, 'TRA');

  SpreadsheetApp.getUi().alert('✅ Score TRA calculé pour ' + resultats.length + ' élèves !');
}

// =============================================================================
// MODULE PART -- Calcul du score de participation orale
// =============================================================================

function calculerPART() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsData = ss.getSheetByName('DATA_NOTES');
  var wsParam = ss.getSheetByName('PARAMÈTRES');

  if (!wsData || wsData.getLastRow() < 3) {
    SpreadsheetApp.getUi().alert('❌ L\'onglet DATA_NOTES est vide.\nCollez d\'abord l\'export Pronote des notes.');
    return;
  }

  var seuilsPART = lireSeuils_(wsParam, 'C', 'D', 43, 46);

  var data = wsData.getDataRange().getValues();
  var resultats = [];

  for (var i = 2; i < data.length; i++) {
    var nom = String(data[i][0]).trim();
    if (!nom) continue;
    var classe = String(data[i][1]).trim();

    var oralAng = parseNote_(data[i][4]);   // Col 5 (index 4) = oral anglais
    var oralLV2 = parseNote_(data[i][14]);  // Col 15 (index 14) = oral LV2

    var notes = [];
    if (oralAng !== null) notes.push(oralAng);
    if (oralLV2 !== null) notes.push(oralLV2);

    var moyOral = notes.length > 0 ? Math.round(notes.reduce(function(a,b){return a+b;}, 0) / notes.length * 100) / 100 : null;
    var scorePART = moyOral !== null ? attribuerScoreNote_(moyOral, seuilsPART) : null;

    resultats.push({
      nom: nom, classe: classe, moyOral: moyOral, scorePART: scorePART
    });
  }

  resultats.sort(function(a, b) {
    return a.classe.localeCompare(b.classe) || a.nom.localeCompare(b.nom);
  });

  ecrireScoreDansSynthese_(ss, resultats, 'PART');

  SpreadsheetApp.getUi().alert('✅ Score PART calculé pour ' + resultats.length + ' élèves !');
}

// =============================================================================
// CALCULER TOUS + GENERER SYNTHESE
// =============================================================================

function calculerTous() {
  calculerABS();
  calculerCOM();
  calculerTRA();
  calculerPART();
  formaterSynthese_();
  SpreadsheetApp.getUi().alert('🎯 Tous les scores ont été calculés et la synthèse est prête !');
}

function genererSynthese() {
  formaterSynthese_();
  SpreadsheetApp.getUi().alert('✅ Synthèse mise en forme !');
}

// =============================================================================
// ECRITURE DANS LA SYNTHESE
// =============================================================================

function ecrireScoreDansSynthese_(ss, resultats, typeScore) {
  var ws = ss.getSheetByName('SYNTHÈSE');
  if (!ws) {
    ws = ss.insertSheet('SYNTHÈSE');
  }

  // Colonnes : A=Nom, B=Classe, C=ABS, D=COM, E=TRA, F=PART
  var colIndex = { 'ABS': 3, 'COM': 4, 'TRA': 5, 'PART': 6 };
  var col = colIndex[typeScore];

  // Si la synthèse est vide, créer les en-têtes
  if (ws.getLastRow() < 2 || ws.getRange('A2').getValue() === '') {
    ws.clear();
    ws.getRange('A1:F1').merge().setValue('🎯 SYNTHÈSE DES SCORES — App Répartition')
      .setBackground('#00B050').setFontColor('white').setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center');
    ws.setRowHeight(1, 40);
    ws.getRange('A2:F2').setValues([['Nom', 'Classe', 'ABS', 'COM', 'TRA', 'PART']]);
    ws.getRange('A2:F2').setBackground('#00B050').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center');
    ws.setFrozenRows(2);
  }

  // Construire un index des noms existants dans la synthèse
  var existingData = ws.getLastRow() >= 3 ? ws.getRange(3, 1, ws.getLastRow() - 2, 2).getValues() : [];
  var nomIndex = {};
  for (var i = 0; i < existingData.length; i++) {
    nomIndex[String(existingData[i][0]).trim()] = i + 3; // numéro de ligne
  }

  // Écrire les résultats
  for (var i = 0; i < resultats.length; i++) {
    var r = resultats[i];
    var ligne = nomIndex[r.nom];

    if (!ligne) {
      // Nouvel élève → ajouter une ligne
      ligne = ws.getLastRow() + 1;
      ws.getRange(ligne, 1).setValue(r.nom);
      ws.getRange(ligne, 2).setValue(r.classe).setHorizontalAlignment('center');
      nomIndex[r.nom] = ligne;
    }

    // Écrire le score
    var score = r['score' + typeScore] || r.scoreABS || r.scoreCOM || r.scoreTRA || r.scorePART;
    if (score !== null && score !== undefined) {
      var cell = ws.getRange(ligne, col);
      cell.setValue(score).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12);

      // Couleur selon le score
      var colors = { 4: '#C6EFCE', 3: '#D6E4F0', 2: '#FFF2CC', 1: '#FFC7CE' };
      var fontColors = { 4: '#006100', 3: '#2F5496', 2: '#BF8F00', 1: '#C00000' };
      cell.setBackground(colors[score] || '#FFFFFF');
      cell.setFontColor(fontColors[score] || '#000000');
    } else {
      ws.getRange(ligne, col).setValue('—').setBackground('#E0E0E0').setFontColor('#A0A0A0')
        .setHorizontalAlignment('center');
    }
  }

  // Trier par Classe puis Nom
  if (ws.getLastRow() >= 4) {
    ws.getRange(3, 1, ws.getLastRow() - 2, 6).sort([
      { column: 2, ascending: true },
      { column: 1, ascending: true }
    ]);
  }

  // Ajuster les largeurs
  ws.setColumnWidth(1, 220);
  ws.setColumnWidth(2, 80);
  for (var c = 3; c <= 6; c++) ws.setColumnWidth(c, 70);
}

function formaterSynthese_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('SYNTHÈSE');
  if (!ws || ws.getLastRow() < 3) return;

  // Bordures sur tout le tableau
  var range = ws.getRange(2, 1, ws.getLastRow() - 1, 6);
  range.setBorder(true, true, true, true, true, true, '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);

  // Alternance de couleur sur les lignes (colonnes Nom et Classe uniquement)
  for (var i = 3; i <= ws.getLastRow(); i++) {
    if (i % 2 === 1) {
      ws.getRange(i, 1, 1, 2).setBackground('#F5F5F5');
    }
  }
}

// =============================================================================
// FONCTIONS UTILITAIRES
// =============================================================================

/**
 * Parse une note (gère les virgules françaises, "Abs", "Disp", etc.)
 */
function parseNote_(val) {
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  if (s === '' || s === 'Abs' || s === 'Disp' || s === 'NE' || s === 'NN' || s === '—') return null;
  s = s.replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/**
 * Lire les seuils ABS depuis l'onglet PARAMETRES (2 colonnes min/max)
 */
function lireSeuilsABS_(ws, colMin, colMax, ligneDebut, ligneFin) {
  var seuils = [];
  for (var i = ligneDebut; i <= ligneFin; i++) {
    seuils.push({
      score: ws.getRange('A' + i).getValue(),
      min: ws.getRange(colMin + i).getValue(),
      max: ws.getRange(colMax + i).getValue()
    });
  }
  return seuils;
}

/**
 * Lire les seuils génériques (COM)
 */
function lireSeuils_(ws, colMin, colMax, ligneDebut, ligneFin) {
  var seuils = [];
  for (var i = ligneDebut; i <= ligneFin; i++) {
    seuils.push({
      score: parseInt(ws.getRange('A' + i).getValue()),
      min: parseFloat(ws.getRange(colMin + i).getValue()),
      max: parseFloat(ws.getRange(colMax + i).getValue())
    });
  }
  return seuils;
}

/**
 * Attribuer un score basé sur une valeur et des seuils min/max
 * Pour ABS : la valeur doit être entre min et max pour obtenir le score
 */
function attribuerScore_(valeur, seuils) {
  for (var i = 0; i < seuils.length; i++) {
    if (valeur >= seuils[i].min && valeur <= seuils[i].max) {
      return seuils[i].score;
    }
  }
  return 1; // Par défaut, score le plus bas
}

/**
 * Attribuer un score pour COM (inversé : plus les points sont élevés, plus le score est bas)
 */
function attribuerScoreInverse_(valeur, seuils) {
  for (var i = 0; i < seuils.length; i++) {
    if (valeur >= seuils[i].min && valeur <= seuils[i].max) {
      return seuils[i].score;
    }
  }
  return 1;
}

/**
 * Attribuer un score pour TRA/PART (basé sur la moyenne)
 */
function attribuerScoreNote_(moyenne, seuils) {
  for (var i = 0; i < seuils.length; i++) {
    if (moyenne >= seuils[i].min && moyenne <= seuils[i].max) {
      return seuils[i].score;
    }
  }
  return 1;
}
