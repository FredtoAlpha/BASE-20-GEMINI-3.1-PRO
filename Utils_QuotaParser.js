/**
 * ===================================================================
 * UTILITAIRES DE PARSING/FORMATAGE DES QUOTAS
 * ===================================================================
 * 
 * Fonctions partagées entre OPTI et LEGACY pour garantir
 * la cohérence du format OPTIONS dans _STRUCTURE
 */

/**
 * Formate les quotas au format standard pour _STRUCTURE
 * @param {Object} quotas - { "ITA": 6, "CHAV": 10, "ESP": 5 }
 * @returns {string} - "ITA=6,CHAV=10,ESP=5" (sans espaces)
 */
function formatQuotasToString_(quotas) {
  if (!quotas || typeof quotas !== 'object') {
    return '';
  }
  
  const pairs = [];
  Object.keys(quotas).forEach(function(key) {
    const value = parseInt(quotas[key]) || 0;
    if (value > 0) {
      pairs.push(key.toUpperCase() + '=' + value);
    }
  });
  
  return pairs.join(',');  // ✅ Pas d'espaces
}

/**
 * Parse une chaîne OPTIONS depuis _STRUCTURE
 * @param {string} optionsStr - "ITA=6,CHAV=10,ESP=5" ou "ITA=6, CHAV=10" (tolère espaces)
 * @returns {Object} - { "ITA": 6, "CHAV": 10, "ESP": 5 }
 */
function parseQuotasFromString_(optionsStr) {
  const quotas = {};
  
  if (!optionsStr || typeof optionsStr !== 'string') {
    return quotas;
  }
  
  // Split par virgule
  const pairs = optionsStr.split(',');
  
  pairs.forEach(function(pair) {
    // Trim pour enlever les espaces
    const trimmedPair = pair.trim();
    if (!trimmedPair) return;
    
    // Split par =
    const parts = trimmedPair.split('=');
    if (parts.length !== 2) return;
    
    // Extraire clé et valeur
    const key = parts[0].trim().toUpperCase();
    const value = parseInt(parts[1].trim()) || 0;
    
    if (key && value > 0) {
      quotas[key] = value;
    }
  });
  
  return quotas;
}

/**
 * Valide le format d'une chaîne OPTIONS
 * @param {string} optionsStr - Chaîne à valider
 * @returns {Object} - { valid: boolean, errors: string[] }
 */
function validateQuotasString_(optionsStr) {
  const result = {
    valid: true,
    errors: []
  };
  
  if (!optionsStr || typeof optionsStr !== 'string') {
    return result;  // Vide est valide
  }
  
  const pairs = optionsStr.split(',');
  
  pairs.forEach(function(pair, index) {
    const trimmedPair = pair.trim();
    if (!trimmedPair) {
      result.errors.push('Paire vide à l\'index ' + index);
      result.valid = false;
      return;
    }
    
    const parts = trimmedPair.split('=');
    if (parts.length !== 2) {
      result.errors.push('Format invalide: "' + trimmedPair + '" (attendu: CLÉ=VALEUR)');
      result.valid = false;
      return;
    }
    
    const key = parts[0].trim();
    const valueStr = parts[1].trim();
    
    if (!key) {
      result.errors.push('Clé vide dans: "' + trimmedPair + '"');
      result.valid = false;
    }
    
    const value = parseInt(valueStr);
    if (isNaN(value) || value < 0) {
      result.errors.push('Valeur invalide dans: "' + trimmedPair + '" (attendu: nombre positif)');
      result.valid = false;
    }
  });
  
  return result;
}

/**
 * Normalise une chaîne OPTIONS (enlève espaces superflus)
 * @param {string} optionsStr - "ITA=6, CHAV = 10 , ESP=5"
 * @returns {string} - "ITA=6,CHAV=10,ESP=5"
 */
function normalizeQuotasString_(optionsStr) {
  const quotas = parseQuotasFromString_(optionsStr);
  return formatQuotasToString_(quotas);
}

/**
 * Teste les fonctions utilitaires
 */
function testQuotasUtils_() {
  console.log('=== TEST QUOTAS UTILS ===');
  
  // Test 1: Format standard
  const test1 = parseQuotasFromString_('ITA=6,CHAV=10,ESP=5');
  console.log('Test 1:', JSON.stringify(test1));
  console.assert(test1.ITA === 6 && test1.CHAV === 10 && test1.ESP === 5, 'Test 1 échoué');
  
  // Test 2: Avec espaces
  const test2 = parseQuotasFromString_('ITA=6, CHAV = 10 , ESP=5');
  console.log('Test 2:', JSON.stringify(test2));
  console.assert(test2.ITA === 6 && test2.CHAV === 10 && test2.ESP === 5, 'Test 2 échoué');
  
  // Test 3: Format
  const test3 = formatQuotasToString_({ ITA: 6, CHAV: 10, ESP: 5 });
  console.log('Test 3:', test3);
  console.assert(test3 === 'ITA=6,CHAV=10,ESP=5', 'Test 3 échoué');
  
  // Test 4: Normalisation
  const test4 = normalizeQuotasString_('ITA=6, CHAV = 10 , ESP=5');
  console.log('Test 4:', test4);
  console.assert(test4 === 'ITA=6,CHAV=10,ESP=5', 'Test 4 échoué');
  
  // Test 5: Validation OK
  const test5 = validateQuotasString_('ITA=6,CHAV=10');
  console.log('Test 5:', JSON.stringify(test5));
  console.assert(test5.valid === true, 'Test 5 échoué');
  
  // Test 6: Validation KO
  const test6 = validateQuotasString_('ITA=6,CHAV=abc');
  console.log('Test 6:', JSON.stringify(test6));
  console.assert(test6.valid === false, 'Test 6 échoué');
  
  console.log('=== TOUS LES TESTS PASSÉS ===');
}
