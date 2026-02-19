# AUDIT COMPLET — BASE-19-REFACTOR-B
## Date : 2026-02-12 | Branche : `claude/optimize-engine-restart-iEcOL`

**Périmètre** : 90 fichiers (61 .js + 25 .html + configs), ~67 500 lignes de code
**Méthodologie** : analyse statique exhaustive (grep patterns, lecture de code, mapping des dépendances)

---

# TABLE DES MATIÈRES

1. [RÉSUMÉ EXÉCUTIF](#1-résumé-exécutif)
2. [SÉCURITÉ](#2-sécurité)
3. [FONCTIONS DUPLIQUÉES](#3-fonctions-dupliquées)
4. [ARCHITECTURE & DETTE TECHNIQUE](#4-architecture--dette-technique)
5. [PERFORMANCE](#5-performance)
6. [FRONTEND & DÉPENDANCES CDN](#6-frontend--dépendances-cdn)
7. [QUALITÉ DE CODE](#7-qualité-de-code)
8. [PLAN D'ACTION PRIORISÉ](#8-plan-daction-priorisé)

---

# 1. RÉSUMÉ EXÉCUTIF

| Catégorie | Critique | Haute | Moyenne | Basse | Total |
|-----------|----------|-------|---------|-------|-------|
| Sécurité | 3 | 2 | 2 | 1 | **8** |
| Duplications | 2 | 4 | 3 | — | **9** |
| Architecture | 1 | 3 | 2 | — | **6** |
| Performance | — | 2 | 3 | 1 | **6** |
| Frontend | 1 | 2 | 2 | 1 | **6** |
| Qualité code | — | 2 | 3 | 2 | **7** |
| **TOTAL** | **7** | **15** | **15** | **5** | **42** |

**Verdict** : L'application est fonctionnelle mais présente **7 problèmes critiques** qui doivent être traités en priorité. Les duplications massives de fonctions (>30 fonctions dupliquées entre 3 fichiers) constituent le risque n°1 de régression.

---

# 2. SÉCURITÉ

## SEC-01 — CRITIQUE : Mot de passe admin en clair dans le code source

**Fichiers concernés :**
- `Config.js:19` → `ADMIN_PASSWORD_DEFAULT: "admin123"`
- `ConsolePilotageV3.html:359` → `<input type="password" id="pwd" value="admin123">`
- `ConsolePilotageV3_Server.js:175` → `const expectedPassword = config.ADMIN_PASSWORD || config.ADMIN_PASSWORD_DEFAULT || "admin123"`
- `InterfaceV2_CoreScript.html:3785` → `adminPassword: 'admin123'`

**Impact** : Tout utilisateur ayant accès au code source (ou à la console navigateur) peut voir le mot de passe admin par défaut. Le fallback en dur `"admin123"` est présent à 4 endroits différents.

**Correctif recommandé** :
1. Supprimer tous les fallbacks `"admin123"` du code
2. Forcer la définition du mot de passe dans `_CONFIG` à l'initialisation
3. Hasher le mot de passe (SHA-256 via `Utilities.computeDigest()`)
4. Comparer les hashes côté serveur uniquement

---

## SEC-02 — CRITIQUE : Mot de passe admin envoyé au client (fuite côté frontend)

**Fichier** : `InterfaceV2_CoreScript.html:3586-3597`

```javascript
.withSuccessHandler(function(configPassword){
  resolve(pwd === configPassword || (window.CONFIG && pwd === CONFIG.adminPassword));
})
.getAdminPasswordFromConfig();
```

**Impact** : La fonction `getAdminPasswordFromConfig()` envoie le mot de passe **en clair** au navigateur du client. N'importe quel utilisateur peut ouvrir la console et appeler `google.script.run.getAdminPasswordFromConfig()` pour récupérer le mot de passe.

**Correctif recommandé** : Créer une fonction serveur `verifierMotDePasseAdmin(password)` qui retourne `true/false` (elle existe déjà dans `Code.js:946`). Utiliser exclusivement celle-ci côté client.

---

## SEC-03 — CRITIQUE : Fonction `getAdminPasswordFromConfig` dupliquée avec comportement différent

**Fichiers :**
- `Code.js:918` → lit `configSheet.getRange('B2').getValue()`
- `AdminPasswordHelper.js:5` → lit `configSheet.getRange('B3').getValue()`

**Impact** : Deux fonctions avec le **même nom** lisent des **cellules différentes** (B2 vs B3). En GAS, la dernière chargée écrase l'autre. Le comportement est **non-déterministe** selon l'ordre de chargement des fichiers. Un utilisateur pourrait croire avoir changé son mot de passe alors que la mauvaise cellule est lue.

**Correctif recommandé** : Supprimer `AdminPasswordHelper.js` et ne garder qu'une seule implémentation dans `Code.js`.

---

## SEC-04 — HAUTE : innerHTML sans sanitisation (risque XSS)

**~80+ occurrences** de `innerHTML = ...` à travers les fichiers HTML. Exemples à risque :

| Fichier | Ligne | Code dangereux |
|---------|-------|---------------|
| `ConsolePilotageV3.html` | 704 | `txt.innerHTML = msg` (msg vient de données utilisateur) |
| `InterfaceV2.html` | 1994 | `content.innerHTML = '...Erreur: ' + error.message + '...'` |
| `InterfaceV2.html` | 1592 | `modal.innerHTML = ...` (construction dynamique) |
| `InterfaceV2_CoreScript.html` | 5132 | `header.innerHTML = ...escapeHtml(title)...` (OK — sanitisé) |

**Note** : Certains usages utilisent `escapeHtml()` (bien), mais la majorité ne le font pas.

**Impact** : Si des données élèves contiennent du HTML/JS malveillant (ex: nom `<script>alert(1)</script>`), elles seront exécutées dans le navigateur de l'administrateur.

**Correctif recommandé** : Remplacer `innerHTML` par `textContent` partout où du texte simple est injecté. Pour le HTML structuré, toujours utiliser `escapeHtml()` sur les données dynamiques.

---

## SEC-05 — HAUTE : `document.write()` dans des fenêtres popup

**Fichiers :**
- `GroupsInterfaceV4.html:4313` → `printWindow.document.write(content)`
- `InterfaceV2_CoreScript.html:6983` → `groupsWindow.document.write(html)`

**Impact** : `document.write()` dans un nouveau contexte de fenêtre peut permettre une injection XSS si le contenu inclut des données utilisateur non-sanitisées.

**Correctif recommandé** : Utiliser des templates HTML avec `textContent` pour les données dynamiques.

---

## SEC-06 — MOYENNE : Logs serveur exposant le mot de passe

**Fichiers :**
- `DiagnosticConfig.js:195` → `Logger.log('ADMIN_PASSWORD = "${rapport.config.ADMIN_PASSWORD}"')`
- `DiagnosticConfig.js:205` → `Logger.log('config.ADMIN_PASSWORD = "${config.ADMIN_PASSWORD}"')`
- `Code.js:964` → `Logger.log(... input(${inputPassword.length}) vs config(${adminPassword.length})...)`

**Impact** : Le mot de passe admin apparaît en clair dans les logs Stackdriver. Les longueurs du mot de passe sont aussi loguées (facilite le brute-force).

**Correctif recommandé** : Ne jamais loguer de mot de passe, même partiellement. Logger uniquement `"[REDACTED]"`.

---

## SEC-07 — MOYENNE : Accès web app ouvert à tous

**Fichier** : `appsscript.json`
```json
"webapp": {
  "executeAs": "USER_DEPLOYING",
  "access": "ANYONE"
}
```

**Impact** : N'importe qui ayant l'URL peut accéder à l'application. Combiné avec le mot de passe admin faible/en clair, cela crée un vecteur d'attaque.

**Correctif recommandé** : Restreindre à `"ANYONE_WITHIN_DOMAIN"` si possible, ou au minimum documenter ce choix d'architecture.

---

## SEC-08 — BASSE : `insertAdjacentHTML` sans vérification

**Fichier** : `InterfaceV2_CoreScript.html:94` → `document.body.insertAdjacentHTML('beforeend', progressHTML)`

**Impact** : Risque limité car `progressHTML` est construit en code, pas depuis des données utilisateur. Surveiller si ce pattern est réutilisé avec des données dynamiques.

---

# 3. FONCTIONS DUPLIQUÉES

C'est le problème le plus grave pour la maintenabilité. **Plus de 30 fonctions** sont dupliquées entre `Orchestration_V14I.js`, `App.Core.js`, `App.Context.js`, `App.CacheManager.js`, et d'autres fichiers.

## DUP-01 — CRITIQUE : Fonctions identiques entre Orchestration_V14I.js et App.*.js (collision de noms GAS)

En Google Apps Script, **toutes les fonctions sont globales**. Si deux fichiers définissent `function foo()`, la dernière écrase la première. Les fichiers suivants définissent les **mêmes fonctions** :

### Fonctions dupliquées Orchestration_V14I.js ↔ App.Context.js (MÊME NOM EXACT) :
| Fonction | Orchestration_V14I.js | App.Context.js |
|----------|----------------------|----------------|
| `makeCtxFromSourceSheets_()` | ligne 32 | ligne 30 |
| `readModeFromUI_()` | ligne 465 | ligne 174 |
| `readNiveauxFromUI_()` | ligne 490 | ligne 194 |
| `readQuotasFromUI_()` | ligne 504 | ligne 211 |
| `readTargetsFromUI_()` | ligne 667 | ligne 239 |
| `readParityToleranceFromUI_()` | ligne 737 | ligne 264 |
| `readMaxSwapsFromUI_()` | ligne 753 | ligne 274 |
| `readClassAuthorizationsFromUI_()` | ligne 770 | ligne 288 |
| `readSourceToDestMapping_()` | ligne 604 | ligne 307 |
| `buildClassOffers_(ctx)` | ligne 2731 | ligne 380 |
| `buildOfferWithQuotas_(ctx)` | ligne 3073 | ligne 418 |
| `computeAllow_(eleve, classOffers)` | ligne 2762 | ligne 458 |

**→ 12 fonctions en collision directe**

### Fonctions dupliquées Orchestration_V14I.js ↔ App.CacheManager.js :
| Fonction | Orchestration_V14I.js | App.CacheManager.js |
|----------|----------------------|---------------------|
| `readElevesFromCache_(ctx)` | ligne 914 | ligne 31 |
| `readElevesFromSelectedMode_(ctx)` | ligne 866 | ligne 66 |
| `forceCacheInUIAndReload_(ctx)` | ligne 788 | ligne 121 |
| `setInterfaceModeCACHE_(ctx)` | ligne 814 | ligne 152 |
| `activateFirstCacheTabIfAny_(ctx)` | ligne 831 | ligne 174 |
| `triggerUIReloadFromCACHE_()` | ligne 843 | ligne 190 |
| `openCacheTabs_(ctx)` | ligne 3010 | ligne 206 |
| `computeMobilityFlags_(ctx)` | ligne 2795 | ligne 273 |
| `auditCacheAgainstStructure_(ctx)` | ligne 3106 | ligne 406 |

**→ 9 fonctions en collision directe**

### Fonctions dupliquées Orchestration_V14I.js ↔ App.Core.js :
| Fonction | Orchestration_V14I.js | App.Core.js |
|----------|----------------------|-------------|
| `_u_(v)` | ligne 2710 | ligne 127 |
| `_arr(v)` | ligne 2711 | ligne 143 |
| `parseCodes_(rowObj)` | ligne 2781 | ligne 161 |
| `findEleveByGenre_()` | ligne 1884 | ligne 187 |
| `calculateClassMetric_LEGACY_()` | ligne 2182 | ligne 222 |
| `computeClassState_LEGACY_()` | ligne 2620 | ligne 245 |
| `simulateSwapState_LEGACY_()` | ligne 2660 | ligne 296 |
| `calculateClassScores_LEGACY_()` | ligne 2165 | ligne 341 |
| `isPlacementLV2OPTOK_()` | ligne 1995 | ligne 374 |
| `isMoveAllowed_()` | ligne 2316 | ligne 400 |
| `isEleveMobile_LEGACY_()` | ligne 2527 | ligne 462 |
| `eligibleForSwap_LEGACY_()` | ligne 2376 | ligne 495 |
| `isSwapValid_LEGACY_()` | ligne 2384 | ligne 512 |
| `computeCountsFromState_()` | ligne 2275 | ligne 559 |
| `computeMobilityStats_LEGACY_()` | ligne 2424 | ligne 596 |
| `calculateSwapScore_LEGACY_()` | ligne 2561 | ligne 632 |
| `logLine()` | ligne 2699 | ligne 701 |

**→ 17 fonctions en collision directe**

### Aussi dupliquées avec Mobility_System.js :
| Fonction | Orchestration_V14I.js | Mobility_System.js |
|----------|----------------------|--------------------|
| `_u_()` | ligne 2710 | ligne 30 |
| `buildClassOffers_()` | ligne 2731 | ligne 57 |
| `computeAllow_()` | ligne 2762 | ligne 91 |
| `parseCodes_()` | ligne 2781 | ligne 112 |
| `computeMobilityFlags_()` | ligne 2795 | ligne 131 |

**→ 5 fonctions en collision directe**

**Impact total** : **~38 fonctions** sont définies dans 2 ou 3 fichiers simultanément. Le comportement dépend de l'ordre de chargement GAS (non garanti). Cela signifie que **modifier une fonction dans un fichier n'aura peut-être aucun effet** car c'est la copie de l'autre fichier qui est exécutée.

---

## DUP-02 — CRITIQUE : `getClassesData` définie 3 fois

| Fichier | Ligne | Signature |
|---------|-------|-----------|
| `Backend_Core.js` | 33 | `getClassesData(mode)` |
| `Backend_Eleves.js` | 111 | `getClassesData(mode = 'source')` |
| `Code.js` | 704 | `getClassesData(mode = 'source')` |

**Impact** : Quelle version le frontend appelle-t-il ? Impossible à déterminer sans connaître l'ordre de chargement GAS.

---

## DUP-03 — HAUTE : `logLine` définie 6 fois

| Fichier | Ligne |
|---------|-------|
| `App.Core.js` | 701 |
| `Logger.js` | 418 |
| `Orchestration_V14I.js` | 2699 |
| `Phase4_Ultimate.js` | 987 |
| `BASEOPTI_System.js` | 22 (nested) |
| `NiveauxDynamiques.js` | 201 (nested) |

---

## DUP-04 — HAUTE : `logParityDecision` définie 3 fois

| Fichier | Ligne |
|---------|-------|
| `Phases_BASEOPTI_V3_COMPLETE.js` | **899 ET 936** (deux fois dans le même fichier !) |
| `Phase3_PariteAdaptive_V3.js` | 311 (suffixée `_V3`) |

---

## DUP-05 — HAUTE : `findLeastPopulatedClass` en 4 variantes

| Fichier | Ligne | Nom |
|---------|-------|-----|
| `LEGACY_Phase2_DissoAsso.js` | 264 | `findLeastPopulatedClass_LEGACY` |
| `LEGACY_Phase3_Parite.js` | 447 | `findLeastPopulatedClass_Phase3` |
| `Phases_BASEOPTI_V3_COMPLETE.js` | 319 | `findLeastPopulatedClass_V3` |
| `BASEOPTI_Architecture_V3.js` | 350 | `findLeastPopulatedClass_V3` (doublon !) |

Note : `findLeastPopulatedClass_V3` est définie **deux fois** avec le même nom (collision GAS).

---

## DUP-06 — HAUTE : `findClassWithoutCodeD` en 3 variantes dont 2 identiques

| Fichier | Ligne | Nom |
|---------|-------|-----|
| `LEGACY_Phase2_DissoAsso.js` | 291 | `findClassWithoutCodeD_LEGACY` |
| `Phases_BASEOPTI_V3_COMPLETE.js` | 452 | `findClassWithoutCodeD_V3` |
| `BASEOPTI_Architecture_V3.js` | 389 | `findClassWithoutCodeD_V3` (doublon !) |

---

## DUP-07 — MOYENNE : `validateDISSOConstraints` en 3 variantes

| Fichier | Ligne | Nom |
|---------|-------|-----|
| `LEGACY_Phase3_Parite.js` | 550 | `validateDISSOConstraints_Phase3` |
| `Phases_BASEOPTI_V3_COMPLETE.js` | 738 | `validateDISSOConstraints_V3_` |
| `Phase4_Ultimate.js` | 995 | `validateDISSOConstraints_Ultimate` |

---

## DUP-08 — MOYENNE : Backend fonctions dupliquées en triple

| Fonction | `Backend_Core.js` | `Backend_Eleves.js` | `Code.js` |
|----------|-------------------|---------------------|-----------|
| `getLastCacheInfo()` | ligne 106 | ligne 313 | ligne 718 |
| `getBridgeContextAndClear()` | ligne 147 | ligne 428 | ligne 742 |
| `saveCacheData()` | ligne 185 | ligne 339 | ligne 765 |
| `loadCacheData()` | ligne 233 | ligne 353 | ligne 855 |
| `saveElevesSnapshot()` | ligne 249 | ligne 288 | ligne 873 |

**→ 5 fonctions × 3 copies = 15 définitions pour 5 fonctions**

---

## DUP-09 — MOYENNE : `openCacheTabs_` dupliquée

| Fichier | Ligne |
|---------|-------|
| `Orchestration_V14I.js` | 3010 |
| `Orchestration_V14I_Stream.js` | 205 |
| `App.CacheManager.js` | 206 |

---

# 4. ARCHITECTURE & DETTE TECHNIQUE

## ARCH-01 — CRITIQUE : Orchestration_V14I.js est un "God file" (3 352 lignes)

`Orchestration_V14I.js` contient :
- Le pipeline complet d'optimisation
- Les fonctions de contexte (`makeCtx`, `readMode`, `readQuotas`, etc.)
- Les fonctions de cache (`readElevesFromCache`, `openCacheTabs`, etc.)
- Les fonctions utilitaires (`_u_`, `_arr`, `parseCodes_`, `logLine`)
- Les fonctions de mobilité (`computeMobilityFlags_`, `buildClassOffers_`)
- Les fonctions LEGACY recopiées (`calculateClassScores_LEGACY_`, etc.)

Ce fichier est une copie monolithique de presque tout le projet. Il rend les fichiers modulaires (`App.Core.js`, `App.Context.js`, `App.CacheManager.js`) **inopérants** car il redéfinit toutes leurs fonctions.

**Correctif recommandé** : Supprimer toutes les fonctions de `Orchestration_V14I.js` qui existent déjà dans `App.*.js`. Ne garder que la logique propre à l'orchestration (pipeline, phases).

---

## ARCH-02 — HAUTE : 11 fichiers LEGACY encore présents (~8 000+ lignes)

| Fichier | Lignes |
|---------|--------|
| `LEGACY_Pipeline.js` | 663 |
| `LEGACY_Phase1_OptionsLV2.js` | ~350 |
| `LEGACY_Phase2_DissoAsso.js` | ~300 |
| `LEGACY_Phase3_Parite.js` | ~560 |
| `LEGACY_Phase4_JulesCodex.js` | ~800 |
| `LEGACY_Phase4_Optimisation.js` | ~700 |
| `LEGACY_Context.js` | ~630 |
| `LEGACY_Init_Onglets.js` | ~520 |
| `LEGACY_Mobility_Calculator.js` | ~300 |
| `LEGACY_Mobility.js` | ~200 |
| `LEGACY_Diagnostic.js` | ~400 |

**Impact** : ~5 000 lignes de code mort ou partiellement utilisé qui alourdit le namespace global GAS et augmente le temps de chargement.

---

## ARCH-03 — HAUTE : État global non contractualisé

`InterfaceV2_CoreScript.html:3785` définit un objet `CONFIG` avec des valeurs par défaut :
```javascript
adminPassword: 'admin123',
```

Cet objet est modifié de façon incontrôlée par le code frontend. Il n'y a pas de pattern Redux/Store ou même de simple getter/setter protégé.

---

## ARCH-04 — HAUTE : Double fichier de pipeline — `BASEOPTI_Architecture_V3.js` vs `Phases_BASEOPTI_V3_COMPLETE.js`

Ces deux fichiers définissent les mêmes fonctions (`findLeastPopulatedClass_V3`, `findClassWithoutCodeD_V3`). L'un semble être une version "complète" de l'autre, mais les deux sont chargés. Collision GAS garantie.

---

## ARCH-05 — MOYENNE : Aucun test automatisé

- 0 fichier de test
- 0 framework de test
- 0 CI/CD
- Tests manuels uniquement via le tableur

**Impact** : Chaque modification est un pari. Les 38+ fonctions dupliquées ne peuvent pas être refactorées sans risque tant qu'il n'y a pas de tests.

---

## ARCH-06 — MOYENNE : `getConfig()` monolithique (>250 lignes dans `Config.js`)

La fonction `getConfig()` dans `Config.js` lit toute la configuration depuis `_CONFIG`, avec des fallbacks en cascade. Elle est trop longue et difficile à tester ou modifier.

---

# 5. PERFORMANCE

## PERF-01 — HAUTE : Lecture cellule par cellule au lieu de batch

**Fichier** : `Config.js:644`
```javascript
const data = configSheet.getRange("A:A").getValues().flat();
```
C'est bien (lecture batch). Mais ailleurs :

**Fichier** : `AdminPasswordHelper.js:16`
```javascript
const password = configSheet.getRange('B3').getValue();
```
**Fichier** : `Code.js:928`
```javascript
const password = configSheet.getRange('B2').getValue();
```
**Fichier** : `App.Context.js:181`
```javascript
const value = uiSheet.getRange('B2').getValue();
```

Chaque `.getValue()` unitaire = 1 appel API Sheets (~100-200ms). Les patterns `getValues()` en batch sont nettement préférables.

---

## PERF-02 — HAUTE : Risque de timeout 6 minutes

`Orchestration_V14I.js:227` (`runOptimizationV14FullI`) orchestre les 4 phases en séquence. Avec des jeux de données importants, le temps d'exécution peut dépasser la limite GAS de 6 minutes.

Le mode streaming (`Orchestration_V14I_Stream.js`) est la bonne solution, mais son adoption n'est pas systématique.

---

## PERF-03 — MOYENNE : Données rechargées plusieurs fois

Les fonctions `readElevesFromSelectedMode_()`, `readElevesFromCache_()` et `loadAllStudentsData()` sont appelées séparément par différentes phases. Chaque appel relit les données depuis les Sheets au lieu d'utiliser un cache en mémoire pendant l'exécution.

---

## PERF-04 — MOYENNE : Pas d'utilisation du CacheService GAS

Google Apps Script offre `CacheService.getScriptCache()` pour mettre en cache des données entre exécutions (jusqu'à 6h). Cette API n'est pas utilisée — seuls des onglets CACHE sont utilisés (beaucoup plus lents).

---

## PERF-05 — MOYENNE : `computeMobilityFlags_` recalcule à chaque phase

Définie dans 3 fichiers (`Orchestration_V14I.js:2795`, `App.CacheManager.js:273`, `Mobility_System.js:131`), cette fonction est potentiellement appelée plusieurs fois dans le pipeline, recalculant les drapeaux de mobilité pour chaque élève à chaque fois.

---

## PERF-06 — BASSE : Tailwind CSS en mode CDN

```html
<script src="https://cdn.tailwindcss.com"></script>
```
Le mode CDN de Tailwind génère les styles **à la volée dans le navigateur**, ce qui est lent pour les gros fichiers HTML. En production, un build Tailwind pré-compilé serait nettement plus rapide.

---

# 6. FRONTEND & DÉPENDANCES CDN

## CDN-01 — CRITIQUE : Aucune intégrité SRI (Subresource Integrity)

**Fichier** : `InterfaceV2.html:62-94`

| Bibliothèque | URL | Version | SRI |
|--------------|-----|---------|-----|
| Tailwind CSS | `cdn.tailwindcss.com` | `@latest` ⚠️ | ❌ Aucun |
| SortableJS | `cdn.jsdelivr.net` | 1.15.0 | ❌ Aucun |
| Font Awesome | `cdnjs.cloudflare.com` | 6.4.0 | ❌ Aucun |
| Chart.js | `cdn.jsdelivr.net` | `@latest` ⚠️ | ❌ Aucun |
| SheetJS/XLSX | `cdnjs.cloudflare.com` | 0.18.5 | ❌ Aucun |
| jsPDF | `cdnjs.cloudflare.com` | 2.5.1 | ❌ Aucun |
| html2canvas | `cdnjs.cloudflare.com` | 1.4.1 | ❌ Aucun |
| Shepherd.js | `cdn.jsdelivr.net` | `@latest` ⚠️ | ❌ Aucun |

**Impact** : Sans SRI, si un CDN est compromis, du code malveillant peut être injecté dans votre application. 3 bibliothèques utilisent `@latest` (pas de version fixée), ce qui signifie qu'une mise à jour cassante peut survenir à tout moment sans avertissement.

**Aussi** dans `ConsolePilotageV3.html:9` :
- Font Awesome 6.5.1 (vs 6.4.0 dans InterfaceV2 — versions incohérentes)

**Correctif recommandé** :
1. Fixer toutes les versions (remplacer `@latest` par une version précise)
2. Ajouter les attributs `integrity="sha384-..."` et `crossorigin="anonymous"` à chaque balise `<script>` et `<link>`
3. Harmoniser la version de Font Awesome

---

## CDN-02 — HAUTE : `google.script.run` sans `withFailureHandler` systématique

**Statistiques** :
- `google.script.run` : **39 occurrences** dans 7 fichiers
- `withSuccessHandler` : **25 occurrences** dans 6 fichiers
- `withFailureHandler` : **25 occurrences** dans 7 fichiers

→ **14 appels `google.script.run` n'ont ni failureHandler ni successHandler**

**Impact** : Si le serveur renvoie une erreur, l'utilisateur ne voit rien — l'UI reste figée sans message d'erreur. L'application semble "morte".

**Correctif recommandé** : Auditer chaque `google.script.run` et s'assurer qu'un `withFailureHandler` est toujours défini.

---

## CDN-03 — HAUTE : `escapeHtml()` utilisé de façon incohérente

Dans `InterfaceV2_CoreScript.html:5132`, on voit un bon exemple :
```javascript
header.innerHTML = `...${escapeHtml(title)} - ${escapeHtml(classe)}...`
```

Mais la majorité des ~80 `innerHTML` n'utilisent PAS `escapeHtml()`. Il n'y a pas de convention claire sur quand sanitiser et quand ne pas le faire.

---

## CDN-04 — MOYENNE : Font Awesome chargé en version différente

- `InterfaceV2.html:83` → Font Awesome **6.4.0**
- `ConsolePilotageV3.html:9` → Font Awesome **6.5.1**

Risque de rendu incohérent entre les interfaces.

---

## CDN-05 — MOYENNE : Pas de fallback en cas d'indisponibilité CDN

Si un CDN est en panne (cdnjs, jsdelivr, etc.), l'application ne fonctionnera plus du tout. Aucun mécanisme de fallback (chargement local) n'est prévu.

---

## CDN-06 — BASSE : `document.write` bloqué puis utilisé

`InterfaceV2.html:12-15` bloque `document.write` pour des raisons de sécurité, mais `GroupsInterfaceV4.html:4313` et `InterfaceV2_CoreScript.html:6983` l'utilisent quand même (dans des fenêtres popup). Approche incohérente.

---

# 7. QUALITÉ DE CODE

## QC-01 — HAUTE : `logParityDecision` définie 2 fois dans le même fichier

**Fichier** : `Phases_BASEOPTI_V3_COMPLETE.js`
- Ligne 899 : première définition
- Ligne 936 : deuxième définition (écrase la première)

La première définition est **morte** — elle ne sera jamais exécutée.

---

## QC-02 — HAUTE : Fonctions mega-longues

| Fichier | Fonction | Lignes |
|---------|----------|--------|
| `Orchestration_V14I.js` | `runOptimizationV14FullI()` | ~175 |
| `Orchestration_V14I.js` | `dispatchElevesWithQuotas_()` | ~170 |
| `Config.js` | `getConfig()` | ~250+ |
| `InterfaceV2_CoreScript.html` | Multiple functions | 100+ chacune |
| `Orchestration_V14I.js` | `computeMobilityFlags_()` | ~210 |
| `Orchestration_V14I.js` | `auditCacheAgainstStructure_()` | ~210 |

**Impact** : Fonctions très difficiles à tester, comprendre et modifier. Risque élevé de régression.

---

## QC-03 — MOYENNE : Debug logging excessif en production

Des centaines d'appels `Logger.log()` et `console.log()` parsèment le code de production, incluant des données potentiellement sensibles (noms d'élèves, mots de passe).

---

## QC-04 — MOYENNE : Nommage incohérent

- Fonctions françaises : `verifierMotDePasseAdmin`, `chargerContraintes`, `creerOngletsSourcesVides`
- Fonctions anglaises : `calculateGlobalScore`, `findBestSwap`, `computeMobilityFlags`
- Mix des deux : `readQuotasFromUI_LEGACY`, `compterEffectifsOptionsEtLangues`
- Suffixes incohérents : `_V3`, `_LEGACY`, `_Phase3`, `_Ultimate`, `_V3_`, `_LEGACY_`

---

## QC-05 — MOYENNE : Magic strings/numbers

Exemples :
- `Config.js:19` : `"admin123"` (mot de passe par défaut)
- `Code.js:928` : `'B2'` (position hardcodée de la cellule config)
- `AdminPasswordHelper.js:16` : `'B3'` (position **différente** pour le même champ !)
- Noms d'onglets hardcodés : `"_CONFIG"`, `"_STRUCTURE"`, `"_BASEOPTI"` dans de nombreux fichiers

---

## QC-06 — BASSE : Commentaires obsolètes

`Code.js:928` : `// ✅ Corrigé: B2 au lieu de B3` — ce commentaire est correct, mais `AdminPasswordHelper.js:16` lit toujours B3. Le "correctif" n'a été appliqué que dans un fichier.

---

## QC-07 — BASSE : Pas de JSDoc/types

Aucune annotation de type cohérente. Quelques fichiers ont des `@param` JSDoc, mais la majorité du code n'a aucune documentation de type, rendant la compréhension et le refactoring difficiles.

---

# 8. PLAN D'ACTION PRIORISÉ

## Phase 1 — Urgences sécurité (immédiat)

| # | Action | Fichiers | Effort |
|---|--------|----------|--------|
| 1 | Supprimer tous les `"admin123"` hardcodés | Config.js, ConsolePilotageV3.html, ConsolePilotageV3_Server.js, InterfaceV2_CoreScript.html | 1h |
| 2 | Supprimer `getAdminPasswordFromConfig` du frontend — utiliser uniquement `verifierMotDePasseAdmin` | InterfaceV2_CoreScript.html, Code.js | 2h |
| 3 | Supprimer `AdminPasswordHelper.js` (doublon dangereux B3 vs B2) | AdminPasswordHelper.js | 10min |
| 4 | Supprimer les logs de mot de passe | DiagnosticConfig.js, Code.js | 30min |

## Phase 2 — Dé-duplication critique (1-2 jours)

| # | Action | Impact |
|---|--------|--------|
| 5 | Purger les fonctions recopiées dans `Orchestration_V14I.js` → importer depuis `App.*.js` | Élimine ~38 collisions |
| 6 | Supprimer `Backend_Core.js` (toutes ses fonctions existent dans `Backend_Eleves.js` et `Code.js`) | -5 doublons |
| 7 | Fusionner `BASEOPTI_Architecture_V3.js` dans `Phases_BASEOPTI_V3_COMPLETE.js` | -2 collisions |
| 8 | Supprimer la double définition de `logParityDecision` dans `Phases_BASEOPTI_V3_COMPLETE.js` | Fix bug |

## Phase 3 — Nettoyage LEGACY (2-3 jours)

| # | Action | Impact |
|---|--------|--------|
| 9 | Identifier quelles fonctions LEGACY sont encore appelées | Cartographie |
| 10 | Supprimer les fichiers LEGACY non-utilisés | -5000 lignes |
| 11 | Renommer les fonctions LEGACY encore utilisées avec un suffixe clair | Clarté |

## Phase 4 — Frontend & CDN (1 jour)

| # | Action | Impact |
|---|--------|--------|
| 12 | Fixer les versions CDN (`@latest` → version précise) | Stabilité |
| 13 | Ajouter SRI à toutes les balises CDN | Sécurité |
| 14 | Ajouter `withFailureHandler` aux 14 appels `google.script.run` manquants | UX |
| 15 | Harmoniser la version Font Awesome | Cohérence |

## Phase 5 — Qualité à long terme

| # | Action | Impact |
|---|--------|--------|
| 16 | Créer des tests unitaires pour les fonctions critiques (au minimum Phase 4 swap logic) | Fiabilité |
| 17 | Refactorer `Orchestration_V14I.js` en modules (<500 lignes chacun) | Maintenabilité |
| 18 | Standardiser le nommage (choisir FR ou EN, pas les deux) | Lisibilité |

---

*Rapport généré le 2026-02-12 par audit statique automatisé.*
