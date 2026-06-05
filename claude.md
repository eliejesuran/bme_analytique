# Analytique ↔ Budget — Documentation projet

## Contexte
Outil standalone (HTML/CSS/JS pur) pour coordinateurs d'une ASBL bruxelloise d'événementiel (~70 événements/an). Met à jour la colonne ANALYTIQUE d'un budget Excel depuis une extraction Winbooks, sans serveur.

**Deux inputs :**
- **Budget** (`EVENT_XXXXX.xlsx`) — structure variable selon le collaborateur ; colonnes ANALYTIQUE et COMMENTAIRE détectées dynamiquement
- **Extraction Winbooks** — structure fixe, tout fichier `.xlsx` accepté

## Fichiers
| Fichier | Rôle |
|---|---|
| `index.html` | Standalone complet |
| `logo_white.png` | Logo header |
| `EVENT_20261.xlsx` | Template budget de référence |

## Métadonnées
`author: Elie JESURAN` · `copyright: © 2026 Elie JESURAN` · `date: 2026-05-22`

---

## Stack
- **xlsx.js 0.18.5** (CDN Cloudflare) — parsing Excel côté client
- **Pyodide 0.25.0 + openpyxl** (CDN jsdelivr + micropip) — export Excel avec styles
- **DM Sans + DM Mono** (Google Fonts)
- **localStorage `abv4`** — session persistée ; incrémenter si structure `S{}` change

---

## Budget Excel

| Cellule/Col | Contenu |
|---|---|
| A3 | Nom événement (header) |
| ANALYTIQUE (dynamique, défaut col I=8) | Détectée par regex, ≥2 occurrences |
| COMMENTAIRE (dynamique, défaut col M=12) | Détectée par label exact |
| COMMENTAIRE+1 | N° Factures — **colonne ajoutée par l'outil** (`factCol = commentCol + 1`) |

### Détection colonne ANALYTIQUE
Regex : `/^ANALYTIQUE(?:\s*(20)?(\d{2}))?$/` — 25 premières lignes, colonne avec ≥2 occurrences, année la plus grande (égalité → colonne la plus à droite).

Calcul année : `m[2] ? parseInt((m[1]||'20')+m[2]) : 0` — donne 0 pour `ANALYTIQUE` nu, 2026 pour `ANALYTIQUE 26` ou `ANALYTIQUE 2026`.

### Types de lignes (couleur fond col A)
| Fond RGB | Classe | Type |
|---|---|---|
| `FFD4EAF3` | `rh` | En-tête section |
| `FFCEDBE6` | `rs` | Sous-section |
| bold + fond `FF*` non-noir | `rt` | Total |
| bold + pas de fond | `rsum` | Résumé |
| pas bold | `rd` | Data (cliquable) |

**Limite** : détection via `.fgColor.rgb` uniquement — les remplissages thème+tint ne sont pas lus ici (contrairement à l'affichage HTML qui utilise `resolveFill`). Un budget avec couleurs thème pures peut avoir des en-têtes classées `rd` (cliquables à tort).

### Sections budget (signe des montants)
- **PRODUITS / RECETTES** : signe inversé (`effectiveIsCr = !isCr`) — recettes en négatif dans le budget
- **CHARGES / DÉPENSES** : signe direct
- Section `null` (avant tout header) : traité comme CHARGES
- Détection : scan des 4 premières colonnes de chaque ligne, correspondance exacte insensible à la casse

### SUBTOTAL
`=SUBTOTAL(9, plage)` recalculé en live, y compris imbriqués (récursivement). Seule la variante à une plage `A1:A10` est supportée.

---

## Extraction Winbooks (structure fixe)
Colonnes lues : `Libellé`, `Nom`, `Période`, `Journal`, `N°Doc`, `Date`, `Solde`, `Commentaire`.
Détection par nom, insensible accents + casse. Lignes ignorées : N°Doc vide (totaux) ou entièrement vides.

**Clé dédup** : `N°Doc § round(|solde|×100) § commentaire[:40]`

---

## État persisté (`localStorage['abv4']`)
```js
{
  processed: { [extKey]: { skipped, assignments:[{rowNum,amount,ndoc}], fromFile? } },
  rowData:   { [rowNum]:  { parts:[{signed,abs,isCr}], ndocs:[string] } },
  curIdx: number,
  _alreadyProcessedNdocs: [string]
}
```

`isCr` dans `parts` est déjà corrigé pour la section (inversion PRODUITS appliquée avant stockage).

---

## Reprise de session
Au chargement du budget, `detectAlreadyProcessed(bytes)` lit la col N° Factures, extrait les N°Doc pipe-séparés → `S._alreadyProcessedNdocs`. Dans `startSession()`, ces N°Doc marquent les lignes d'extraction comme `fromFile:true` (non affichées dans le tableau, juste sautées par `advance()`). `S.rowData` n'est **pas** reconstruit depuis le fichier.

---

## Export (Pyodide + openpyxl)
1. Chargement paresseux (~15s, une fois par session) ; retry possible si erreur
2. `uint8ToBase64()` chunké 8192 bytes (évite stack overflow)
3. `PY_EXPORT` : préserve styles, écrit `=base+amt1+amt2*-1` (`isCr=true` → `*-1`), N°Doc pipe-séparés en col N
4. Autres feuilles du workbook préservées (openpyxl natif)
5. `\\\\` dans `number_format` : double-échappement JS→Python→openpyxl intentionnel
6. Séparateur `|` (avec espaces) : synchronisé entre `getExistingFactures()` et `PY_EXPORT` — ne pas modifier l'un sans l'autre

---

## Résolution couleurs (affichage HTML)
- `fgColor.rgb` → strip alpha `FF`
- `fgColor.theme + tint` → `themeColors[theme]` + `applyTint()`
- Luminance < 0.08 → ignoré (artefact dk1 sur cellules blanches)
- `loadThemeColors(wb)` : lit `wb.Themes` (xlsx.js, `cellStyles:true`) ; fallback OFFICE_THEME standard

---

## Bugs connus

| ID | Bug | Localisation | Impact |
|---|---|---|---|
| **B3** | Split-undo no-op — filtre `p=>true` ne supprime aucune part | `undoLast()` `split-complete` ~l.1145 | Undo split ne restaure rien |
| **B4** | Pas de garde sur split partiel en cours lors de Annuler / Passer | `toggleSplit()`, `skipLine()` | Peut quitter un split sans finaliser |
| **N3** | Row type detection ignore `theme+tint` fills | `parseBudget()` l.545-547 | Budgets avec fonds thème : en-têtes `rd` (cliquables à tort) |
| **N4** | `split-complete` : `markProc` enregistre seulement le dernier `rowNum` | `assign()` split branch | `S.processed[key].assignments` incomplet pour splits multi-lignes |
| **N5** | Cellules formule avec `v=0` légitimes re-évaluées par `evalArithFormula` | `parseBudget()` l.575 | Un vrai zéro peut être substitué par un résultat de formule |

---

## Décisions de design
- **localStorage key `abv4`** — incrémenter si structure `S{}` change
- **factCol = commentCol + 1** — dérivé dynamiquement, pas fixe
- **Export via Pyodide** — pour préserver fidèlement les styles Excel (xlsx.js ne peut pas)
- **Undo en mémoire seulement** — perdu au rechargement
- **`navExt` avec `renderExt(true)`** — la navigation ‹/› n'appelle pas `advance()` pour permettre de revoir des lignes déjà traitées

---

## Backlog

| Priorité | Tâche |
|---|---|
| 🔴 B3 | Fix split-undo : stocker `{rowNum, partIdx}` par affectation dans history |
| 🔴 B4 | Bloquer/avertir Annuler et Passer sur split partiel en cours (`splitUsed > 0`) |
| 🟡 N3 | Row type : étendre détection aux `theme+tint` fills (via `resolveFill`) |
| 🟡 N4 | `split-complete` : stocker tous les `{rowNum, parts}` dans history |
| 🟡 N5 | `evalArithFormula` : ne substituer que si `v` est absent ou réellement 0 intentionnel |
| 🟢 | Raccourcis clavier : Entrée = affecter, Espace = passer, Échap = fermer split |
| 🟢 | Curseurs visuels extraction : rouge = passé, vert = affecté, clic pour revenir |
| 🟢 | Filtrer extraction par période / journal |
| 🟢 | Possibilité d'ajouter des lignes au budget |
| 🟢 | Visualisation finale de l'output |
| 🟢 | Sous-totaux pour autres colonnes (REEL, MAJ…) |
| 🟢 | Support multi-feuilles budget |

---

## Historique versions
| Version | Changements |
|---|---|
| v1 | Structure de base : upload, affectation, export xlsx.js |
| v2 | Pyodide + openpyxl, déduplication inter-sessions |
| v3 | Couleurs fond Excel, SUBTOTAL live, split |
| v4 | Undo, masquage colonnes, nom événement header |
| v5 | Détection ANALYTIQUE YYYY, fix double evalArith |
| v6–v6.2 | Couleurs thème+tint, regex YY/YYYY, meta author |
| v7 | Reprise de session via col N° Factures, navigation ‹/›, mode révision |
| v8 | Affichage factures dans tableau, encodage chunked base64, sections produits/charges (B2) |
| v8.1 | Fix year detection, dead code, resumeSession index, navExt save+advance |
