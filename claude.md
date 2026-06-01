# Analytique ↔ Budget — Documentation projet

## Contexte
Outil standalone (HTML/CSS/JS pur) pour coordinateurs de projet d'une ASBL communale bruxelloise d'événementiel (~70 événements/an). Permet la mise à jour de la colonne "Analytique" d'un budget Excel depuis une extraction Winbooks, sans serveur.

## Fichiers
| Fichier | Rôle |
|---|---|
| `index.html` | Outil complet (standalone) |
| `logo_white.png` | Logo header (même dossier que index.html) |
| `EVENT_20261.xlsx` | Template budget de référence |

## Métadonnées
`author: Elie JESURAN` · `copyright: © 2026 Elie JESURAN` · `date: 2026-04-28`

---

## Stack technique
- **xlsx.js 0.18.5** (CDN Cloudflare) — parsing Excel côté client
- **Pyodide 0.25.0** + **openpyxl** (CDN jsdelivr + micropip) — export Excel avec styles
- **DM Sans + DM Mono** (Google Fonts)
- **localStorage `abv4`** — session persistée ; incrémenter la clé si structure de `S{}` change

---

## Structure budget Excel
| Cellule/Col | Contenu |
|---|---|
| A3 | Nom événement (affiché dans le header) |
| Col I (0-based: 8) | ANALYTIQUE — détectée automatiquement |
| Col M (0-based: 12) | Commentaire — dernière colonne d'origine |
| Col N (0-based: 13) | N° Factures — **colonne ajoutée par l'outil** (pas fixe : `factCol = commentCol + 1`) |

### Détection colonne ANALYTIQUE
Regex : `/^ANALYTIQUE(?:\s*(20)?(\d{2}))?$/` — cherche dans les 25 premières lignes, retient la colonne apparaissant ≥ 2 fois avec l'année la plus grande (égalité → colonne la plus à droite).

**BUG CONNU** : `m[1]` capture `"20"` ou `undefined`, pas l'année entière. `parseInt(m[1])` donne `20` ou `0` — la comparaison d'années est cassée. Fix : `parseInt((m[1]||'20') + m[2])`.

### Types de lignes (détection couleur fond col A)
| Couleur RGB | CSS class | Type |
|---|---|---|
| `FFD4EAF3` | `rh` | En-tête section |
| `FFCEDBE6` | `rs` | Sous-section |
| bold + fond coloré | `rt` | Total |
| bold + pas de fond | `rsum` | Résumé |
| pas bold | `rd` | Data — cliquable |

### Formules SUBTOTAL
`=SUBTOTAL(9,I19:I26)` recalculées en live. Les SUBTOTAL imbriqués sont gérés récursivement (non testés en profondeur). Seule la variante à une plage est supportée.

---

## Flux extraction Winbooks
Colonnes lues : `Libellé`, `Nom`, `Période`, `Journal`, `N°Doc`, `Date`, `Solde`, `Commentaire`. Détection insensible accents + casse. N'importe quel nom de fichier .xlsx accepté.

Lignes ignorées : sans `N°Doc` (totaux) ou entièrement vides.

**Clé déduplication** : `N°Doc + "§" + round(|solde|×100) + "§" + commentaire[0:40]`

**Reprise de session** : au chargement du budget, `detectAlreadyProcessed(bytes)` lit la col N° Factures, extrait les N°Doc pipe-séparés → stockés dans `S._alreadyProcessedNdocs`. Dans `startSession()`, ces N°Doc marquent les lignes d'extraction comme déjà traitées (`fromFile:true`).

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

---

## Export Excel (Pyodide + openpyxl)
1. Chargement paresseux Pyodide (~15s, une seule fois par session)
2. `uint8ToBase64()` chunké par 8192 bytes (évite stack overflow sur gros fichiers)
3. Script Python `PY_EXPORT` dans Pyodide : préserve styles, écrit formules additives `=base+amt1+amt2*-1`, écrit N°Doc pipe-séparés en col N
4. Les autres feuilles du workbook sont **préservées** (openpyxl ne les supprime pas)
5. `\\\\` dans `number_format` : intentionnel (double-échappement JS→Python→openpyxl)
6. Séparateur `|` (avec espaces) parsé dans `getExistingFactures()` ET dans `PY_EXPORT` — ne pas changer sans mettre à jour les deux

---

## Résolution couleurs (affichage HTML)
- `fgColor.rgb` → strip alpha FF, utiliser directement
- `fgColor.theme + tint` → `themeColors[theme]` + `applyTint()`
- Luminance < 0.08 → ignoré (artefact dk1 sur cellules blanches)
- `loadThemeColors(wb)` : lit `wb.Themes` (xlsx.js, cellStyles:true) ; fallback sur OFFICE_THEME standard

---

## Bugs connus

| Bug | Localisation | Impact |
|---|---|---|
| **Year detection cassé** | `parseBudget()` l.501 — `m[1]` = `"20"` pas l'année | Détection ANALYTIQUE YYYY/YY compare mal les années |
| **Split-undo no-op** | `undoLast()` type `split-complete` l.1129 — filtre `p=>true` ne supprime rien | Undo split ne restaure aucun état |
| **`stillUsed` mort** | `undoLast()` type `assign` l.1097 — `some(()=>true)` toujours vrai, var inutilisée | Dead code, à supprimer |
| **`ndocWasNew` mort** | `assign()` l.986 — calculé après push, jamais utilisé | Dead code, à supprimer |
| **`resumeSession` mauvais index** | l.1419 — toujours `extRows.length - 1`, pas la dernière non traitée | Revient sur dernière ligne et non dernière en attente |
| **Doc `detectAlreadyProcessed` périmée** | claude.md (ancienne version) — signature `(ws,analCol,factCol)` incorrecte | La vraie signature est `(bytes)` ; ne reconstruit pas `S.rowData` |

---

## Décisions de design critiques
- **localStorage key = `abv4`** — incrémenter si structure `S{}` change
- **factCol = commentCol + 1** — pas fixe, dérivé de la détection de "COMMENTAIRE"
- **Export via Pyodide** (pas xlsx.js) pour préserver fidèlement les styles
- **Undo en mémoire seulement** — perdu au rechargement

---

## Backlog / Prochaine session
| Priorité | Tâche |
|---|---|
| B1 🔴 | Fix year detection bug (`parseInt((m[1]||'20') + m[2])`) |
| B2 🔴 | Si affectation dans les PRODUITS (premiere partie du tableau budgétaire), ajouter l'opposé du montant (si -5000 --> ajouter 5000 par exemple) |
| B3 🔴 | Fix split-undo (stocker les rowNum+parts affectés dans history) |
| B4 🔴 | Bouton Annuler : bloquer ou avertir sur split partiel en cours |
| B5 🟡 | Supprimer dead code : `stillUsed`, `ndocWasNew` |
| B6 🟡 | `resumeSession` : pointer sur dernière ligne non traitée |
| B7 🟡 | Curseurs visuels : rouge = passé, vert = affecté, cliquable pour revenir |
| B8 🟢 | Possibilité d'ajouter des lignes au budget |
| B9 🟢 | Visualisation finale de l'output |
| B10 🟢 | Sous-totaux pour autres colonnes (REEL, MAJ…) |
| B11 🟢 | Support multi-feuilles budget |

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
| v8 | Affichage factures dans tableau, encodage chunked base64 |
