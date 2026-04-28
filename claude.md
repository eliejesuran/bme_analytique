# Analytique ↔ Budget — Documentation projet

## Contexte

Outil développé pour le département finance d'une **ASBL communale bruxelloise d'événementiel** (~70 événements/an en espace public). Il permet aux coordinateurs de projet de faire le suivi budgétaire sans avoir à mettre à jour manuellement, ligne par ligne, leur colonne "Analytique" depuis une extraction Winbooks.

---

## Fichiers

| Fichier | Rôle |
|---|---|
| `index.html` | L'outil complet (HTML + CSS + JS, standalone) |
| `logo_white.png` | Logo chargé dans le header (591×407 px, fond transparent) |
| `EVENT_20261.xlsx` | Exemple de fichier budget (template de référence) |
| `Historique_pe_riodique_croise__par_compte_du_01_12_2025_a__13_55_10.xlsx` | Exemple d'extraction Winbooks |
| `claude.md` | Documentation technique du projet (ce fichier) |

Les fichiers `index.html` et `logo_white.png` doivent être dans le **même dossier** pour que le logo s'affiche.

---
## Métadonnées & authorship

Balises meta présentes dans `index.html` :

| Meta | Valeur |
|---|---|
| `author` | Elie JESURAN |
| `copyright` | © 2026 Elie JESURAN |
| `date` | 2026-04-28 |

---

## Architecture technique

### Stack
- **HTML/CSS/JS pur** — aucun framework, aucun serveur requis
- **xlsx.js 0.18.5** (CDN Cloudflare) — parsing des fichiers Excel côté client
- **Pyodide 0.25.0** (CDN jsdelivr) — Python dans le navigateur pour l'export Excel
- **openpyxl** (via micropip dans Pyodide) — génération de l'Excel de sortie avec préservation des styles
- **DM Sans + DM Mono** (Google Fonts) — typographie

### Bibliothèques CDN
```
https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
https://cdn.jsdelivr.net/pyodide/v0.25.0/full/pyodide.js
https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600
```

### Persistance
- **localStorage** (`abv4`) — session sauvegardée automatiquement à chaque action
- Le logo n'est **pas** persisté en localStorage (il est chargé via `src="logo_white.png"`)
- L'historique undo est **en mémoire uniquement** (perdu au rechargement)

---

## Structure du fichier budget Excel

Basé sur le template `EVENT_XXXXX.xlsx` :

| Cellule | Contenu |
|---|---|
| A3 | Nom de l'événement (affiché dans le header de l'outil) |
| A5 | Exercice / année |
| **Col I (0-based: 8)** | ANALYTIQUE — colonne mise à jour par l'outil |
| **Col M (0-based: 12)** | Commentaire — dernière colonne d'origine |
| **Col N (0-based: 13)** | N° Factures — **nouvelle colonne ajoutée par l'outil** |

### Détection automatique de la colonne ANALYTIQUE
L'outil cherche les cellules contenant `ANALYTIQUE` ou `ANALYTIQUE YYYY` (ex: `ANALYTIQUE 2025`) dans les 25 premières lignes. Il retient la colonne qui apparaît **au moins 2 fois** (une fois dans le bloc PRODUITS, une dans CHARGES) avec **l'année la plus grande**. En cas d'égalité, la colonne la plus à droite est choisie.
Regex : /^ANALYTIQUE(?:\s*(20)?(\d{2}))?$/
Reconnaît : "ANALYTIQUE", "ANALYTIQUE 2025", "ANALYTIQUE 25" (avec ou sans le préfixe "20")

### Types de lignes (détection par couleur de fond col A)
| Couleur RGB | Classe CSS | Signification |
|---|---|---|
| `FFD4EAF3` | `rh` | En-tête de section (bleu clair) |
| `FFCEDBE6` | `rs` | Sous-section (bleu moyen) |
| bold + fond coloré | `rt` | Ligne de total |
| bold + pas de fond | `rsum` | Résumé/total général |
| pas bold + pas de fond | `rd` | **Ligne data — cliquable** |

### Formules SUBTOTAL
Les cellules `=SUBTOTAL(9,I19:I26)` sont recalculées en temps réel dans l'outil. L'algorithme :
1. Parse la plage `I19:I26`
2. Pour chaque cellule : utilise `cell.v` si numérique non-nul, sinon évalue les formules arithmétiques simples (`=245.45+400.00`)
3. Ajoute les deltas de session (`S.rowData[rowNum].parts`)
4. Gère les SUBTOTAL imbriqués récursivement sans double-comptage

### Colonnes visibles vs cachées (dans le template EVENT)
Certaines colonnes sont masquées dans Excel (hidden:true) :
B, D, F, H, I, J, K sont cachées dans le template de référence.
L'outil les affiche toutes mais elles peuvent être masquées via
le mécanisme de toggle de colonnes (clic sur en-tête).

---

## Flux de l'extraction Winbooks

### Colonnes lues
`Libellé`, `Nom`, `Période`, `Journal`, `N°Doc`, `Date`, `Solde`, `Commentaire`

La détection des colonnes est insensible aux accents (normalize NFD) et à la casse.

### Ligne ignorées automatiquement
- Lignes sans `N°Doc` (lignes de total en fin d'extraction)
- Lignes entièrement vides

### Clé d'unicité (déduplication inter-sessions)
```
key = N°Doc + "§" + round(|solde| * 100) + "§" + commentaire[0:40]
```

### Détection des lignes déjà traitées (reprise de session)
Au chargement du budget, l'outil lit la colonne N° Factures. Tout `N°Doc` trouvé dans cette colonne est marqué comme déjà traité → les lignes d'extraction correspondantes sont automatiquement sautées.

### Note sur le fichier d'extraction
N'importe quel nom de fichier .xlsx est accepté (plus de dépendance
au nom "Historique périodique..."). La détection des colonnes se fait
uniquement par le contenu des en-têtes (insensible accents + casse).

---

## Session & actions utilisateur

### État persisté (`localStorage['abv4']`)
```js
{
  processed: {
    [extKey]: { skipped: bool, assignments: [{rowNum, amount, ndoc}] }
  },
  rowData: {
    [rowNum]: { parts: [{signed, abs, isCr}], ndocs: [string] }
  },
  curIdx: number,
  _alreadyProcessedNdocs: [string]
}
```

### Actions disponibles
| Action | Effet |
|---|---|
| **Clic sur ligne budgétaire** | Affecte le montant courant à cette ligne |
| **⏩ Passer** | Marque la ligne d'extraction comme skippée |
| **✂ Split** | Mode partiel : saisir un montant < total, cliquer sur plusieurs lignes |
| **↩ Annuler** | Annule la dernière action (assign ou skip) — en mémoire seulement |
| **Clic sur en-tête colonne** | Masque/affiche la colonne (protégé : col A, Analytique, N° Factures) |
| **↺ Réinitialiser** | Efface localStorage et recharge |

---

## Export Excel (via Pyodide + openpyxl)

Le bouton "Exporter Excel" déclenche :
1. Chargement **paresseux** de Pyodide (~15s, une seule fois par session navigateur)
2. Installation d'openpyxl via micropip
3. Encodage base64 chunked du fichier budget (chunks de 8192 octets, évite le stack overflow)
4. Exécution du script Python `PY_EXPORT` dans Pyodide

### Ce que le script Python fait
- Charge le workbook original avec openpyxl (préserve TOUS les styles, couleurs, bordures)
- Ajoute l'en-tête "N° Factures" en col N avec le style copié de col M
- Pour chaque ligne affectée :
  - **Analytique** : écrit `=(ancien_contenu)+245.45+400.00*-1+...` (formule additive)
  - **Factures** : écrit `"25002395 | 25002400 | ..."` (pipe-séparé, dédupliqué)
- Copie le style de la cellule M correspondante vers la cellule N

### Préservation des styles
openpyxl préserve les couleurs de fond, polices, bordures et formats numériques de toutes les cellules non modifiées. Les cellules modifiées (col I, col N) reçoivent le style copié de leur voisine (col M pour N, style existant pour I).

---

## Résolution des couleurs (affichage HTML)

xlsx.js lit les styles Excel en format XF (indices). Les couleurs peuvent être :
- **RGB direct** (`fgColor.rgb`) → utilisé directement (strip alpha FF)
- **Thème + tint** (`fgColor.theme` + `fgColor.tint`) → résolu via table de thème + formule de tint Excel
- **Indexé** → ignoré (couleurs legacy)

### Table de thème par défaut (Office standard)
```
0:dk1=#000000  1:lt1=#FFFFFF  2:dk2=#44546A  3:lt2=#E7E6E6
4:acc1=#4472C4  5:acc2=#ED7D31  6:acc3=#A5A5A5  7:acc4=#FFC000
8:acc5=#5B9BD5  9:acc6=#70AD47
```

### Formule de tint
- Tint ≥ 0 : `canal = canal + (255 - canal) × tint` (éclaircit vers blanc)
- Tint < 0 : `canal = canal × (1 + tint)` (assombrit vers noir)

### Filtre anti-artefact
Les couleurs résolues avec une luminance < 0.08 (quasi-noires) sont ignorées pour l'affichage — ce sont des artefacts de lecture de cellules blanches avec `theme:0` (dk1).

### Fonction `loadThemeColors(wb)`
Appelée juste après `XLSX.read()`. Tente de lire `wb.Themes` (tableau
exposé par xlsx.js en mode cellStyles:true). Si absent ou mal formé,
reste sur la table Office standard `OFFICE_THEME`.

### Ordre de priorité dans `resolveFill()`
1. Si `patternType` absent ou `'none'` → pas de fond
2. Si `fgColor.rgb` présent et non-transparent → utiliser directement
3. Si `fgColor.theme` présent → résoudre via themeColors + applyTint()
4. Si luminance résolue < 0.08 → ignorer (artefact dk1)

---

## Points connus / limitations

- L'historique undo ne survit pas au rechargement de page
- Le split ne peut être annulé qu'en bloc (pas partie par partie)
- Les formules SUBTOTAL supportées : `=SUBTOTAL(9, plage)` uniquement (pas les variantes avec plusieurs plages)
- Pyodide nécessite une connexion internet au premier export de la session
- Le logo doit être dans le même dossier que `index.html` (pas de fallback localStorage)

---

## Versioning informel

| Version | Changements clés |
|---|---|
| v1 | Structure de base : upload, affectation ligne par ligne, export xlsx.js |
| v2 | Pyodide + openpyxl pour export fidèle, déduplication inter-sessions |
| v3 | Couleurs de fond Excel, SUBTOTAL live, split avec avertissement |
| v4 | Bouton undo, masquage de colonnes par clic header, nom événement dans header |
| v5 | Détection ANALYTIQUE YYYY, fix double evalArith, description extraction générique |
| v6 | Résolution couleurs thème (theme+tint), nettoyage code, documentation |
| v6.1 | Regex ANALYTIQUE étendue (format YY et YYYY), balises meta author/copyright/date |
| v6.1 | Regex ANALYTIQUE étendue `/^ANALYTIQUE(?:\s*(20)?(\d{2}))?$/`,
         balises meta author/copyright/date (Elie JESURAN, 2026-04-28) |
| v6.2 | (en cours) Couleurs de fond via résolution thème+tint,
         nettoyage double evalArith, description extraction générique |


---
## État au 28 avril 2026 — fin de session

### Travail en cours (non terminé)
- **Couleurs de fond des cellules** : chantier initié mais pas finalisé.
  La fonction `resolveFill()` est en place et gère RGB + thème+tint.
  Problème identifié : les lignes "data" (blanches dans Excel) ont parfois
  `theme:0 tint:0` (dk1=noir) comme fgColor → filtré via luminance < 0.08.
  À tester sur un vrai fichier avec des couleurs variées pour valider.

### Bugs connus
- Le split-undo est conservateur : il supprime toutes les parts du rowData
  mais ne peut pas identifier précisément quelles parts appartiennent à quel
  split. L'utilisateur doit réaffecter depuis zéro après un undo de split.
- `stillUsed` dans undoLast (type:'assign') est calculé mais jamais utilisé
  pour la décision — la décision réelle repose sur `otherAssignmentsUseNdoc`.
  Variable à nettoyer.
- Les SUBTOTAL imbriqués (ex: I27 = SUBTOTAL de I18:I26 qui contient lui-même
  des SUBTOTAL) sont gérés récursivement mais non testés en profondeur.

### Décisions de design importantes (à ne pas oublier)
- L'export passe par **Pyodide + openpyxl** (pas xlsx.js) pour préserver
  fidèlement les styles Excel. Pyodide se charge paresseusement au premier
  clic "Exporter".
- La clé localStorage est `abv4`. Si tu changes la structure de S{},
  incrémente la clé pour éviter des conflits avec des sessions anciennes.
- La colonne N° Factures utilise `|` comme séparateur (avec espaces autour).
  Ce séparateur est parsé dans `getExistingFactures()` et dans le script
  Python `PY_EXPORT`. Ne pas changer sans mettre à jour les deux endroits.
- Le script Python `PY_EXPORT` est une string template dans le JS.
  Les `\\\\` dans `number_format` sont intentionnels (double-échappement
  JS→Python→openpyxl).

### Ce qui a été testé et fonctionne
- Chargement budget + extraction, affectation, skip, split, undo
- Export Pyodide avec formule additive dans col I
- Reprise de session : détection des N°Doc déjà dans col N
- SUBTOTAL recalculés en temps réel après affectation
- evalArithFormula pour lire les formules =245.45+400.00 écrites par l'export
- Masquage/affichage de colonnes par clic sur l'en-tête
- Nom de l'événement lu depuis A3 affiché dans le header


---
## Prochaine session

### À faire en priorité
1. **Valider les couleurs de fond** sur un fichier budget réel avec des
   couleurs variées — vérifier que `resolveFill()` donne le bon résultat
   visuel pour les lignes rh (bleu clair), rs (bleu moyen), rt (totaux),
   et que les data rows restent blanches.
2. **Nettoyer `stillUsed`** dans `undoLast` (type:'assign') — variable
   calculée mais inutilisée.
3. **Tester ANALYTIQUE 2025 / ANALYTIQUE 25** — la nouvelle regex
   `/^ANALYTIQUE(?:\s*(20)?(\d{2}))?$/` n'a pas encore été testée sur
   un vrai fichier avec ce format.
4. **Faire chercher la colonne N° Facture**  soit en créer une, soit la réutiliser, ce n'est pas fixe

### Idées / demandes en attente
- Visualisation finale léchée de l'output (mentionné en début de projet,
  reporté "à terme")
- Sous-totaux pour d'autres colonnes que ANALYTIQUE (REEL, MAJ...)
- Support multi-feuilles dans le budget

