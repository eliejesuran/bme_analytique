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
