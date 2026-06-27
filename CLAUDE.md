# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Projet

**SCC Bilingual Formatter** — outil Python en ligne de commande qui prend une décision de la Cour suprême du Canada et génère un document Word (`.docx`) bilingue : français à gauche, anglais à droite, paragraphes alignés par numéro.

Entrée actuelle = **item ID Lexum** (ex. `20264`), pas la référence neutre (résolution `2024 CSC 5` → item ID encore à faire, voir Fetcher). Projet et doc en français ; code (noms) en anglais. Dev via venv : `.venv/bin/python`.

## Stack

Python 3 · **httpx** (téléchargement PDFs Lexum) · **pdfplumber** (extraction texte) · **python-docx** (génération `.docx`).

## Commandes

```bash
pip install -r requirements.txt          # httpx, pdfplumber, python-docx

python main.py 20264                      # → scc_20264_bilingue.docx (FR | EN)
python main.py 20264 --lang-order en      # anglais à gauche (défaut: fr)
python main.py 20264 --output chemin.docx
python main.py 20264 --pdf-en en.pdf --pdf-fr fr.pdf   # fallback PDFs locaux
```

## Architecture

Pipeline en 5 étapes, chacune un module avec une fonction-livrable :

```
item ID Lexum ("20264")
  → fetch_pdfs(item_id) -> (pdf_en, pdf_fr)             # fetcher.py
  → parse_pdf(pdf_bytes) -> Decision                    # parser.py (×2)
  → align(decision_fr, decision_en) -> [AlignedSection] # aligner.py
  → render_docx(sections, metadata, output_path)        # renderer.py
                                                         # main.py orchestre
```

### Modèle de données (models.py)

- **Decision** : `title`, `neutral_citation`, `sections: List[Section]`
- **Section** : `type` (majority|concurring|dissent|headnotes|other), `author`, `paragraphs`
- **Paragraph** : `number`, `text`, `contains_quote`, `contains_citation`, `headings: List[str]`
- **ParagraphPair** : `number`, `section_type`, `fr/en: Optional[Paragraph]` (un côté peut être None)
- **AlignedSection** (sortie Aligner) : `type`, `author_fr`, `author_en`, `pairs` — auteur dans les deux langues, pour les en-têtes et bandeaux du Renderer

### Points d'attention

**Fetcher** — PDF à `https://decisions.scc-csc.ca/scc-csc/scc-csc/{lang}/{item_id}/1/document.do` (`{lang}` ∈ en/fr, **même item_id pour les deux langues**). Le site renvoie **403 sans User-Agent de navigateur** (en-tête obligatoire). PDFs gardés en mémoire. *À faire* : résoudre référence neutre → item_id (l'API Cludo et les pages « Case in Brief » de scc-csc.ca exposent l'ID, mais aucune route directe et fiable n'a été retenue).

**Parser** — la mise en page CSC est un gabarit stable :
- **Couverture et bloc d'opinions sur deux colonnes** : on sépare par coordonnée `x0` des mots (pas `extract_text()`, qui les entrelace). Indispensable car la citation peut être coupée par la colonne de droite (« R. c. Wolfe, 2024 » … « CSC 34 »). Seuils ≈ x0 310 (en-tête), 160 (opinions).
- **Bloc d'opinions** (sous `CORAM`) = **source de vérité des sections** (le corps ne marque pas fiablement les frontières) : type + auteur + plage de paragraphes (« paras. 1 to 92 »). Type via mots-clés (`DISSID`→dissent, `CONCORD`→concurring, sinon majority).
- **Paragraphes** : marqueurs `[N]` en début de ligne, **séquentiels depuis [1]** (écarte les années `[2017]` des citations). Rattachés à une section par leur plage.
- **Sous-titres** (`II. Contexte`, `A. …`, `(1) …`) : aucune marque typographique distinctive. Détection **ligne à ligne** : une ligne au patron de plan n'est un titre que si elle précède *immédiatement* un `[N]` (sinon réabsorbée dans la prose → écarte les initiales d'auteurs « N. Metallic, … »). Stockés dans `Paragraph.headings` (niveaux empilables).
- **Queue** : dispositif (« Appeal allowed » / « Pourvoi accueilli ») et liste des procureurs retirés du dernier paragraphe (`_strip_tail`).
- `contains_quote` / `contains_citation` : heuristiques (formatage). **Non utilisées pour styliser** — elles matcheraient ~83 % des paragraphes ; styliser un *bloc cité en retrait* exigerait la structure sous-paragraphe, à faire.

Fixtures de test dans `samples/` : `20264` (unanime, « La Cour ») et `20701` (divisée, majorité + dissidence).

**Aligner** — appariement ¶N_FR ↔ ¶N_EN par numéro, regroupé en `AlignedSection`. Type/auteurs pris sur la version FR en priorité (repli EN). Un paragraphe présent dans une seule version → paire dont l'autre côté est `None`. Nouvelle section quand `(type, author_fr, author_en)` change.

**Renderer** (python-docx) — `DocMetadata` porte titres + citations **dans les deux langues** (`title_fr/en`, `citation_fr/en`) + `lang_order`.
- **En-tête courant alterné par page** (anglais sur impaires, français sur paires ; page 1 = anglais) via `w:evenAndOddHeaders`. Contenu : *nom de la décision* (italique) — référence — **juge rédacteur (rôle)**. Le juge de la page est un champ **`STYLEREF`** pointant des marqueurs **masqués** (`w:vanish`, styles `OpinionRefEN`/`OpinionRefFR`, texte « Auteur (rôle) ») posés au début de chaque opinion **au niveau du corps** (hors tableau, pour que STYLEREF les retrouve). Un marqueur précoce avant le titre évite l'erreur STYLEREF en page 1. Ligne horizontale = bordure basse du paragraphe d'en-tête.
- **Rôle** : `unanime`/`unanimous` si une seule opinion, sinon `majoritaire`/`dissident`/`concordant`. Auteur = ce que nomme la Cour (« La Cour » ou le juge, tronqué avant « (avec l'accord…) »).
- **Corps** : un tableau **par opinion** (pas de saut de page entre elles), 2 colonnes, texte **justifié**, **interligne simple**, alignement vertical haut. Bandeaux de début d'opinion **sans couleur** (gras). Sous-titres en lignes dédiées (gras, FR/EN) au-dessus du paragraphe.
- **Réglages globaux** : Times New Roman partout, marges 0,4 po, bordures de tableau blanches, espace inter-colonnes via marges de cellule, pied de page = champ `PAGE`.
- ⚠️ Les champs Word (`STYLEREF`/`PAGE`) se rafraîchissent à l'ouverture ou à l'impression ; sinon `Cmd+A` puis `F9`.
