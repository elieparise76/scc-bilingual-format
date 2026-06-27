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

python main.py 20264                      # → scc_20264_bilingue.docx (EN | FR)
python main.py 20264 --lang-order fr      # français à gauche (défaut: en)
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

- **Decision** : `title`, `neutral_citation`, `sections: List[Section]`, `hearing_date`, `decision_date`, `appeal_from`, `catchwords` (mots-clés du sommaire), `held` (mention « Held / Arrêt »)
- **Section** : `type` (majority|concurring|dissent|headnotes|other), `author`, `paragraphs`, `lead_in` (mention d'attribution verbatim, ex. « English version of the judgment delivered by\nTHE COURT — »)
- **TextRun** : `text`, `italic`, `bold` — fragment de texte stylé
- **Paragraph** : `number`, `text` (brut), `contains_quote`, `contains_citation`, `headings: List[str]`, `runs: List[TextRun]` (texte découpé en fragments italique/gras)
- **ParagraphPair** : `number`, `section_type`, `fr/en: Optional[Paragraph]` (un côté peut être None)
- **AlignedSection** (sortie Aligner) : `type`, `author_fr`, `author_en`, `lead_in_fr/en`, `pairs` — auteurs et mentions d'attribution dans les deux langues, pour les en-têtes et bandeaux du Renderer

### Points d'attention

**Fetcher** — PDF à `https://decisions.scc-csc.ca/scc-csc/scc-csc/{lang}/{item_id}/1/document.do` (`{lang}` ∈ en/fr, **même item_id pour les deux langues**). Le site renvoie **403 sans User-Agent de navigateur** (en-tête obligatoire). PDFs gardés en mémoire. *À faire* : résoudre référence neutre → item_id (l'API Cludo et les pages « Case in Brief » de scc-csc.ca exposent l'ID, mais aucune route directe et fiable n'a été retenue).

**Parser** — la mise en page CSC est un gabarit stable :
- **Couverture et bloc d'opinions sur deux colonnes** : on sépare par coordonnée `x0` des mots (pas `extract_text()`, qui les entrelace). Indispensable car la citation peut être coupée par la colonne de droite (« R. c. Wolfe, 2024 » … « CSC 34 »). Seuils ≈ x0 310 (en-tête), 160 (opinions). Les **dates d'audition et de jugement** s'extraient en **recadrant** la colonne droite de la couverture (`_cover_right_text`, `re.S` car la date chevauche le retour à la ligne ; le label « JUGEMENT RENDU » sur 2 lignes serait sinon entrelacé).
- **Mention d'appel + mots-clés (tags)** (`appeal_from`, `catchwords`) : depuis le sommaire, ancrés sur « ON APPEAL FROM / EN APPEL DE ». Les **mots-clés sont en italique** dans le PDF (la prose du sommaire qui suit est en romain) → frontière nette : on prend le bloc italique (`extract_words` + `fontname`).
- **Mention « Held / Arrêt »** (`held`) : 1re phrase (dispositif) après « Held (…): » / « Arrêt (…) : » (regex `_HELD`, lookahead sur fin de phrase). La parenthèse des dissidents est conservée si présente.
- **Bloc d'opinions** (sous `CORAM`) = **source de vérité des sections** (le corps ne marque pas fiablement les frontières) : type + auteur + plage de paragraphes (« paras. 1 to 92 »). Type via mots-clés (`DISSID`→dissent, `CONCORD`→concurring, sinon majority).
- **Paragraphes** : extraction **ligne à ligne sur les mots** (`extract_words(extra_attrs=["fontname"])`, pas `extract_text()`) pour conserver les **italiques/gras** : chaque mot devient un `TextRun` selon sa police (`...Italic...`→italique, `...Bold...`→gras), fragments voisins fusionnés. Marqueurs `[N]` **séquentiels depuis [1]** (écarte les années `[2017]`). Rattachés à une section par leur plage. *(Soulignement/surlignage non exposés par pdfplumber — et absents des PDF CSC.)*
- **Sous-titres** (`II. Contexte`, `A. …`, `(1) …`) : aucune marque typographique distinctive. Une ligne au patron de plan n'est un titre que si elle précède *immédiatement* un `[N]` (sinon réabsorbée dans la prose → écarte les initiales d'auteurs « N. Metallic, … »). Stockés dans `Paragraph.headings` (niveaux empilables).
- **Mention d'attribution** (`Section.lead_in`) : capturée verbatim avant la 1re ¶ de chaque opinion (`_extract_lead_ins`), ancrée sur la ligne d'auteur finissant par « — » (« THE COURT — », « LA JUGE MOREAU — ») ; le préambule = lignes au-dessus jusqu'à la dernière finissant par « . ». Reproduite telle quelle dans le bandeau (varie selon original/traduction).
- **Queue** : dispositif (« Appeal allowed » / « Pourvoi accueilli ») et liste des procureurs retirés du dernier paragraphe (`_strip_tail_runs`, tronque les runs).
- `contains_quote` / `contains_citation` : heuristiques (formatage), non utilisées pour styliser.

Fixtures de test dans `samples/` : `20264` (unanime, « La Cour ») et `20701` (divisée, majorité + dissidence).

**Aligner** — appariement ¶N_FR ↔ ¶N_EN par numéro, regroupé en `AlignedSection`. Type/auteurs pris sur la version FR en priorité (repli EN). Un paragraphe présent dans une seule version → paire dont l'autre côté est `None`. Nouvelle section quand `(type, author_fr, author_en)` change.

**Renderer** (python-docx) — `DocMetadata` porte titres, citations, dates, mentions d'appel et mots-clés **dans les deux langues** (`title_fr/en`, `citation_fr/en`, `date_fr/en`, `appeal_fr/en`, `catchwords_fr/en`) + `lang_order`.
- **Page de garde bilingue** (`_add_front_matter`, suivie d'un saut de page) : **un élément par ligne de tableau** (`_bi_row`) pour que les deux langues démarrent au même niveau (comme les paragraphes du corps). De haut en bas : nom de cause, citation, date d'audition, date du jugement, mention d'appel, mots-clés/tags (italique), mention « Held / Arrêt », puis la **table des matières des motifs** = liste des opinions (rôle + auteur + plage `[début]-[fin]`, **même si unanime** : « Reasons of the Court / Motifs de la Cour »). La plage est alignée à droite par taquet à points de conduite (`_add_toc_row`, `_TOC_TAB`, `WD_TAB_LEADER.DOTS`).
- **Table des matières par opinion** (`_section_toc` + `_add_contents_toc` dans `_add_opinion_table`) : au début de chaque opinion, sous le bandeau, le plan de cette opinion (sous-titres + n° de paragraphe), indenté par niveau (`_heading_level`). **Seulement s'il y a des sous-titres détectés.**
- **En-tête courant à 3 zones par taquets** (pas de barres verticales) : **n° de page** à gauche (champ `PAGE`), ***nom de la décision* (italique) — juge rédacteur** au centre, **référence** à droite. Taquets : centre à `_COL_W`, **droite à `_CONTENT_W` (bord droit/marge)** → la citation est au bord droit. ⚠️ **Deux pièges de taquets** : (1) le style intégré `Header` a ses propres taquets par défaut (droite à 6,5 po) que Word **fusionne** avec les nôtres ; comme 6,5 < 7,7 po, la citation s'y accroche → on **retire les taquets des styles Header/Footer** (`_set_base_font`). (2) **Pas de retrait de paragraphe** : un retrait décale l'origine des taquets et le taquet droit déborde la zone de texte (Word l'ignore → citation pas à droite). Pour que le trait du bas (bordure basse, pleine largeur) s'aligne quand même sur le texte, les **colonnes vont bord à bord** : marges de cellule **extérieures nulles, gap uniquement au centre** (`_set_cell_margins` par cellule, `_col_margins`). Le **nom de cause est tronqué** (`_header_case`, `_HEADER_TITLE_MAX`) : un titre long déborde les taquets et empêche la citation d'atteindre la droite. Le juge est un champ **`STYLEREF`** pointant des marqueurs (styles `OpinionRefEN`/`OpinionRefFR`, **juge seul, sans rôle**) posés au début de chaque opinion **au niveau du corps** (hors tableau). L'en-tête **alterne par page** (anglais sur impaires, français sur paires ; page 1 = anglais) via `w:evenAndOddHeaders`. ⚠️ Les marqueurs sont rendus invisibles par **blanc + 1 pt, PAS `w:vanish`** : le texte masqué n'est pas mis en page, or STYLEREF dépend de la pagination → un marqueur masqué reste introuvable et l'en-tête sort vide. Un marqueur précoce avant le titre évite l'erreur STYLEREF en page 1.
- **Champs Word** : `STYLEREF` et `PAGE` portent un **résultat en cache** (`_add_field(..., cached=…)`) + le doc a `w:updateFields=true` → les valeurs s'affichent dès l'ouverture (sinon un champ sans cache paraît **vide** tant qu'il n'est pas recalculé). Sinon, `Cmd+A` puis `F9`.
- **Corps** : un tableau **par opinion** (pas de saut de page entre elles), 2 colonnes **strictement égales** (`_COL_W = _CONTENT_W // 2`, la colonne de droite atteint la marge), texte **justifié**, **interligne simple**, alignement vertical haut, italiques/gras des `runs` reproduits. **Bandeau de début d'opinion = 1re ligne du tableau (2 colonnes, non fusionnée)** portant la mention d'attribution verbatim (`AlignedSection.lead_in_fr/en`), **chaque langue de son côté** (repli sur libellé synthétique si absente). Sous-titres en lignes dédiées (gras, FR/EN) au-dessus du paragraphe.
- **Réglages globaux** : ordre des langues par défaut **anglais à gauche** (`lang_order="en"`), Times New Roman partout, marges 0,4 po, **colonnes strictement égales bord à bord** (texte 0 → `_CONTENT_W`, gap au centre seulement), bordures de tableau blanches. **Pas de numéro de page en pied** (déplacé dans l'en-tête).
