# SCC Bilingual Formatter

*Read in [English](README.md).*

Outil en ligne de commande qui transforme une décision de la **Cour suprême du
Canada** en un document Word (`.docx`) **bilingue côte à côte** : anglais à gauche,
français à droite, paragraphes alignés par numéro.

À partir du seul identifiant de la décision, l'outil télécharge les deux versions
officielles, en extrait la structure (juges, opinions, paragraphes, sous-titres) et
produit un document prêt à relire ou à imprimer.

## Fonctionnalités

- 📥 **Téléchargement automatique** des PDF anglais et français depuis `decisions.scc-csc.ca`
- 🧩 **Extraction structurée** : titre, référence neutre, dates d'audition et de
  jugement, mention « en appel de » et mots-clés du sommaire ; opinions (majorité /
  concordance / dissidence) avec leur juge rédacteur et leur plage de paragraphes ; et
  les **sous-titres** du plan (`I.`, `A.`, `(1)`…, y compris non numérotés)
- ↔️ **Alignement** des paragraphes ¶N anglais ↔ ¶N français, avec **réconciliation
  bilingue des sous-titres** — là où une langue détecte un titre que l'autre a manqué,
  on le récupère depuis la version parallèle (parité exacte sur toutes les décisions
  de test)
- ↳ **Texte en retrait** — citations en bloc, extraits législatifs et listes énumérées
  sont détectés dans le PDF source et reproduits indentés et en corps légèrement plus
  petit
- 🔗 **Liens CanLII cliquables** sur chaque référence neutre (`AAAA SCC N` / `AAAA CSC N`)
  du corps — l'URL est construite **de façon déterministe à partir de la citation**,
  sans aucun appel réseau (règle de citation CanLII), et les italiques et le gras
  environnants sont préservés
- ⚖️ **Liens Justice Canada cliquables** sur les citations de lois révisées (`R.S.C. 1985, c. C-50`
  / `L.R.C. 1985, c. C-50`) du corps, vers le site **officiel** laws-lois.justice.gc.ca —
  le chapitre cité est le slug d'URL, construit de façon déterministe sans appel réseau.
  Limité aux L.R.C. 1985 (les lois annuelles et les suppléments n'ont pas de slug
  déterministe par chapitre → laissés sans lien, aucun lien mort)
- 📄 **Document Word soigné** : page de garde bilingue (nom de cause, référence, dates,
  mention d'appel, mots-clés en italique, dispositif *Held / Arrêt*, et une table des
  matières des opinions) ; une table des matières par opinion ; corps en deux colonnes ;
  en-tête courant qui indique le juge des motifs de la page et **alterne anglais /
  français à chaque page** ; Times New Roman ; texte justifié ; italiques et gras inline
  préservés

## Installation

```bash
git clone https://github.com/elieparise76/scc-bilingual.git
cd scc-bilingual
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Utilisation

Entrer la **référence neutre** ou l'**item ID Lexum** :

```bash
python main.py "2024 CSC 5"
# → scc_20264_bilingue.docx   (anglais | français)

python main.py 20264           # équivalent — item ID issu de l'URL de la décision

# Options
python main.py "2024 CSC 5" --lang-order fr               # français à gauche (défaut : en)
python main.py "2024 CSC 5" --output chemin/sortie.docx
python main.py 20264 --pdf-en en.pdf --pdf-fr fr.pdf      # fournir les PDF localement
```

Les formes `CSC` (français) et `SCC` (anglais) de la référence sont toutes deux
acceptées et renvoient au même document. L'item ID est le nombre dans l'URL de la
décision sur `decisions.scc-csc.ca/.../item/<ID>/index.do`.

> **Astuce Word** : l'en-tête utilise des champs dynamiques (numéro de page, juge
> courant). Ils sont générés avec une valeur en cache pour s'afficher à l'ouverture ;
> s'ils s'affichent quand même vides, sélectionne tout (`Cmd/Ctrl+A`) puis appuie sur
> `F9`, ou lance un aperçu avant impression, pour les rafraîchir.

## Fonctionnement

Un pipeline en cinq étapes, chacune un module avec une fonction-livrable claire :

```
item ID Lexum ("20264")
  → fetch_pdfs(item_id) -> (pdf_en, pdf_fr)              # fetcher.py
  → parse_pdf(pdf_bytes) -> Decision                     # parser.py  (×2)
  → align(decision_fr, decision_en) -> [AlignedSection]  # aligner.py
  → render_docx(sections, metadata, output) -> .docx     # renderer.py
                                                          # main.py orchestre
```

La mise en page des décisions de la CSC est un gabarit stable, exploité finement :
séparation des colonnes de la couverture par position des mots (pas `extract_text`,
qui les entrelace) ; lecture du bloc d'opinions sous `CORAM` comme source de vérité des
sections ; détection des sous-titres par leur structure (position juste avant un
marqueur de paragraphe) plutôt que par typographie ; détection des blocs en retrait par
leur écart horizontal à la marge du corps. Voir [`CLAUDE.md`](CLAUDE.md) pour les
détails d'implémentation.

## Stack

Python 3 · [httpx](https://www.python-httpx.org/) · [pdfplumber](https://github.com/jsvine/pdfplumber) · [python-docx](https://python-docx.readthedocs.io/)

## État et suite

Le pipeline complet est fonctionnel, y compris la résolution de référence neutre.

Décisions de test dans [`samples/`](samples/) : `20264` (unanime, « La Cour »),
`20701` (divisée, majorité + dissidence) et `20546` (longue, avec sous-titres non
numérotés en anglais et une table des matières imprimée — le cas de test de la
détection des sous-titres).

## Licence

À déterminer. Les décisions de la Cour suprême du Canada sont reproductibles sans frais
(Décret sur la reproduction de la législation fédérale).
