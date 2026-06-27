# SCC Bilingual Formatter

Outil en ligne de commande qui transforme une décision de la **Cour suprême du Canada** en un document Word (`.docx`) **bilingue côte à côte** : français à gauche, anglais à droite, paragraphes alignés par numéro.

À partir du seul identifiant de la décision, l'outil télécharge les deux versions officielles, en extrait la structure (juges, opinions, paragraphes, sous-titres) et produit un document prêt à relire ou à imprimer.

## Fonctionnalités

- 📥 **Téléchargement automatique** des PDF anglais et français depuis `decisions.scc-csc.ca`
- 🧩 **Extraction structurée** : titre, référence neutre, opinions (majorité / concordance / dissidence) avec leur juge rédacteur et leur plage de paragraphes, et les **sous-titres** du plan (`I.`, `A.`, `(1)`…)
- ↔️ **Alignement** des paragraphes ¶N français ↔ ¶N anglais
- 📄 **Document Word soigné** : tableau deux colonnes, en-tête courant qui indique le juge des motifs de la page (et alterne anglais / français à chaque page), table des matières des opinions, Times New Roman, texte justifié

## Installation

```bash
git clone https://github.com/<ton-utilisateur>/scc-bilingual.git
cd scc-bilingual
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Utilisation

L'entrée est l'**item ID Lexum** de la décision — le nombre dans l'URL de la décision sur
`decisions.scc-csc.ca/.../item/<ID>/index.do` (ex. `20264` pour *2024 CSC 5*).

```bash
python main.py 20264
# → scc_20264_bilingue.docx   (français | anglais)

# Options
python main.py 20264 --lang-order en              # anglais à gauche (défaut : fr)
python main.py 20264 --output chemin/sortie.docx
python main.py 20264 --pdf-en en.pdf --pdf-fr fr.pdf   # fournir les PDF localement
```

> **Astuce Word** : l'en-tête utilise des champs dynamiques (numéro de page, juge courant). S'ils s'affichent vides à l'ouverture, sélectionne tout (`Cmd/Ctrl+A`) puis appuie sur `F9`, ou lance un aperçu avant impression, pour les rafraîchir.

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

La mise en page des décisions de la CSC est un gabarit stable, exploité finement : séparation
des colonnes de la couverture par position des mots, lecture du bloc d'opinions sous `CORAM`
comme source de vérité des sections, détection des sous-titres par leur position juste avant un
marqueur de paragraphe. Voir [`CLAUDE.md`](CLAUDE.md) pour les détails d'implémentation.

## Stack

Python 3 · [httpx](https://www.python-httpx.org/) · [pdfplumber](https://github.com/jsvine/pdfplumber) · [python-docx](https://python-docx.readthedocs.io/)

## État et suite

Le pipeline complet est fonctionnel. À venir :

- Résolution **référence neutre → item ID** (entrer `2024 CSC 5` au lieu de `20264`)
- Mise en forme distincte des **blocs cités en retrait** (nécessite la structure sous-paragraphe)

Décisions de test : `20264` (unanime) et `20701` (divisée), dans [`samples/`](samples/).

## Licence

À déterminer. Les décisions de la Cour suprême du Canada sont reproductibles sans frais
(Décret sur la reproduction de la législation fédérale).
