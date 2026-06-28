"""Liens hypertexte CanLII sur les références neutres de la Cour suprême.

Détecte les citations neutres « AAAA SCC N » (anglais) et « AAAA CSC N »
(français) dans le corps du document et transforme chaque occurrence en
hyperlien cliquable vers CanLII.

URL **déterministe**, sans aucun appel réseau : la règle de citation CanLII pose
que « when a decision has a neutral citation assigned by the issuing court, the
CanLII citation is entirely based upon the neutral citation ». Le docID = la
citation, espaces retirés, en minuscules. D'où le patron :

  « 2024 SCC 5 » → https://www.canlii.org/en/ca/scc/doc/2024/2024scc5/2024scc5.html
  « 2024 CSC 5 » → https://www.canlii.org/fr/ca/csc/doc/2024/2024csc5/2024csc5.html

⚠️ CanLII renvoie 403 à tout accès automatisé (anti-bot) : l'URL ne peut donc pas
être validée par requête HTTP. C'est attendu — elle est garantie par la règle de
citation ci-dessus et reste valide au clic humain dans Word.

Point d'injection : `renderer._emit_runs`, seul endroit qui écrit les runs du
corps (prose ET blocs en retrait, donc les citations à l'intérieur d'une citation
en bloc sont couvertes), délègue à `emit_runs` ci-dessous. python-docx n'a pas
d'API hyperlien : on manipule l'OOXML (relation externe + `<w:hyperlink>`).
"""

from __future__ import annotations

import re

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Patron d'URL CanLII par cour. La langue suit la cour : SCC → /en/ca/scc,
# CSC → /fr/ca/csc (le PDF anglais cite « SCC », le français « CSC »).
_CANLII = {
    "SCC": "https://www.canlii.org/en/ca/scc/doc/{y}/{y}scc{n}/{y}scc{n}.html",
    "CSC": "https://www.canlii.org/fr/ca/csc/doc/{y}/{y}csc{n}/{y}csc{n}.html",
}
# Référence neutre CSC : groupes (année, cour, numéro). Identique à
# parser._CITATION. Le `\b...\d+\b` écarte les marqueurs « [156] » (pas de cour)
# et les années « [2017] » (pas de cour) — seul « AAAA SCC/CSC N » est capté.
_CITE_LINK = re.compile(r"\b(\d{4})\s+(SCC|CSC)\s+(\d+)\b")

# --- Style du lien (constantes uniques, faciles à changer) ------------------- #
# Défaut : bleu hyperlien Word + souligné, la convention reconnaissable
# « cliquable ». Pour un rendu sobre (noir, sans soulignement) :
#   _LINK_COLOR = None   et   _LINK_UNDERLINE = False
# La mise en forme du texte environnant (italique/gras/corps) est toujours
# préservée sur le lien (un lien dans une citation en bloc reste en 10 pt).
_LINK_COLOR = "0563C1"      # hex RVB du bleu hyperlien Word ; None = couleur du texte
_LINK_UNDERLINE = True


def canlii_url(year: str, court: str, num: str) -> str:
    """Construit l'URL CanLII d'une référence neutre (aucun appel réseau)."""
    return _CANLII[court].format(y=year, n=num)


def _add_hyperlink(para, url, text, *, italic=False, bold=False, size=None) -> None:
    """Ajoute un run-hyperlien à la **fin** du paragraphe.

    `para.add_run()` et `para._p.append(...)` ajoutent tous deux en fin de
    `<w:p>` : en traitant les segments de gauche à droite et en les ajoutant
    aussitôt, l'ordre d'origine est préservé. `relate_to` déduplique les
    relations externes identiques (une seule entrée .rels par URL).
    """
    r_id = para.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    if _LINK_COLOR:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), _LINK_COLOR)
        rPr.append(c)
    if _LINK_UNDERLINE:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)
    if italic:
        rPr.append(OxmlElement("w:i"))
    if bold:
        rPr.append(OxmlElement("w:b"))
    if size is not None:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size.pt * 2)))  # taille en demi-points
        rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")  # ne pas perdre d'espace de bord
    t.text = text
    r.append(t)
    link.append(r)
    para._p.append(link)


def emit_runs(para, runs, size=None) -> None:
    """Écrit les fragments stylés dans `para`, en transformant chaque référence
    neutre CSC en hyperlien CanLII (le reste en runs normaux), dans l'ordre.

    `size` (un `Pt`, ou None) force le corps des runs (texte en retrait, plus
    petit). Les citations tiennent dans un **seul** run roman (le nom de cause
    italique est un run distinct), donc le découpage intra-run suffit : on
    n'essaie pas de recoller une citation coupée entre deux runs (cas très rare).
    """
    for r in runs:
        pos = 0
        for m in _CITE_LINK.finditer(r.text):
            before = r.text[pos:m.start()]
            if before:
                run = para.add_run(before)
                run.italic, run.bold = r.italic, r.bold
                if size is not None:
                    run.font.size = size
            year, court, num = m.groups()
            _add_hyperlink(
                para, canlii_url(year, court, num), m.group(0),
                italic=r.italic, bold=r.bold, size=size,
            )
            pos = m.end()
        rest = r.text[pos:]
        if rest or pos == 0:  # le reste, OU le run entier s'il n'a aucune citation
            run = para.add_run(rest if pos else r.text)
            run.italic, run.bold = r.italic, r.bold
            if size is not None:
                run.font.size = size
