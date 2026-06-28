"""Phase 2 — Parser.

Transforme un PDF de décision CSC en structure `Decision` exploitable
(métadonnées, sections d'opinion, paragraphes numérotés).

Stratégie (la mise en page CSC est un gabarit stable) :

1. **Métadonnées** (titre, citation) — depuis la page de couverture. L'en-tête
   est sur deux colonnes (titre/citation à gauche, dates/dossier à droite) ;
   on isole la colonne de gauche par position horizontale (x0) car la citation
   peut être coupée par la colonne de droite (« R. c. Wolfe, 2024 » … « CSC 34 »).

2. **Structure des opinions** — depuis le bloc « REASONS FOR JUDGMENT / MOTIFS »
   sous le CORAM. Ce bloc liste, pour chaque opinion, son type, son auteur et sa
   *plage de paragraphes* (« paras. 1 to 92 » / « par. 93 à 144 »). Il est aussi
   sur deux colonnes (libellé à gauche, auteur à droite) — séparation par x0.
   C'est la source de vérité pour découper en sections (plus fiable que de
   détecter les en-têtes dans le corps).

3. **Paragraphes** — marqueurs « [N] » en début de ligne, séquentiels à partir
   de [1] (ce qui écarte les années entre crochets « [2017] » des citations).
   Chaque paragraphe est rattaché à une section selon sa plage.
"""

from __future__ import annotations

import io
import re
from collections import Counter
from dataclasses import dataclass
from typing import Dict, List, Optional

import pdfplumber

from models import Decision, Paragraph, Section, SectionType, TextBlock, TextRun

# Seuils de séparation des colonnes (coordonnée x0, en points).
_HEADER_COL_SPLIT = 310   # en-tête p.1 : colonne de droite (dates/dossier) ~325
_OPINION_COL_SPLIT = 160  # bloc d'opinions : colonne auteur ~177

_PARA_RANGE = re.compile(
    r"\(par(?:as|a)?\.?\s*(\d+)\s*(?:to|à|–|-)\s*(\d+)\s*\)", re.IGNORECASE
)
_CITATION = re.compile(r"\b(\d{4})\s+(SCC|CSC)\s+(\d+)\b")
# Ligne de marqueur de paragraphe « [N] … » (ancrée en début de ligne).
_LINE_MARKER = re.compile(r"^\s*\[(\d+)\]\s?")
# Patron de sous-titre du plan CSC : « II. … », « A. … », « (1) … »,
# « (a) … », « (i) … ». Même police/taille/marge que le corps — c'est le
# patron + la position (juste avant un [N]) qui les identifient.
# Patron de sous-titre numéroté. La parenthèse ouvrante est **optionnelle** :
# l'anglais écrit « (a) … » / « (1) … » mais le français « a) … » / « (1) … ».
_HEADING = re.compile(
    r"^\s*(?:[IVXL]+\.|[A-Z]\.|\(?\d+\)|\(?[a-z]\)|\(?[ivxl]+\))\s+\S"
)

# Queue à retirer du dernier paragraphe : dispositif puis liste des procureurs.
# Dispositif : « Appeal allowed », « Pourvoi accueilli »… (Majuscule volontaire
# pour ne pas confondre avec « the appeal » / « le pourvoi » dans le texte).
_DISPOSITION = re.compile(
    r"\b(?:Appeals?|Pourvois?)\b[^.]{0,90}?"
    r"\b(?:allowed|dismissed|accueillis?|rejetés?|accueilli|rejeté)\b"
)
# Bloc des avocats/procureurs (toujours en toute fin).
_COUNSEL = re.compile(r"\b(?:Solicitors?|Procureurs?|Counsel|Avocats?)\b")
_LABEL_KEYWORD = re.compile(
    r"\b(REASONS|MOTIFS|JUDGMENT|JUGEMENT|DISSENTING|CONCURRING|"
    r"DISSIDENT|CONCORDANT|JOINT|CONJOINT)\b",
    re.IGNORECASE,
)
_BODY_LABELS = re.compile(r"^(BETWEEN|AND BETWEEN|ENTRE|ET ENTRE|- and -|- et -)\b")

# --- Détection du retrait (« texte en retrait » : citations en bloc, listes) ---
# La mise en page CSC place le corps au ras de la marge (x0 ≈ 90) et indente les
# citations en bloc / extraits législatifs à ~148 (offset +58) et les listes
# imbriquées à ~162 (+72). On classe chaque ligne en niveau de retrait via
# l'écart à la marge de base (calculée par document, voir _extract_paragraphs).
_INDENT_L2 = 64  # offset (pts) au-delà duquel : niveau 2 (~162, liste imbriquée)
_INDENT_L1 = 34  # offset au-delà duquel : niveau 1 (~148, citation en bloc)
_INDENT_LIST = 16  # offset minimal pour qu'une **ligne à puce** compte comme L1
# Surcroît d'interligne (pts) — par rapport à l'interligne propre du bloc — au-delà
# duquel on ouvre un nouvel alinéa dans un retrait. Relatif (et non absolu) car
# l'interligne varie : ~2 pt dans une citation (un alinéa saute à ~14), ~16 pt
# dans une liste du corps à interligne double (aucun saut anormal entre lignes).
_GAP_PARA = 7.0
# Puce / numéro de liste en début de ligne (extrait législatif ou énumération) :
# « (4) », « (a) », « (iv) », « a) », « i) », « 1. », « 570 (1) », « 320.24(4) ».
# Sert à ouvrir un bloc à chaque item ; volontairement étroit (pas de simple
# nombre « 27 octobre » qui débuterait une ligne de continuation d'une source).
_LIST_ITEM = re.compile(
    r"^\s*(?:"
    r"\(\d{1,3}\)"            # (4) (12)
    r"|\([a-z]{1,2}\)"        # (a) (bb)
    r"|\([ivxlcdm]+\)"        # (iv)
    r"|[a-z]\)"               # a) b)   (style français : parenthèse seule)
    r"|[ivxlcdm]+\)"          # i) ii)
    r"|\d{1,3}\."             # 1. 2.   (énumération, ex. l'image)
    r"|\d[\d.]*\s*\(\d{1,3}\)"  # 570 (1), 320.24(4)
    r")\s"
)
# Une puce n'ouvre un bloc que si la ligne précédente **termine** un item
# (« ; », « . », « : »…) : distingue un vrai item « b) … » (précédé de « … ; »)
# d'un fragment « (3) ou 320.15(2) » qui poursuit « … 320.14(2) ou ».
_LIST_PREV_END = re.compile(r"[.;:!?»”’)\]]\s*$|—\s*$")

# Détection citations/guillemets (heuristique, pour formatage distinct).
_QUOTE_CHARS = ("«", "»", "“", "”", "[TRANSLATION]", "[TRADUCTION]")
_CITATION_HINT = re.compile(
    r"\b\d{4}\s+(?:SCC|CSC)\s+\d+\b"  # citation neutre
    r"|\[\d{4}\]\s*\d*\s*(?:S\.?C\.?R\.?|R\.?C\.?S\.?)"  # recueils
    r"|\bpara(?:s)?\.\s*\d+|\bpar\.\s*\d+",  # renvois
)


@dataclass
class _Opinion:
    type: SectionType
    author: str
    start: int
    end: int


# --------------------------------------------------------------------------- #
# Utilitaires de mise en page
# --------------------------------------------------------------------------- #
def _lines(words: List[dict], tol: float = 3.0) -> List[dict]:
    """Regroupe les mots en lignes (par coordonnée `top`), triées par x0."""
    rows: List[dict] = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if rows and abs(w["top"] - rows[-1]["top"]) <= tol:
            rows[-1]["words"].append(w)
        else:
            rows.append({"top": w["top"], "words": [w]})
    for r in rows:
        r["words"].sort(key=lambda w: w["x0"])
        r["text"] = " ".join(w["text"] for w in r["words"])
    return rows


def _col_text(row: dict, split: float, side: str) -> str:
    """Texte d'une ligne du côté gauche ('left') ou droit ('right') de `split`."""
    if side == "left":
        ws = [w for w in row["words"] if w["x0"] < split]
    else:
        ws = [w for w in row["words"] if w["x0"] >= split]
    return " ".join(w["text"] for w in ws)


# --------------------------------------------------------------------------- #
# 1. Métadonnées
# --------------------------------------------------------------------------- #
def _parse_header(page) -> tuple[str, str]:
    """Extrait (titre, citation neutre) depuis la couverture."""
    left_parts: List[str] = []
    for row in _lines(page.extract_words()):
        if _BODY_LABELS.match(row["text"]):
            break
        left_parts.append(_col_text(row, _HEADER_COL_SPLIT, "left"))
    blob = re.sub(r"\s+", " ", " ".join(left_parts)).strip()

    cite_m = _CITATION.search(blob)
    citation = (
        f"{cite_m.group(1)} {cite_m.group(2)} {cite_m.group(3)}" if cite_m else ""
    )

    label_m = re.search(r"(CITATION|R[ÉE]F[ÉE]RENCE)\s*:", blob)
    start = label_m.end() if label_m else 0
    end = cite_m.start() if cite_m else len(blob)
    title = blob[start:end].strip().rstrip(",").strip()
    return title, citation


def _cover_right_text(page) -> str:
    """Texte de la colonne droite de la couverture (dates/dossier), recadré.

    On recadre car le label sur deux lignes (« JUGEMENT RENDU ») serait sinon
    entrelacé par `extract_text()` global."""
    crop = page.crop((310, 230, page.width, 480))
    return re.sub(r"[ \t]+", " ", crop.extract_text() or "")


def _find_cover_date(text: str, label: str) -> str:
    """Date suivant un libellé ; `re.S` car elle chevauche le saut de ligne."""
    m = re.search(label + r"\s*:?\s*(.+?\d{4})", text, re.S)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


_HELD = re.compile(
    r"(?:Held|Arrêt)\s*(?:\([^)]*\))?\s*:\s*.{1,400}?\.(?=\s+[A-ZÀ-Ü]|\s*$)", re.S
)


def _extract_held(pages) -> str:
    """Mention « Held (…): … » / « Arrêt (…) : … » (1re phrase = dispositif)."""
    full = "\n".join((p.extract_text() or "") for p in pages[2:8])
    m = _HELD.search(full)
    return re.sub(r"\s+", " ", m.group(0)).strip() if m else ""


def _is_italic(word: dict) -> bool:
    return "Italic" in (word.get("fontname") or "")


def _extract_appeal_and_catchwords(pages) -> tuple[str, str]:
    """(« ON APPEAL FROM … », mots-clés) depuis le sommaire.

    Ancre = « ON APPEAL FROM » / « EN APPEL DE » (mot APPEAL/APPEL en capitales).
    La mention d'appel = mots **non italiques** jusqu'au 1er italique ; les
    mots-clés = le **bloc italique** qui suit (le sommaire en prose qui suit est
    en romain → frontière nette)."""
    words: List[dict] = []
    for pg in pages[2:10]:
        words += pg.extract_words(extra_attrs=["fontname"])

    anchor = next(
        (i for i, w in enumerate(words) if w["text"] in ("APPEAL", "APPEL")), None
    )
    if anchor is None:
        return "", ""

    i = max(0, anchor - 1)  # inclure « ON » / « EN »
    appeal: List[str] = []
    while i < len(words) and not _is_italic(words[i]):
        appeal.append(words[i]["text"])
        i += 1

    catch: List[str] = []
    while i < len(words):
        w = words[i]
        if _is_italic(w):
            catch.append(w["text"])
        elif re.search(r"[A-Za-zÀ-ÿ]{2,}", w["text"]):
            break  # 1er mot de prose (romain) → fin des mots-clés
        else:
            catch.append(w["text"])  # ponctuation entre fragments italiques
        i += 1

    appeal_from = re.sub(r"\s+", " ", " ".join(appeal)).strip()
    catchwords = re.sub(r"\s+", " ", " ".join(catch)).strip()
    return appeal_from, catchwords


# --------------------------------------------------------------------------- #
# 2. Structure des opinions
# --------------------------------------------------------------------------- #
def _classify(label: str) -> SectionType:
    L = label.upper()
    dissent = "DISSID" in L or "DISSENT" in L
    concur = "CONCORD" in L or "CONCURR" in L
    if dissent:
        return SectionType.DISSENT
    if concur:
        return SectionType.CONCURRING
    return SectionType.MAJORITY


def _clean_author(author: str) -> str:
    a = _PARA_RANGE.sub("", author)
    a = re.sub(r"\b\d+\)", "", a)  # fragments de plage ayant débordé (« 144) »)
    a = re.sub(r"\s+", " ", a).strip()
    return a


def _find_coram_page(pages):
    for p in pages[:5]:
        if "CORAM" in (p.extract_text() or ""):
            return p
    return None


def _parse_opinions(pages) -> List[_Opinion]:
    page = _find_coram_page(pages)
    if page is None:
        return []

    rows = _lines(page.extract_words())
    # Région du bloc : entre la ligne CORAM et la ligne NOTE/note de bas (*).
    block: List[dict] = []
    seen_coram = False
    for r in rows:
        if not seen_coram:
            if "CORAM" in r["text"]:
                seen_coram = True
            continue
        if r["text"].startswith("NOTE") or r["text"].startswith("*"):
            break
        block.append(r)

    opinions: List[_Opinion] = []
    cur: Optional[dict] = None
    for r in block:
        label = _col_text(r, _OPINION_COL_SPLIT, "left")
        author = _col_text(r, _OPINION_COL_SPLIT, "right")
        if cur is None:
            if not _LABEL_KEYWORD.search(label):
                continue  # débordement du CORAM (liste des juges)
            cur = {"label": [label], "author": [author], "full": [r["text"]]}
        else:
            cur["label"].append(label)
            cur["author"].append(author)
            cur["full"].append(r["text"])

        m = _PARA_RANGE.search(" ".join(cur["full"]))
        if m:
            opinions.append(
                _Opinion(
                    type=_classify(" ".join(cur["label"])),
                    author=_clean_author(" ".join(cur["author"])),
                    start=int(m.group(1)),
                    end=int(m.group(2)),
                )
            )
            cur = None
    return opinions


# --------------------------------------------------------------------------- #
# 3. Paragraphes
# --------------------------------------------------------------------------- #
# Type de la sortie de _extract_paragraphs : n -> (blocs, titres, candidats).
_ParaData = "tuple[List[TextBlock], List[str], List[str]]"


def _word_style(word: dict) -> tuple[bool, bool]:
    """(italique, gras) d'après le nom de police (ex. TimesNewRomanPS-ItalicMT)."""
    f = (word.get("fontname") or "").lower()
    return ("italic" in f or "oblique" in f), ("bold" in f)


def _line_runs(words: List[dict]) -> List[TextRun]:
    """Fragments stylés d'une ligne ; chaque mot porte une espace finale."""
    runs: List[TextRun] = []
    for w in words:
        italic, bold = _word_style(w)
        token = w["text"] + " "
        if runs and runs[-1].italic == italic and runs[-1].bold == bold:
            runs[-1].text += token
        else:
            runs.append(TextRun(token, italic, bold))
    return runs


def _normalize_runs(runs: List[TextRun]) -> List[TextRun]:
    """Espaces simples, fusion des fragments voisins de même style, bords nettoyés."""
    out: List[TextRun] = []
    for r in runs:
        t = re.sub(r"\s+", " ", r.text)
        if not t:
            continue
        if out and out[-1].italic == r.italic and out[-1].bold == r.bold:
            out[-1].text += t
        else:
            out.append(TextRun(t, r.italic, r.bold))
    if out:
        out[0].text = out[0].text.lstrip()
        out[-1].text = out[-1].text.rstrip()
    return [r for r in out if r.text]


def _runs_text(runs: List[TextRun]) -> str:
    return "".join(r.text for r in runs)


def _tail_cut_index(text: str) -> int:
    """Position où couper le dernier paragraphe (dispositif + procureurs)."""
    counsel = _COUNSEL.search(text)
    region = text[: counsel.start()] if counsel else text
    disp = _DISPOSITION.search(region)
    if disp:
        return disp.start()
    if counsel:
        return counsel.start()
    return len(text)


# Fin de phrase, deux formes :
#   • « .?! » + guillemets fermants éventuels (« threshold. », « Act? », « … 393). ») ;
#   • une parenthèse/crochet fermant précédé d'un caractère **autre que « . »**
#     (« … p. 154) » = fin de citation).
# On exclut ainsi « … JJ.A.) » / « … para. 38.] » (« .) » / « .] » = abréviation
# dans une parenthèse, débordement de titre) qui ne sont PAS des fins de phrase.
_SENT_END = re.compile(r"[.?!][\"'”’»]*$|[^.\s][)\]][\"'”’»]*$")
_COLON_END = re.compile(r":\s*$")
_DASH_END = re.compile(r"—\s*$")  # ligne d'auteur d'opinion (lead-in)
_BARE_MARK = re.compile(r"\[\d+\]")


def _find_boundary(lines: List[tuple]) -> Optional[int]:
    """Index de la dernière ligne de **prose** (fin de phrase ou « : »).

    Les lignes au **préfixe de plan** sont ignorées : un titre peut finir par
    « ? » (« A. … de la Loi? ») sans être une fin de prose. Le débordement d'un
    titre finissant par « .) » (« … JJ.A.) ») n'est pas non plus une frontière —
    c'est `_SENT_END` qui l'écarte (« .) » ≠ fin de phrase, contrairement à
    « ). »). Tout ce qui suit la frontière = le bloc de sous-titres.

    NB : on n'essaie PAS de « suivre » les débordements de titre ici. Une ligne
    de prose-citation commençant par une initiale (« J. P. J. Maingot, … »)
    matche `_HEADING` par accident ; la sauter est sans conséquence (elle ne
    finit pas une phrase), mais propager un état « titre en cours » avalerait la
    vraie fin de phrase suivante."""
    boundary = None
    for i, (text, _w) in enumerate(lines):
        if _HEADING.match(text):
            continue
        t = text.rstrip()
        if _SENT_END.search(t) or _COLON_END.search(t):
            boundary = i
    return boundary


def _strip_lead_in(block: List[tuple]) -> List[tuple]:
    """Retire le lead-in d'opinion (jusqu'à la dernière ligne d'auteur « — »)."""
    start = 0
    for i, (text, _w) in enumerate(block):
        if _DASH_END.search(text.rstrip()):
            start = i + 1
    return block[start:]


def _is_heading_block(block: List[tuple]) -> bool:
    """Le bloc commence-t-il par un vrai sous-titre ?

    Un sous-titre commence soit par un **préfixe de plan**, soit (non numéroté)
    par une **ligne courte commençant par une lettre** (« Charter Interpretation »,
    « Definition »). Un bloc cité qui a fui (prose pleine largeur, ou ligne de
    citation « (Voir… » / note « [Nous soulignons.] ») est rejeté → il reste dans
    le paragraphe précédent."""
    for text, words in _strip_lead_in(block):
        t = text.strip()
        if not t:
            continue
        if _HEADING.match(text):
            return True
        return _line_x1(words) <= _FULL_WIDTH and t[:1].isalpha()
    return False


def _block_split(lines: List[tuple]) -> tuple:
    """(prose, bloc-titre) d'un segment. Le bloc-titre = lignes après la
    frontière de prose, **si elles forment un vrai bloc de sous-titres**. Sinon
    tout reste en prose : frontière finissant par « : » (liste/citation
    énumérée), ou bloc cité ayant fui (`_is_heading_block`)."""
    b = _find_boundary(lines)
    if b is None or _COLON_END.search(lines[b][0].rstrip()):
        return lines, []
    tail = lines[b + 1 :]
    if not _is_heading_block(tail):
        return lines, []
    return lines[: b + 1], tail


_FULL_WIDTH = 490  # x1 au-delà duquel une ligne est « pleine » (donc déborde)
# Une opinion longue imprime sa propre table des matières (avec n° de page) entre
# le lead-in et son 1er paragraphe : ce n'est PAS un sous-titre à capter.
_PRINTED_TOC = re.compile(r"TABLE\s+(?:OF\s+CONTENTS|DES\s+MATIÈRES)", re.IGNORECASE)
# Année seule entre parenthèses → signature de citation (« (2023), 101 R. … »).
# « (2022 QCCA 185) » d'un vrai titre ne matche pas (chiffres après l'année).
_CITATION_YEAR = re.compile(r"\((?:19|20)\d{2}\)")


def _line_x1(words: List[dict]) -> float:
    return max((w["x1"] for w in words), default=0.0)


def _split_into_headings(block: List[tuple]) -> List[str]:
    """Regroupe un bloc-titre en sous-titres.

    Retire d'abord le lead-in d'opinion (jusqu'à la dernière ligne d'auteur
    finissant par « — »). Puis : un nouveau sous-titre commence à chaque ligne au
    préfixe de plan, ou à un titre **non numéroté** (« Charter Interpretation »,
    « Definition »). Une ligne sans préfixe ne **prolonge** le titre précédent
    que si celui-ci était **pleine largeur** (il a donc débordé) — sinon c'est un
    nouveau titre non numéroté. Cela distingue « a) Souveraineté… / parlementaire »
    (débordement) de « A. … Privilege / Definition » (deux titres)."""
    block = _strip_lead_in(block)
    # Table des matières imprimée de l'opinion → on n'en tire aucun sous-titre.
    if any(_PRINTED_TOC.search(text) for text, _w in block):
        return []

    headings: List[str] = []
    prev_full = False
    for text, words in block:
        t = re.sub(r"\s+", " ", text).strip()
        if not t:
            continue
        is_prefix = bool(_HEADING.match(text))
        if headings and not is_prefix and prev_full:
            headings[-1] = f"{headings[-1]} {t}"  # débordement du titre précédent
        else:
            headings.append(t)  # préfixe, 1re ligne, ou nouveau titre non numéroté
        prev_full = _line_x1(words) > _FULL_WIDTH
    # Écarte les fragments de citation captés à tort : parenthèses imbriquées
    # (« … (G. Boniface)). ») ou année entre parenthèses (« … (2023), 101 R. du
    # B. can. … ») — un vrai sous-titre n'a ni « )) » ni « (AAAA) » isolée.
    return [h for h in headings if "))" not in h and not _CITATION_YEAR.search(h)]


def _indent_level(x0: float, baseline: float, is_list: bool) -> int:
    """Niveau de retrait d'une ligne d'après l'écart de son x0 à la marge.

    `is_list` (la ligne débute par une puce/numéro) abaisse le seuil : une
    énumération faiblement indentée compte quand même comme « en retrait »."""
    off = x0 - baseline
    if off >= _INDENT_L2:
        return 2
    if off >= _INDENT_L1:
        return 1
    if off >= _INDENT_LIST and is_list:
        return 1
    return 0


def _group_indent_blocks(lines: List[tuple], baseline: float) -> List[dict]:
    """Regroupe les lignes d'un paragraphe en blocs par niveau de retrait.

    Le **niveau de rendu** d'un bloc = niveau du x0 *minimal* de ses lignes :
    une ligne plus indentée (1er retrait d'un nouvel alinéa de citation, à ~162)
    est ainsi absorbée par le bloc dont le corps revient à ~148.

    Un nouveau bloc s'ouvre quand : (a) on franchit la frontière corps/retrait ;
    (b) un **saut vertical** marque un changement d'alinéa ; (c) une ligne en
    retrait débute par une **puce/numéro** (nouvel item de liste ou de citation)."""
    groups: List[dict] = []
    cur: Optional[dict] = None
    prev_bottom: Optional[float] = None
    prev_text = ""
    for text, words in lines:
        if not words:
            continue
        x0, top, bottom = words[0]["x0"], words[0]["top"], words[0]["bottom"]
        gap = (top - prev_bottom) if prev_bottom is not None else None
        is_list = bool(_LIST_ITEM.match(text))
        lvl = _indent_level(x0, baseline, is_list)
        if cur is None:
            boundary = True
        elif (cur["level"] >= 1) != (lvl >= 1):
            boundary = True  # (a) corps <-> retrait
        elif (
            lvl >= 1
            and cur["base_gap"] is not None
            and gap is not None
            and gap > cur["base_gap"] + _GAP_PARA
        ):
            # (b) nouvel alinéa *dans* un retrait : un saut nettement plus grand
            # que l'interligne propre du bloc. Comparaison **relative** (et non
            # absolue) car l'interligne varie : citations à interligne simple
            # (~2 pt, un alinéa saute à ~14), listes du corps à interligne double
            # (~16 pt partout — aucun saut anormal, donc aucune coupe parasite).
            boundary = True
        elif is_list and lvl >= 1 and _LIST_PREV_END.search(prev_text.rstrip()):
            # (c) nouvel item de liste/citation : puce précédée d'une fin d'item
            # (évite de couper un renvoi « (3) ou … » au fil d'une phrase).
            boundary = True
        else:
            boundary = False

        if boundary:
            cur = {"lines": [words], "min_x0": x0, "level": lvl, "base_gap": None}
            groups.append(cur)
        else:
            cur["lines"].append(words)
            if cur["base_gap"] is None and gap is not None and gap >= 0:
                cur["base_gap"] = gap  # 1er interligne intra-bloc = référence
            if x0 < cur["min_x0"]:  # le corps du bloc fixe son niveau de rendu
                cur["min_x0"] = x0
                cur["level"] = _indent_level(x0, baseline, False)
        prev_bottom = bottom
        prev_text = text
    return groups


def _lines_to_blocks(lines: List[tuple], baseline: float) -> List[TextBlock]:
    """Construit les blocs (prose + retraits) d'un paragraphe ; retire « [N] ».

    Le marqueur « [N] » est **conservé** pour le calcul du retrait (il est au ras
    de la marge, x0 ≈ baseline) : il ancre la 1re ligne au niveau 0 malgré le
    retrait de 1re ligne du texte (le texte est tabulé après le marqueur, à ~148,
    alors que les lignes de continuation reviennent à la marge). On ne retire le
    marqueur que des runs, juste avant de les normaliser."""
    blocks: List[TextBlock] = []
    for gi, grp in enumerate(_group_indent_blocks(lines, baseline)):
        runs: List[TextRun] = []
        for li, words in enumerate(grp["lines"]):
            if gi == 0 and li == 0 and words and _BARE_MARK.fullmatch(words[0]["text"]):
                words = words[1:]  # retire le mot-marqueur « [N] »
            runs.extend(_line_runs(words))
        runs = _normalize_runs(runs)
        if runs:
            blocks.append(TextBlock(runs=runs, indent=grp["level"]))
    return blocks


def _strip_tail_blocks(blocks: List[TextBlock]) -> List[TextBlock]:
    """Tronque les blocs du dernier paragraphe au dispositif/liste des procureurs.

    Le dispositif (« Pourvoi accueilli ») et la liste des procureurs sont en
    retrait dans le PDF (donc capturés comme blocs) : on les retire en coupant à
    l'indice calculé par `_tail_cut_index` sur le texte à plat des blocs."""
    flat = "".join(_runs_text(b.runs) for b in blocks)
    cut = _tail_cut_index(flat)
    if cut >= len(flat):
        return blocks
    out: List[TextBlock] = []
    total = 0
    for b in blocks:
        btext = _runs_text(b.runs)
        if total + len(btext) <= cut:
            out.append(b)
            total += len(btext)
        else:
            kept = _strip_tail_runs_to(b.runs, cut - total)
            if kept:
                out.append(TextBlock(runs=kept, indent=b.indent))
            break
    return out


def _strip_tail_runs_to(runs: List[TextRun], keep: int) -> List[TextRun]:
    """Tronque une liste de runs aux `keep` premiers caractères (bords nettoyés)."""
    out: List[TextRun] = []
    total = 0
    for r in runs:
        if total + len(r.text) <= keep:
            out.append(r)
            total += len(r.text)
        else:
            head = r.text[: keep - total]
            if head:
                out.append(TextRun(head, r.italic, r.bold))
            break
    if out:
        out[-1].text = out[-1].text.rstrip()
    return [r for r in out if r.text]


def _candidate_block(lines: List[tuple]) -> List[tuple]:
    """Bloc-titre **permissif** en fin de segment (avant le marqueur suivant).

    Trouvé en **remontant** : la dernière ligne (juste avant « [N] ») est toujours
    incluse — c'est le titre ou sa dernière ligne de débordement ; on continue de
    remonter sur les lignes au **préfixe de plan** et sur les lignes qui ne
    **terminent pas** une phrase (débordements de titre, titre non numéroté qui
    déborde) ; on s'arrête à la 1re ligne de **prose** finissant une phrase.

    Récupère les titres que `_block_split` rejette (garde `_is_heading_block`, ou
    frontière placée sur un débordement finissant par « ). », ex. « (la juge
    Dysart). »). N'est utilisé qu'en **réconciliation** (l'autre langue confirme)."""
    block: List[tuple] = []
    for text, words in reversed(lines):
        if _LINE_MARKER.match(text):
            break  # ne pas avaler la ligne du marqueur « [N] »
        t = text.rstrip()
        ends = bool(_SENT_END.search(t) or _COLON_END.search(t))
        if block and not _HEADING.match(text) and ends:
            break  # ligne de prose au-dessus du bloc de titres
        block.append((text, words))
    block.reverse()
    return block


def _extract_paragraphs(pages) -> Dict[int, "_ParaData"]:
    """Extrait, par numéro, les fragments stylés du paragraphe et ses sous-titres.

    Travail **ligne à ligne** sur les mots (avec police) pour conserver les
    italiques/gras. Découpe le corps en segments par marqueur « [N] » séquentiel
    (écarte les années « [2017] »). Le **bloc de sous-titres** d'un paragraphe =
    tout ce qui suit la dernière phrase du paragraphe précédent, jusqu'à son
    marqueur — voir `_block_split` / `_split_into_headings`."""
    body_lines: List[tuple[str, List[dict]]] = []
    for p in pages:
        for row in _lines(p.extract_words(extra_attrs=["fontname"])):
            body_lines.append((row["text"], row["words"]))

    baseline = _body_baseline(body_lines)

    # Segments : lignes de [N] (incluse) jusqu'avant [N+1]. `pre` = avant [1].
    pre: List[tuple] = []
    segments: List[dict] = []
    expected = 1
    cur: Optional[dict] = None
    for text, words in body_lines:
        m = _LINE_MARKER.match(text)
        if m and int(m.group(1)) == expected:
            cur = {"n": expected, "lines": [(text, words)]}
            segments.append(cur)
            expected += 1
        elif cur is None:
            pre.append((text, words))
        else:
            cur["lines"].append((text, words))

    if not segments:
        return {}

    # Le bloc-titre en fin d'un segment appartient au paragraphe SUIVANT.
    _, pending = _block_split(pre)  # titres confirmés du 1er paragraphe
    pending_cand = _candidate_block(pre)  # titres candidats (réconciliation)
    result: Dict[int, "_ParaData"] = {}
    for i, seg in enumerate(segments):
        prose, tail = _block_split(seg["lines"])
        blocks = _lines_to_blocks(prose, baseline)
        if i + 1 == len(segments):  # dernier paragraphe → nettoyer la queue
            blocks = _strip_tail_blocks(blocks)
        result[seg["n"]] = (
            blocks,
            _split_into_headings(pending),
            _split_into_headings(pending_cand),
        )
        pending = tail
        pending_cand = _candidate_block(seg["lines"])
    return result


def _body_baseline(body_lines: List[tuple]) -> float:
    """Marge de gauche du corps (x0 le plus fréquent) ; ~90 pour les PDF CSC.

    Sert d'origine au calcul du retrait. Le corps domine largement les lignes,
    donc le x0 modal est la marge de base même en présence de citations/listes."""
    counts = Counter(
        round(words[0]["x0"]) for _t, words in body_lines if words
    )
    return float(counts.most_common(1)[0][0]) if counts else 90.0


# Ligne d'attribution d'opinion (« THE COURT — », « LA JUGE MOREAU — »).
_AUTHOR_DASH = "—"
_LEAD_IN_LOOKBACK = 80  # lignes max à remonter pour trouver l'attribution


def _extract_lead_ins(pages, opinions: List[_Opinion]) -> Dict[int, str]:
    """Mention d'attribution verbatim avant la 1re ¶ de chaque opinion.

    Ancrée sur la ligne d'auteur en fin de « — » (« THE COURT — ») ; le
    préambule = les lignes au-dessus jusqu'à la dernière ligne finissant par
    « . » (fin de la liste des avocats). Reproduite « préambule\\nauteur ».
    """
    lines = "\n".join((p.extract_text() or "") for p in pages).split("\n")
    result: Dict[int, str] = {}
    for op in opinions:
        idx = next(
            (i for i, l in enumerate(lines) if re.match(rf"^\s*\[{op.start}\]\s", l)),
            None,
        )
        if idx is None:
            result[op.start] = ""
            continue
        j = idx - 1
        while j >= 0 and idx - j <= _LEAD_IN_LOOKBACK and not lines[j].rstrip().endswith(
            _AUTHOR_DASH
        ):
            j -= 1
        if j < 0 or idx - j > _LEAD_IN_LOOKBACK:
            result[op.start] = ""
            continue
        author_line = lines[j].strip()
        preamble: List[str] = []
        k = j - 1
        while k >= 0 and lines[k].strip() and not lines[k].rstrip().endswith("."):
            preamble.insert(0, lines[k].strip())
            k -= 1
        pre = re.sub(r"\s+", " ", " ".join(preamble)).strip()
        result[op.start] = f"{pre}\n{author_line}" if pre else author_line
    return result


def _make_paragraph(number: int, data: "_ParaData") -> Paragraph:
    blocks, headings, candidates = data
    runs = [r for b in blocks for r in b.runs]
    text = " ".join(_runs_text(b.runs) for b in blocks).strip()
    return Paragraph(
        number=number,
        text=text,
        contains_quote=any(q in text for q in _QUOTE_CHARS),
        contains_citation=bool(_CITATION_HINT.search(text)),
        headings=headings,
        heading_candidates=candidates,
        runs=runs,
        blocks=blocks,
    )


# --------------------------------------------------------------------------- #
# Assemblage
# --------------------------------------------------------------------------- #
def _build_sections(
    opinions: List[_Opinion],
    paras: Dict[int, "_ParaData"],
    lead_ins: Dict[int, str],
) -> List[Section]:
    if not opinions:
        # Repli : une seule section majoritaire regroupant tous les paragraphes.
        all_paras = [_make_paragraph(n, paras[n]) for n in sorted(paras)]
        return [Section(type=SectionType.MAJORITY, author="", paragraphs=all_paras)]

    sections: List[Section] = []
    for op in opinions:
        sec_paras = [
            _make_paragraph(n, paras[n])
            for n in range(op.start, op.end + 1)
            if n in paras
        ]
        sections.append(
            Section(
                type=op.type,
                author=op.author,
                paragraphs=sec_paras,
                lead_in=lead_ins.get(op.start, ""),
            )
        )
    return sections


def parse_pdf(pdf_bytes: bytes) -> Decision:
    """Extrait la structure d'une décision depuis un PDF.

    Args:
        pdf_bytes: contenu du PDF en mémoire.

    Returns:
        Decision avec ses sections et paragraphes numérotés.
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages = pdf.pages
        title, citation = _parse_header(pages[0])
        right = _cover_right_text(pages[0])
        hearing_date = _find_cover_date(right, r"(?:APPEALS? HEARD|APPELS? ENTENDUS?)")
        decision_date = _find_cover_date(right, r"(?:JUDGMENT RENDERED|JUGEMENT RENDU)")
        appeal_from, catchwords = _extract_appeal_and_catchwords(pages)
        held = _extract_held(pages)
        opinions = _parse_opinions(pages)
        paras = _extract_paragraphs(pages)
        lead_ins = _extract_lead_ins(pages, opinions)

    return Decision(
        title=title,
        neutral_citation=citation,
        sections=_build_sections(opinions, paras, lead_ins),
        hearing_date=hearing_date,
        decision_date=decision_date,
        appeal_from=appeal_from,
        catchwords=catchwords,
        held=held,
    )
