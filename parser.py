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
from dataclasses import dataclass
from typing import Dict, List, Optional

import pdfplumber

from models import Decision, Paragraph, Section, SectionType, TextRun

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
# Type de la sortie de _extract_paragraphs : n -> (runs, sous-titres).
_ParaData = "tuple[List[TextRun], List[str], List[str]]"  # runs, titres, candidats


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


def _strip_tail_runs(runs: List[TextRun]) -> List[TextRun]:
    """Tronque les runs au dispositif/liste des procureurs (dernier paragraphe)."""
    cut = _tail_cut_index(_runs_text(runs))
    out: List[TextRun] = []
    total = 0
    for r in runs:
        if total + len(r.text) <= cut:
            out.append(r)
            total += len(r.text)
        else:
            keep = cut - total
            if keep > 0:
                out.append(TextRun(r.text[:keep], r.italic, r.bold))
            break
    if out:
        out[-1].text = out[-1].text.rstrip()
    return out


# Fin de phrase, deux formes :
#   • « .?! » + guillemets fermants éventuels (« threshold. », « Act? », « … 393). ») ;
#   • une parenthèse/crochet fermant précédé d'un caractère **autre que « . »**
#     (« … p. 154) » = fin de citation).
# On exclut ainsi « … JJ.A.) » / « … para. 38.] » (« .) » / « .] » = abréviation
# dans une parenthèse, débordement de titre) qui ne sont PAS des fins de phrase.
_SENT_END = re.compile(r"[.?!][\"'”’»]*$|[^.\s][)\]][\"'”’»]*$")
_COLON_END = re.compile(r":\s*$")
_DASH_END = re.compile(r"—\s*$")  # ligne d'auteur d'opinion (lead-in)
# Ligne d'**attribution d'une opinion** : entièrement en CAPITALES, finissant par
# « — » (« LA JUGE MOREAU — », « MARTIN J. — », « THE COURT — »). L'absence de
# minuscule la distingue des nombreuses lignes de prose, de sommaire (mots-clés)
# ou de procureurs qui finissent aussi par un tiret cadratin (« … Négligence
# criminelle — », « … the Act — », « … Ministère de la Justice — »).
_LEAD_IN_AUTHOR = re.compile(r"^(?=[^a-zà-ÿ]*[A-ZÀ-Þ])[^a-zà-ÿ]*—\s*$")
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


def _opinion_break_index(lines: List[tuple]) -> Optional[int]:
    """Index où **débute le matériel inter-opinions** dans un segment, ou None.

    Quand le dernier paragraphe d'une opinion est immédiatement suivi d'une autre
    opinion, son segment englobe (après sa prose) la mention d'attribution de
    l'opinion suivante, sa **table des matières imprimée** éventuelle, puis le 1er
    sous-titre de cette opinion. Repère le plus tôt de : la **ligne d'auteur en
    capitales finissant par « — »** (`_LEAD_IN_AUTHOR`) ou la **table des matières
    imprimée** (`_PRINTED_TOC`). Le préambule du lead-in (« Version française des
    motifs… ») est *au-dessus* de la ligne d'auteur ; il est récupéré par la
    recherche de frontière de prose dans `_block_split` (il ne finit pas une
    phrase, donc il tombe naturellement après la frontière)."""
    for i, (text, _w) in enumerate(lines):
        if _LEAD_IN_AUTHOR.match(text.strip()) or _PRINTED_TOC.search(text):
            return i
    return None


def _block_split(lines: List[tuple]) -> tuple:
    """(prose, bloc-titre) d'un segment. Le bloc-titre = lignes après la
    frontière de prose, **si elles forment un vrai bloc de sous-titres**. Sinon
    tout reste en prose : frontière finissant par « : » (liste/citation
    énumérée), ou bloc cité ayant fui (`_is_heading_block`).

    Cas particulier (dernier ¶ d'une opinion suivi d'une autre) : si le segment
    contient le lead-in / la table des matières de l'opinion suivante
    (`_opinion_break_index`), la frontière de prose est cherchée **uniquement
    avant** ce point — sinon les entrées de la table imprimée (« par. 320.24(4) »,
    dont « 4) » ressemble à une fin de phrase) leurraient `_find_boundary` et tout
    le bloc inter-opinions restait collé à la prose. Tout ce qui suit la frontière
    (préambule + ligne d'auteur + table imprimée + 1er sous-titre de l'opinion
    suivante) part en `tail` ; `_split_into_headings` l'écarte ensuite via
    `_PRINTED_TOC` / `_strip_lead_in`."""
    brk = _opinion_break_index(lines)
    if brk is not None:
        b = _find_boundary(lines[:brk])
        cut = b + 1 if b is not None else brk
        return lines[:cut], lines[cut:]
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


def _lines_to_runs(lines: List[tuple]) -> List[TextRun]:
    """Construit les runs d'un paragraphe ; retire le marqueur « [N] » initial."""
    runs: List[TextRun] = []
    for idx, (_text, words) in enumerate(lines):
        ws = words
        if idx == 0 and ws and _BARE_MARK.fullmatch(ws[0]["text"]):
            ws = ws[1:]
        runs.extend(_line_runs(ws))
    return _normalize_runs(runs)


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
        runs = _lines_to_runs(prose)
        if i + 1 == len(segments):  # dernier paragraphe → nettoyer la queue
            runs = _strip_tail_runs(runs)
        result[seg["n"]] = (
            runs,
            _split_into_headings(pending),
            _split_into_headings(pending_cand),
        )
        pending = tail
        pending_cand = _candidate_block(seg["lines"])
    return result


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
    runs, headings, candidates = data
    text = _runs_text(runs).strip()
    return Paragraph(
        number=number,
        text=text,
        contains_quote=any(q in text for q in _QUOTE_CHARS),
        contains_citation=bool(_CITATION_HINT.search(text)),
        headings=headings,
        heading_candidates=candidates,
        runs=runs,
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
