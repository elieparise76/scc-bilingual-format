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

from models import Decision, Paragraph, Section, SectionType

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
_HEADING = re.compile(
    r"^\s*(?:[IVXL]+\.|[A-Z]\.|\(\d+\)|\([a-z]\)|\([ivxl]+\))\s+\S"
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
def _strip_tail(text: str) -> str:
    """Retire du dernier paragraphe le dispositif et la liste des procureurs."""
    counsel = _COUNSEL.search(text)
    region = text[: counsel.start()] if counsel else text
    disp = _DISPOSITION.search(region)
    if disp:
        return text[: disp.start()].strip()
    if counsel:
        return region.strip()
    return text.strip()


def _extract_paragraphs(pages) -> Dict[int, tuple[str, List[str]]]:
    """Extrait, par numéro, le texte du paragraphe et ses sous-titres.

    Travail **ligne à ligne** (et non sur le texte aplati) pour pouvoir isoler
    les sous-titres, qui n'ont aucune marque typographique distinctive. Un
    sous-titre candidat (patron de plan) n'est validé que s'il précède
    immédiatement un marqueur « [N] » : sinon il est réabsorbé dans la prose
    (écarte « N. Metallic, … » et autres initiales d'auteurs dans les citations).
    """
    lines: List[str] = []
    for p in pages:
        lines.extend((p.extract_text() or "").split("\n"))

    # n -> (lignes de texte, sous-titres). order garde l'ordre d'apparition.
    paras: Dict[int, tuple[List[str], List[str]]] = {}
    order: List[int] = []
    expected = 1
    current: Optional[int] = None
    buffer: List[str] = []  # sous-titres candidats, validés au prochain [N]

    for line in lines:
        m = _LINE_MARKER.match(line)
        if m and int(m.group(1)) == expected:
            n = expected
            paras[n] = ([line[m.end():]], buffer)  # buffer = ses sous-titres
            order.append(n)
            buffer = []
            current = n
            expected += 1
            continue
        if _HEADING.match(line):
            buffer.append(line.strip())
            continue
        # Prose : des titres candidats en attente étaient en fait de la prose.
        if buffer:
            if current is not None:
                paras[current][0].extend(buffer)
            buffer = []
        if current is not None:
            paras[current][0].append(line)

    if buffer and current is not None:  # candidats résiduels en toute fin
        paras[current][0].extend(buffer)

    result: Dict[int, tuple[str, List[str]]] = {}
    for i, n in enumerate(order):
        text_lines, headings = paras[n]
        text = re.sub(r"\s+", " ", " ".join(text_lines)).strip()
        if i + 1 == len(order):  # dernier paragraphe → nettoyer la queue
            text = _strip_tail(text)
        result[n] = (text, headings)
    return result


def _make_paragraph(number: int, data: tuple[str, List[str]]) -> Paragraph:
    text, headings = data
    return Paragraph(
        number=number,
        text=text,
        contains_quote=any(q in text for q in _QUOTE_CHARS),
        contains_citation=bool(_CITATION_HINT.search(text)),
        headings=headings,
    )


# --------------------------------------------------------------------------- #
# Assemblage
# --------------------------------------------------------------------------- #
def _build_sections(
    opinions: List[_Opinion], paras: Dict[int, tuple[str, List[str]]]
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
        sections.append(Section(type=op.type, author=op.author, paragraphs=sec_paras))
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
        opinions = _parse_opinions(pages)
        paras = _extract_paragraphs(pages)

    return Decision(
        title=title,
        neutral_citation=citation,
        sections=_build_sections(opinions, paras),
    )
