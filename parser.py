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
_ParaData = "tuple[List[TextRun], List[str]]"


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


def _extract_paragraphs(pages) -> Dict[int, "_ParaData"]:
    """Extrait, par numéro, les fragments stylés du paragraphe et ses sous-titres.

    Travail **ligne à ligne** sur les mots (avec leur police) — et non sur le
    texte aplati — pour (a) isoler les sous-titres et (b) conserver les
    italiques/gras. Un sous-titre candidat (patron de plan) n'est validé que
    s'il précède immédiatement un marqueur « [N] » ; sinon il est réabsorbé dans
    la prose (écarte « N. Metallic, … » et autres initiales d'auteurs).
    """
    body_lines: List[tuple[str, List[dict]]] = []  # (texte de ligne, mots)
    for p in pages:
        for row in _lines(p.extract_words(extra_attrs=["fontname"])):
            body_lines.append((row["text"], row["words"]))

    paras: Dict[int, dict] = {}  # n -> {"runs": [...], "headings": [...]}
    order: List[int] = []
    expected = 1
    current: Optional[int] = None
    buffer: List[tuple[str, List[dict]]] = []  # sous-titres candidats

    for text, words in body_lines:
        m = _LINE_MARKER.match(text)
        if m and int(m.group(1)) == expected:
            n = expected
            body_words = words
            if body_words and re.fullmatch(r"\[\d+\]", body_words[0]["text"]):
                body_words = body_words[1:]  # retire le marqueur « [N] »
            paras[n] = {
                "runs": _line_runs(body_words),
                "headings": [t.strip() for t, _w in buffer],
            }
            order.append(n)
            buffer = []
            current = n
            expected += 1
            continue
        if _HEADING.match(text):
            buffer.append((text, words))
            continue
        # Prose : des titres candidats en attente étaient en fait de la prose.
        if buffer and current is not None:
            for _t, w in buffer:
                paras[current]["runs"].extend(_line_runs(w))
        buffer = []
        if current is not None:
            paras[current]["runs"].extend(_line_runs(words))

    if buffer and current is not None:  # candidats résiduels en toute fin
        for _t, w in buffer:
            paras[current]["runs"].extend(_line_runs(w))

    result: Dict[int, "_ParaData"] = {}
    for i, n in enumerate(order):
        runs = _normalize_runs(paras[n]["runs"])
        if i + 1 == len(order):  # dernier paragraphe → nettoyer la queue
            runs = _strip_tail_runs(runs)
        result[n] = (runs, paras[n]["headings"])
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
    runs, headings = data
    text = _runs_text(runs).strip()
    return Paragraph(
        number=number,
        text=text,
        contains_quote=any(q in text for q in _QUOTE_CHARS),
        contains_citation=bool(_CITATION_HINT.search(text)),
        headings=headings,
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
