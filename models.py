"""Structures de données partagées par tout le pipeline.

Le Parser produit des `Decision`, l'Aligner produit des `ParagraphPair`,
le Renderer consomme les deux.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class SectionType(str, Enum):
    MAJORITY = "majority"
    CONCURRING = "concurring"
    DISSENT = "dissent"
    HEADNOTES = "headnotes"
    OTHER = "other"


@dataclass
class TextRun:
    """Un fragment de texte avec sa mise en forme inline (italique/gras).

    Permet de reproduire dans le `.docx` les italiques (noms d'arrêts, termes
    latins, emphase…) et le gras présents dans le PDF source. Le soulignement
    et le surlignage ne sont pas exposés par pdfplumber (et n'apparaissent pas
    dans les PDF officiels de la CSC).
    """

    text: str
    italic: bool = False
    bold: bool = False


@dataclass
class Paragraph:
    number: int
    text: str  # texte brut (heuristiques, nettoyage de queue)
    contains_quote: bool = False
    contains_citation: bool = False
    # Sous-titres du plan (ex. « II. Contexte », « A. … », « (1) … ») qui
    # précèdent ce paragraphe dans le corps. Plusieurs niveaux peuvent
    # s'empiler avant un même paragraphe.
    headings: List[str] = field(default_factory=list)
    # Texte découpé en fragments stylés (italique/gras) pour le rendu fidèle.
    runs: List[TextRun] = field(default_factory=list)


@dataclass
class Section:
    type: SectionType
    author: str  # ex. "Le juge Wagner" (depuis le bloc d'opinions de la couverture)
    paragraphs: List[Paragraph] = field(default_factory=list)
    # Mention d'attribution verbatim du corps, ex. « English version of the
    # judgment delivered by\nTHE COURT — » ou « Le jugement suivant a été rendu
    # par\nLA COUR — ». Reproduite telle quelle dans le bandeau d'opinion.
    lead_in: str = ""


@dataclass
class Decision:
    title: str
    neutral_citation: str  # ex. "2024 CSC 5"
    sections: List[Section] = field(default_factory=list)

    @property
    def paragraphs(self) -> List[Paragraph]:
        """Tous les paragraphes, toutes sections confondues."""
        return [p for s in self.sections for p in s.paragraphs]


@dataclass
class ParagraphPair:
    """Une paire FR/EN alignée par numéro de paragraphe.

    L'un des deux côtés peut être None si le paragraphe n'existe que dans
    une seule version.
    """

    number: int
    section_type: SectionType
    fr: Optional[Paragraph] = None
    en: Optional[Paragraph] = None


@dataclass
class AlignedSection:
    """Une section dont les paragraphes FR/EN sont appariés.

    Porte l'auteur dans les deux langues (« La juge Martin » / « Martin J. »)
    pour permettre au Renderer d'écrire des bandeaux bilingues.
    """

    type: SectionType
    author_fr: str
    author_en: str
    pairs: List[ParagraphPair] = field(default_factory=list)
    # Mentions d'attribution verbatim (par. ci-dessus), une par langue.
    lead_in_fr: str = ""
    lead_in_en: str = ""
