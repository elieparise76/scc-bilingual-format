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
class Paragraph:
    number: int
    text: str
    contains_quote: bool = False
    contains_citation: bool = False
    # Sous-titres du plan (ex. « II. Contexte », « A. … », « (1) … ») qui
    # précèdent ce paragraphe dans le corps. Plusieurs niveaux peuvent
    # s'empiler avant un même paragraphe.
    headings: List[str] = field(default_factory=list)


@dataclass
class Section:
    type: SectionType
    author: str  # ex. "Le juge Wagner"
    paragraphs: List[Paragraph] = field(default_factory=list)


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
