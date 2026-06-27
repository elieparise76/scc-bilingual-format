"""Phase 3 — Aligner.

Apparie les paragraphes des versions française et anglaise par numéro, en
regroupant le résultat par section (avec l'auteur dans les deux langues).
"""

from __future__ import annotations

from typing import Dict, List

from models import AlignedSection, Decision, Paragraph, ParagraphPair, Section


def _para_index(decision: Decision) -> Dict[int, Paragraph]:
    return {p.number: p for s in decision.sections for p in s.paragraphs}


def _section_index(decision: Decision) -> Dict[int, Section]:
    return {p.number: s for s in decision.sections for p in s.paragraphs}


def align(decision_fr: Decision, decision_en: Decision) -> List[AlignedSection]:
    """Apparie ¶N_FR ↔ ¶N_EN, regroupé par section.

    Un paragraphe présent dans une seule version donne une paire dont l'autre
    côté est None. Le type de section et les auteurs proviennent de la version
    française en priorité (repli sur l'anglaise).

    Args:
        decision_fr: version française parsée.
        decision_en: version anglaise parsée.

    Returns:
        Liste d'AlignedSection, dans l'ordre des numéros de paragraphe.
    """
    fr_paras = _para_index(decision_fr)
    en_paras = _para_index(decision_en)
    fr_sec = _section_index(decision_fr)
    en_sec = _section_index(decision_en)

    all_numbers = sorted(set(fr_paras) | set(en_paras))

    sections: List[AlignedSection] = []
    current_key = None
    for n in all_numbers:
        fs = fr_sec.get(n)
        es = en_sec.get(n)
        ref = fs or es  # au moins un des deux existe
        section_type = ref.type
        author_fr = fs.author if fs else ""
        author_en = es.author if es else ""

        key = (section_type, author_fr, author_en)
        if key != current_key:
            sections.append(
                AlignedSection(
                    type=section_type,
                    author_fr=author_fr,
                    author_en=author_en,
                    lead_in_fr=fs.lead_in if fs else "",
                    lead_in_en=es.lead_in if es else "",
                    pairs=[],
                )
            )
            current_key = key

        sections[-1].pairs.append(
            ParagraphPair(
                number=n,
                section_type=section_type,
                fr=fr_paras.get(n),
                en=en_paras.get(n),
            )
        )
    return sections
