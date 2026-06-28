"""Phase 5 — CLI.

Orchestre le pipeline complet :
    ref → (résolution) → fetch → parse (×2) → align → render

`ref` accepte un item ID Lexum numérique (ex. '20264') ou une référence neutre
(ex. '2024 SCC 5' / '2024 CSC 5'). La résolution se fait via citation.py.
"""

from __future__ import annotations

import argparse
import re
import sys

from aligner import align
from citation import CitationError, resolve_item_id
from fetcher import fetch_pdfs, fetch_pdfs_from_files
from parser import parse_pdf
from renderer import DocMetadata, render_docx


def default_output_name(item_id: str | int) -> str:
    """'20264' → 'scc_20264_bilingue.docx'."""
    return f"scc_{item_id}_bilingue.docx"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Génère un .docx bilingue d'une décision de la Cour suprême du Canada."
    )
    p.add_argument(
        "ref",
        help=(
            'Item ID Lexum (ex. "20264") ou référence neutre (ex. "2024 SCC 5" / '
            '"2024 CSC 5"). L\'item ID est visible dans l\'URL '
            "decisions.scc-csc.ca/.../item/<ID>/index.do."
        ),
    )
    p.add_argument(
        "--lang-order",
        choices=["fr", "en"],
        default="en",
        help="Langue affichée à gauche (défaut: en).",
    )
    p.add_argument(
        "--output",
        help="Chemin du fichier .docx de sortie (défaut: scc_<item_id>_bilingue.docx).",
    )
    p.add_argument(
        "--pdf-en",
        help="Fallback : chemin local du PDF anglais (court-circuite le fetcher).",
    )
    p.add_argument(
        "--pdf-fr",
        help="Fallback : chemin local du PDF français (court-circuite le fetcher).",
    )
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    # Résolution référence neutre → item ID (si l'entrée n'est pas purement numérique).
    ref = args.ref.strip()
    if ref.isdigit():
        item_id: int | str = ref
    else:
        try:
            item_id = resolve_item_id(ref)
            print(f"Référence {ref!r} → item ID {item_id}")
        except CitationError as exc:
            print(f"Erreur : {exc}", file=sys.stderr)
            return 1

    if args.pdf_en and args.pdf_fr:
        pdf_en, pdf_fr = fetch_pdfs_from_files(args.pdf_en, args.pdf_fr)
    else:
        pdf_en, pdf_fr = fetch_pdfs(item_id)

    decision_en = parse_pdf(pdf_en)
    decision_fr = parse_pdf(pdf_fr)

    sections = align(decision_fr, decision_en)

    output_path = args.output or default_output_name(item_id)
    metadata = DocMetadata(
        title_fr=decision_fr.title,
        title_en=decision_en.title,
        citation_fr=decision_fr.neutral_citation,
        citation_en=decision_en.neutral_citation,
        hearing_fr=decision_fr.hearing_date,
        hearing_en=decision_en.hearing_date,
        date_fr=decision_fr.decision_date,
        date_en=decision_en.decision_date,
        appeal_fr=decision_fr.appeal_from,
        appeal_en=decision_en.appeal_from,
        catchwords_fr=decision_fr.catchwords,
        catchwords_en=decision_en.catchwords,
        held_fr=decision_fr.held,
        held_en=decision_en.held,
        lang_order=args.lang_order,
    )
    written = render_docx(sections, metadata, output_path)
    print(f"Document généré : {written}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
