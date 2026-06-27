"""Phase 5 — CLI.

Orchestre le pipeline complet :
    item_id → fetch → parse (×2) → align → render
"""

from __future__ import annotations

import argparse
import sys

from aligner import align
from fetcher import fetch_pdfs, fetch_pdfs_from_files
from parser import parse_pdf
from renderer import DocMetadata, render_docx


def default_output_name(item_id: str) -> str:
    """'20264' → 'scc_20264_bilingue.docx'."""
    return f"scc_{item_id}_bilingue.docx"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Génère un .docx bilingue d'une décision de la Cour suprême du Canada."
    )
    p.add_argument(
        "item_id",
        help='Item ID Lexum de la décision, ex. "20264" '
        "(visible dans l'URL decisions.scc-csc.ca/.../item/<ID>/index.do).",
    )
    p.add_argument(
        "--lang-order",
        choices=["fr", "en"],
        default="fr",
        help="Langue affichée à gauche (défaut: fr).",
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

    if args.pdf_en and args.pdf_fr:
        pdf_en, pdf_fr = fetch_pdfs_from_files(args.pdf_en, args.pdf_fr)
    else:
        pdf_en, pdf_fr = fetch_pdfs(args.item_id)

    decision_en = parse_pdf(pdf_en)
    decision_fr = parse_pdf(pdf_fr)

    sections = align(decision_fr, decision_en)

    output_path = args.output or default_output_name(args.item_id)
    metadata = DocMetadata(
        title_fr=decision_fr.title,
        title_en=decision_en.title,
        citation_fr=decision_fr.neutral_citation,
        citation_en=decision_en.neutral_citation,
        lang_order=args.lang_order,
    )
    written = render_docx(sections, metadata, output_path)
    print(f"Document généré : {written}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
