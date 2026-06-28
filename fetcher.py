"""Phase 1 — Fetcher.

Récupère les deux PDFs (EN + FR) d'une décision CSC à partir de son
*item ID* Lexum ou de sa référence neutre.

Pattern d'URL : le PDF d'une décision se télécharge directement à
    https://decisions.scc-csc.ca/scc-csc/scc-csc/{lang}/{item_id}/1/document.do
où {lang} ∈ {en, fr}. Les deux langues partagent le même item_id.

Si une référence neutre est fournie ('2024 SCC 5'), elle est résolue en item_id
via citation.resolve_item_id avant le téléchargement.
"""

from __future__ import annotations

from typing import Tuple

import httpx

BASE_URL = "https://decisions.scc-csc.ca/scc-csc/scc-csc/{lang}/{item_id}/1/document.do"

# Le site renvoie 403 sans User-Agent de navigateur.
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
    )
}


def pdf_url(item_id: str | int, lang: str) -> str:
    """Construit l'URL du PDF pour un item_id et une langue ('en'|'fr')."""
    return BASE_URL.format(lang=lang, item_id=item_id)


def _download(client: httpx.Client, item_id: str | int, lang: str) -> bytes:
    url = pdf_url(item_id, lang)
    resp = client.get(url, headers=_HEADERS, follow_redirects=True, timeout=30.0)
    resp.raise_for_status()
    content_type = resp.headers.get("content-type", "")
    if "application/pdf" not in content_type:
        raise ValueError(
            f"URL {url} n'a pas renvoyé un PDF (content-type: {content_type!r}). "
            f"item_id={item_id} est-il valide ?"
        )
    return resp.content


def fetch_pdfs(item_id: str | int) -> Tuple[bytes, bytes]:
    """Télécharge les PDFs anglais et français d'une décision CSC.

    Args:
        item_id: item ID Lexum (ex. "20264") ou référence neutre (ex. "2024 SCC 5").
                 Si la valeur contient des lettres, elle est résolue en item ID via
                 citation.resolve_item_id.

    Returns:
        Tuple (pdf_en, pdf_fr) en bytes (gardés en mémoire).
    """
    # Résolution référence neutre → item ID si l'argument n'est pas purement numérique.
    ref = str(item_id).strip()
    if not ref.isdigit():
        from citation import resolve_item_id
        item_id = resolve_item_id(ref)

    with httpx.Client() as client:
        pdf_en = _download(client, item_id, "en")
        pdf_fr = _download(client, item_id, "fr")
    return pdf_en, pdf_fr


def fetch_pdfs_from_files(path_en: str, path_fr: str) -> Tuple[bytes, bytes]:
    """Fallback : charge les deux PDFs depuis le disque (input manuel)."""
    with open(path_en, "rb") as f_en, open(path_fr, "rb") as f_fr:
        return f_en.read(), f_fr.read()
