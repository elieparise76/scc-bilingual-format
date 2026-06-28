"""Résolution référence neutre → item ID Lexum.

Interroge le moteur de recherche Decisia (decisions.scc-csc.ca) pour retrouver
l'item ID entier à partir d'une référence neutre.

Exemples : '2024 SCC 5' / '2024 CSC 5' → 20264.

URL du endpoint de recherche (iframe interne, pas de JS requis) :
    https://decisions.scc-csc.ca/scc-csc/{lang}/d/s/index.do
        ?cont={citation}&col=1&iframe=true

Le premier résultat dont la balise <span class="citation"> correspond exactement à
la référence recherchée (normalisation espaces + SCC/CSC) donne l'item ID.
"""

from __future__ import annotations

import re
from typing import Optional
from urllib.parse import quote

import httpx

_SEARCH_URL = "https://decisions.scc-csc.ca/scc-csc/{lang}/d/s/index.do"

# Le site renvoie 403 sans User-Agent de navigateur — même contrainte que le fetcher.
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
    )
}


class CitationError(Exception):
    """Levée quand la référence neutre ne peut pas être résolue en item ID."""


def _normalise(citation: str) -> str:
    """Normalise une référence : espaces, casse, SCC et CSC traités comme équivalents."""
    return re.sub(r"\s+", " ", citation.strip()).upper().replace("CSC", "SCC")


def _search_lang(citation: str) -> str:
    """Détermine la langue de recherche : 'fr' pour CSC, 'en' pour SCC."""
    upper = citation.upper()
    # 'CSC' sans 'SCC' (la forme SCC ne contient pas 'CSC' comme sous-chaîne)
    return "fr" if "CSC" in upper and "SCC" not in upper else "en"


def _extract_pairs(html: str) -> list[tuple[int, str]]:
    """Extrait les paires (item_id, citation) depuis le HTML des résultats Decisia.

    Chaque résultat a la structure suivante dans un bloc <h3> :
        <a href="/scc-csc/scc-csc/{lang}/item/{id}/index.do...">Titre</a>
        <span class="citation">{citation}</span>
    """
    pairs: list[tuple[int, str]] = []
    for block in re.findall(r"<h3[^>]*>(.*?)</h3>", html, re.DOTALL):
        m_id = re.search(r"/item/(\d+)/index\.do", block)
        m_cit = re.search(r'class="citation">\s*([^<]+?)\s*</span>', block)
        if m_id and m_cit:
            pairs.append((int(m_id.group(1)), m_cit.group(1).strip()))
    return pairs


def resolve_item_id(
    citation: str, client: Optional[httpx.Client] = None
) -> int:
    """Résout une référence neutre en item ID Lexum.

    Args:
        citation: référence neutre, ex. '2024 SCC 5' ou '2024 CSC 5'.
        client: client httpx existant à réutiliser (évite d'ouvrir une nouvelle session).

    Returns:
        L'item ID entier, ex. 20264.

    Raises:
        CitationError: si aucune correspondance exacte n'est trouvée.
    """
    citation_norm = _normalise(citation)
    lang = _search_lang(citation)
    url = _SEARCH_URL.format(lang=lang)

    own_client = client is None
    if own_client:
        client = httpx.Client()

    try:
        resp = client.get(
            url,
            params={"cont": citation.strip(), "col": "1", "iframe": "true"},
            headers=_HEADERS,
            follow_redirects=True,
            timeout=30.0,
        )
        resp.raise_for_status()
    finally:
        if own_client:
            client.close()

    pairs = _extract_pairs(resp.text)
    for item_id, found_cit in pairs:
        if _normalise(found_cit) == citation_norm:
            return item_id

    found_list = [c for _, c in pairs[:5]]
    raise CitationError(
        f"Référence {citation!r} introuvable dans les résultats Decisia. "
        f"Premiers résultats : {found_list}"
    )
