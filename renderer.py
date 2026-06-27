"""Phase 4 — Renderer.

Génère le document Word (.docx) bilingue à deux colonnes à partir des
sections alignées.

Structure produite :
- en-tête courant **alterné par page** (anglais sur les pages impaires, français
  sur les paires) : « <nom de la décision en italique> — <référence> — <juge
  rédacteur et son rôle sur cette page> », suivi d'une ligne horizontale ;
- pied de page : numéro de page ;
- page de garde (titre + référence) et table des matières des sections ;
- corps : par opinion, un tableau deux colonnes (langue de gauche / langue de
  droite), une ligne par paire de paragraphes, alignées en haut de cellule ;
- bandeaux (sans couleur) au début de chaque opinion ;
- lignes de sous-titres (« II. Contexte », « A. … ») au-dessus des paragraphes.

L'en-tête « juge de cette page » s'appuie sur le champ Word **STYLEREF** : à
chaque début d'opinion on insère deux paragraphes marqueurs *masqués* (un par
langue, styles `OpinionRefEN`/`OpinionRefFR`) contenant « Auteur (rôle) » ;
STYLEREF affiche dans l'en-tête le premier marqueur présent sur la page, ou à
défaut le dernier avant la page — soit exactement l'opinion qui débute (ou se
poursuit) sur la page. Marqueurs au niveau du corps (hors tableau) pour que
STYLEREF les retrouve de façon fiable.

Réglages globaux : Times New Roman partout, interligne simple, texte justifié,
marges 0,4 po, bordures de tableau blanches, espace accru entre les colonnes.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from models import AlignedSection, Paragraph, SectionType

_FONT = "Times New Roman"
_MARGIN = Inches(0.4)
_COL_W = Inches(3.8)  # largeur de chaque colonne (page lettre, marges 0,4 po)
_CELL_LR = Inches(0.12)  # marge interne gauche/droite → espace au centre
_CELL_TB = Inches(0.03)

# Libellés de rôle de section (fr, en) pour bandeaux et table des matières.
_ROLE_LABELS = {
    SectionType.MAJORITY: ("Motifs de la majorité", "Majority reasons"),
    SectionType.CONCURRING: ("Motifs concordants", "Concurring reasons"),
    SectionType.DISSENT: ("Motifs dissidents", "Dissenting reasons"),
    SectionType.HEADNOTES: ("Sommaire", "Headnotes"),
    SectionType.OTHER: ("Autres motifs", "Other reasons"),
}
# Adjectif de rôle (fr, en) pour l'en-tête courant « Auteur (rôle) ».
_ROLE_ADJ = {
    SectionType.MAJORITY: ("majoritaire", "majority"),
    SectionType.CONCURRING: ("concordant", "concurring"),
    SectionType.DISSENT: ("dissident", "dissenting"),
    SectionType.HEADNOTES: ("sommaire", "headnotes"),
    SectionType.OTHER: ("autres", "other"),
}
_COURT_AUTHORS = {"La Cour", "The Court"}


@dataclass
class DocMetadata:
    """Métadonnées d'en-tête pour le document généré."""

    title_fr: str
    title_en: str
    citation_fr: str
    citation_en: str
    lang_order: str = "fr"  # "fr" → FR à gauche ; "en" → EN à gauche

    @property
    def cover_title(self) -> str:
        return self.title_fr if self.lang_order == "fr" else self.title_en

    @property
    def cover_citation(self) -> str:
        return self.citation_fr if self.lang_order == "fr" else self.citation_en


# --------------------------------------------------------------------------- #
# Helpers bas niveau (OOXML)
# --------------------------------------------------------------------------- #
def _add_page_field(paragraph) -> None:
    """Insère un champ « PAGE » (numéro de page dynamique)."""
    _add_field(paragraph, "PAGE")


def _add_field(paragraph, instruction: str):
    """Insère un champ Word arbitraire ; renvoie le run créé."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)
    return run


def _enable_odd_even(doc: Document) -> None:
    settings = doc.settings.element
    if settings.find(qn("w:evenAndOddHeaders")) is None:
        settings.append(OxmlElement("w:evenAndOddHeaders"))


def _para_bottom_border(paragraph) -> None:
    """Ligne horizontale sous le paragraphe (utilisé pour l'en-tête)."""
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pbdr.append(bottom)
    pPr.append(pbdr)


def _set_base_font(doc: Document) -> None:
    """Times New Roman + interligne simple sur les styles utilisés."""
    for name in ("Normal", "Header", "Footer", "List Bullet"):
        if name not in doc.styles:
            continue
        style = doc.styles[name]
        style.font.name = _FONT
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(attr), _FONT)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(11)
    pf = normal.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def _set_table_borders_white(table) -> None:
    """Rend toutes les bordures du tableau blanches (lignes invisibles)."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "FFFFFF")
        borders.append(el)
    tblPr.append(borders)


def _set_cell_margins(table) -> None:
    """Marges internes des cellules → crée l'espace au centre du tableau."""
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (
        ("top", _CELL_TB),
        ("left", _CELL_LR),
        ("bottom", _CELL_TB),
        ("right", _CELL_LR),
    ):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(val.twips)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _create_ref_styles(doc: Document) -> None:
    """Styles masqués portant « Auteur (rôle) » pour les champs STYLEREF."""
    for name in ("OpinionRefEN", "OpinionRefFR"):
        if name in doc.styles:
            continue
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        st.font.size = Pt(1)
        rpr = st.element.get_or_add_rPr()
        rpr.append(OxmlElement("w:vanish"))  # texte masqué (hors flux)
        pf = st.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)


# --------------------------------------------------------------------------- #
# Libellés (rôles, auteurs)
# --------------------------------------------------------------------------- #
def _lead_author(author: str) -> str:
    """« La juge Martin (avec l'accord…) » → « La juge Martin »."""
    return author.split("(")[0].strip() if author else ""


def _section_label(section: AlignedSection, lang: str, unanimous: bool) -> str:
    """Libellé de bandeau/TOC : « Motifs de la majorité — La juge Côté »."""
    if unanimous:
        role = "Motifs de la Cour" if lang == "fr" else "Reasons of the Court"
    else:
        fr, en = _ROLE_LABELS.get(section.type, _ROLE_LABELS[SectionType.OTHER])
        role = fr if lang == "fr" else en
    author = _lead_author(section.author_fr if lang == "fr" else section.author_en)
    if author in _COURT_AUTHORS:  # éviter « Motifs de la Cour — La Cour »
        return role
    return f"{role} — {author}" if author else role


def _author_role(section: AlignedSection, lang: str, unanimous: bool) -> str:
    """Texte du marqueur STYLEREF : « La juge Martin (majoritaire) »."""
    author = _lead_author(section.author_fr if lang == "fr" else section.author_en)
    if unanimous:
        role = "unanime" if lang == "fr" else "unanimous"
    else:
        fr, en = _ROLE_ADJ.get(section.type, _ROLE_ADJ[SectionType.OTHER])
        role = fr if lang == "fr" else en
    return f"{author} ({role})"


# --------------------------------------------------------------------------- #
# En-tête / pied de page
# --------------------------------------------------------------------------- #
def _fill_header(paragraph, title: str, citation: str, ref_style: str) -> None:
    for run in list(paragraph.runs):  # vider sans laisser de run vide
        run._r.getparent().remove(run._r)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = paragraph.add_run(title)
    title_run.italic = True
    paragraph.add_run(f" — {citation} — ")
    _add_field(paragraph, f'STYLEREF "{ref_style}"')
    for run in paragraph.runs:
        run.font.size = Pt(9)
    _para_bottom_border(paragraph)


def _setup_headers_footers(doc: Document, meta: DocMetadata) -> None:
    sec = doc.sections[0]
    odd_header = sec.header
    even_header = sec.even_page_header
    odd_header.is_linked_to_previous = False
    even_header.is_linked_to_previous = False
    # Pages impaires → anglais ; paires → français (page 1 = anglais).
    _fill_header(odd_header.paragraphs[0], meta.title_en, meta.citation_en, "OpinionRefEN")
    _fill_header(even_header.paragraphs[0], meta.title_fr, meta.citation_fr, "OpinionRefFR")

    for footer in (sec.footer, sec.even_page_footer):
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_page_field(fp)
        for run in fp.runs:
            run.font.size = Pt(9)


# --------------------------------------------------------------------------- #
# Page de garde et table des matières
# --------------------------------------------------------------------------- #
def _add_title_block(doc: Document, meta: DocMetadata) -> None:
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(meta.cover_title)
    run.bold = True
    run.font.size = Pt(16)

    cite = doc.add_paragraph()
    cite.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cite.add_run(meta.cover_citation)
    run.font.size = Pt(12)


def _add_toc(
    doc: Document, sections: List[AlignedSection], lang_left: str, unanimous: bool
) -> None:
    lang_right = "en" if lang_left == "fr" else "fr"
    heading = doc.add_paragraph()
    run = heading.add_run("Table des matières / Contents")
    run.bold = True
    run.font.size = Pt(12)

    for sec in sections:
        nums = [p.number for p in sec.pairs]
        rng = f"[{nums[0]}–{nums[-1]}]" if nums else ""
        line = doc.add_paragraph(style="List Bullet")
        line.add_run(f"{_section_label(sec, lang_left, unanimous)}  {rng}").bold = True
        line.add_run(f"\n{_section_label(sec, lang_right, unanimous)}").italic = True


# --------------------------------------------------------------------------- #
# Corps : marqueurs, bandeaux, tableaux
# --------------------------------------------------------------------------- #
def _add_ref_markers(doc: Document, section: AlignedSection, unanimous: bool) -> None:
    """Paragraphes marqueurs masqués pour STYLEREF (en début d'opinion)."""
    doc.add_paragraph(_author_role(section, "en", unanimous), style="OpinionRefEN")
    doc.add_paragraph(_author_role(section, "fr", unanimous), style="OpinionRefFR")


def _add_banner(
    doc: Document, section: AlignedSection, lang_left: str, unanimous: bool
) -> None:
    lang_right = "en" if lang_left == "fr" else "fr"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    top = p.add_run(_section_label(section, lang_left, unanimous))
    top.bold = True
    p.add_run("\n")
    p.add_run(_section_label(section, lang_right, unanimous)).italic = True


def _write_cell(cell, paragraph: Optional[Paragraph], number: int) -> None:
    cell.width = _COL_W
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if paragraph is None:
        run = p.add_run("—")
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        return
    num_run = p.add_run(f"[{number}] ")
    num_run.bold = True
    p.add_run(paragraph.text)


def _add_heading_row(table, pair, lang_left: str) -> None:
    """Ligne de sous-titres au-dessus d'un paragraphe (gras, deux colonnes)."""
    left_para = pair.fr if lang_left == "fr" else pair.en
    right_para = pair.en if lang_left == "fr" else pair.fr
    left_h = left_para.headings if left_para else []
    right_h = right_para.headings if right_para else []
    if not left_h and not right_h:
        return
    row = table.add_row()
    for cell, headings in ((row.cells[0], left_h), (row.cells[1], right_h)):
        cell.width = _COL_W
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(8)
        p.add_run("\n".join(headings)).bold = True


def _add_opinion_table(doc: Document, section: AlignedSection, lang_left: str) -> None:
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.allow_autofit = False
    _set_table_borders_white(table)
    _set_cell_margins(table)
    for pair in section.pairs:
        _add_heading_row(table, pair, lang_left)
        row = table.add_row()
        left_para = pair.fr if lang_left == "fr" else pair.en
        right_para = pair.en if lang_left == "fr" else pair.fr
        _write_cell(row.cells[0], left_para, pair.number)
        _write_cell(row.cells[1], right_para, pair.number)


# --------------------------------------------------------------------------- #
# Point d'entrée
# --------------------------------------------------------------------------- #
def render_docx(
    sections: List[AlignedSection],
    metadata: DocMetadata,
    output_path: str,
) -> str:
    """Produit le fichier .docx bilingue.

    Args:
        sections: sections alignées (sortie de l'Aligner).
        metadata: titres/citations FR+EN, ordre des langues.
        output_path: chemin du fichier .docx à écrire.

    Returns:
        Le chemin du fichier écrit.
    """
    lang_left = metadata.lang_order
    unanimous = len(sections) == 1

    doc = Document()
    for s in doc.sections:
        s.page_width = Inches(8.5)
        s.page_height = Inches(11)
        s.top_margin = s.bottom_margin = _MARGIN
        s.left_margin = s.right_margin = _MARGIN

    _enable_odd_even(doc)
    _set_base_font(doc)
    _create_ref_styles(doc)
    _setup_headers_footers(doc, metadata)

    # Marqueur précoce (masqué) pour la 1re opinion : garantit que STYLEREF se
    # résout dès la page 1 (sinon « Error! No text of specified style »).
    _add_ref_markers(doc, sections[0], unanimous)
    _add_title_block(doc, metadata)
    _add_toc(doc, sections, lang_left, unanimous)

    for section in sections:
        _add_ref_markers(doc, section, unanimous)
        _add_banner(doc, section, lang_left, unanimous)
        _add_opinion_table(doc, section, lang_left)

    doc.save(output_path)
    return output_path
