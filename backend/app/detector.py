"""
Legacy Bengali font detector for DOCX documents.

Detects whether a DOCX file uses legacy Bijoy/ANSI Bengali fonts
(e.g., SutonnyMJ, SutonnyOMJ, Nikosh) and identifies which text runs
use these fonts so they can be selectively converted to Unicode.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run


# ============================================================
# Known legacy Bijoy/ANSI Bengali fonts
# ============================================================
# These fonts use ANSI encoding to represent Bengali characters.
# Text in these fonts needs to be converted to Unicode before
# PDF generation.

LEGACY_BIJOY_FONTS: set[str] = {
    # SutonnyMJ family
    "SutonnyMJ",
    "SutonnyOMJ",
    "SutonnyMJ Bold",
    "SutonnyMJ Regular",
    "SutonnyOMJ Bold",
    "SutonnyOMJ Regular",
    "SutonnySushreeMJ",
    "SutonnyBanglaOMJ",

    # Nikosh family
    "Nikosh",
    "NikoshBAN",
    "NikoshLight",
    "NikoshLightBan",
    "Nikosh Grammar",

    # Other common Bijoy fonts
    "BanglaKontho",
    "BanglaMJ",
    "AdarshaLipiNormal",
    "AponaLohit",
    "BenSenHandwriting",
    "BenSen",
    "Charukola Unicode",
    "KalpurushANSI",
    "Kalpurush",
    "SolaimanLipi",
    "BorakMJ",
    "ChandrabatiMJ",
    "JugantorMJ",
    "KalerKanthoMJ",
    "ProthomAloMJ",
    "SaijaMJ",
    "SurmaANSI",
    "BijoyBanlaFont",
}

# Fonts that are already Unicode Bengali - do NOT convert
UNICODE_BENGALI_FONTS: set[str] = {
    "Noto Sans Bengali",
    "Noto Serif Bengali",
    "Kalpurush",
    "SolaimanLipi",
    "Mukti",
    "Vrinda",
    "Shonar Bangla",
    "Lohit Bengali",
    "Akaash",
    "Bangla",
    "Mitra Mono",
    "Likhan",
}


def _normalize_font_name(font_name: str) -> str:
    """Normalize a font name for comparison.

    Removes extra whitespace and handles case-insensitive matching.
    """
    if not font_name:
        return ""
    return re.sub(r"\s+", " ", font_name.strip())


def is_legacy_bijoy_font(font_name: str | None) -> bool:
    """Check if a font name is a known legacy Bijoy/ANSI font.

    Args:
        font_name: The font name to check.

    Returns:
        True if the font is a known legacy Bijoy font.
    """
    if not font_name:
        return False

    normalized = _normalize_font_name(font_name)

    # Check exact match
    if normalized in LEGACY_BIJOY_FONTS:
        return True

    # Check case-insensitive match
    lower = normalized.lower()
    for legacy_font in LEGACY_BIJOY_FONTS:
        if legacy_font.lower() == lower:
            return True

    # Check if font name contains known Bijoy indicators
    bijoy_indicators = ["sutonnymj", "sutonnyomj", "banlamj", "borakmj", "prothomalo"]
    for indicator in bijoy_indicators:
        if indicator in lower:
            return True

    return False


@dataclass
class FontDetectionResult:
    """Result of font detection on a DOCX document."""

    has_legacy_fonts: bool = False
    legacy_font_names: set[str] = field(default_factory=set)
    has_unicode_bengali: bool = False
    unicode_font_names: set[str] = field(default_factory=set)
    total_runs: int = 0
    legacy_runs: int = 0


def _get_run_font_name(run: Run) -> str | None:
    """Extract the font name from a run, checking multiple sources.

    DOCX can specify fonts in several places:
    1. Direct run formatting (rPr/rFonts)
    2. Paragraph style
    3. Document defaults

    Args:
        run: A python-docx Run object.

    Returns:
        The font name, or None if not found.
    """
    # Check direct run formatting first
    if run.font and run.font.name:
        return run.font.name

    # Check run XML for rFonts element
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            # Check various font attributes
            for attr in ["w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"]:
                font_name = rfonts.get(qn(attr))
                if font_name:
                    return font_name

    return None


def _get_paragraph_default_font(paragraph: Paragraph) -> str | None:
    """Get the default font for a paragraph from its style.

    Args:
        paragraph: A python-docx Paragraph object.

    Returns:
        The default font name, or None if not found.
    """
    if paragraph.style and paragraph.style.font and paragraph.style.font.name:
        return paragraph.style.font.name
    return None


def _get_document_default_font(doc: Document) -> str | None:
    """Get the document's default font.

    Args:
        doc: A python-docx Document object.

    Returns:
        The default font name, or None if not found.
    """
    try:
        styles = doc.styles
        default_style = styles["Normal"]
        if default_style and default_style.font and default_style.font.name:
            return default_style.font.name
    except (KeyError, AttributeError):
        pass
    return None


def detect_fonts(doc: Document) -> FontDetectionResult:
    """Detect all fonts used in a DOCX document.

    Scans all paragraphs and runs in the document (including headers,
    footers, and tables) to identify legacy Bijoy fonts and Unicode
    Bengali fonts.

    Args:
        doc: A python-docx Document object.

    Returns:
        FontDetectionResult with details about detected fonts.
    """
    result = FontDetectionResult()
    doc_default_font = _get_document_default_font(doc)

    # Collect all paragraphs from body, headers, footers, and tables
    all_paragraphs: list[Paragraph] = []

    # Body paragraphs
    all_paragraphs.extend(doc.paragraphs)

    # Table cell paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    # Header/footer paragraphs
    for section in doc.sections:
        if section.header:
            all_paragraphs.extend(section.header.paragraphs)
        if section.footer:
            all_paragraphs.extend(section.footer.paragraphs)

    for paragraph in all_paragraphs:
        para_default_font = _get_paragraph_default_font(paragraph)

        for run in paragraph.runs:
            result.total_runs += 1
            font_name = _get_run_font_name(run)

            # Fall back to paragraph style, then document default
            if not font_name:
                font_name = para_default_font
            if not font_name:
                font_name = doc_default_font

            if font_name:
                if is_legacy_bijoy_font(font_name):
                    result.has_legacy_fonts = True
                    result.legacy_font_names.add(font_name)
                    result.legacy_runs += 1
                elif font_name in UNICODE_BENGALI_FONTS:
                    result.has_unicode_bengali = True
                    result.unicode_font_names.add(font_name)

    return result
