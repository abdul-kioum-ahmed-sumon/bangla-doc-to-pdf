"""
Bijoy/ANSI to Unicode Bengali converter.

Uses the bijoy2unicode library for the actual character mapping,
with additional utilities for DOCX-level text conversion.
"""

from bijoy2unicode import converter


# Initialize the converter once
_converter = converter.Unicode()


def bijoy_to_unicode(text: str) -> str:
    """Convert Bijoy/ANSI encoded Bengali text to Unicode Bengali.

    Uses the bijoy2unicode library which handles:
    - Full character mapping (vowels, consonants, conjuncts)
    - Vowel sign reordering (ই-কার, এ-কার placement)
    - Hasanta/virama handling
    - Bengali digit conversion
    - Punctuation mapping

    Args:
        text: Bijoy-encoded text string (e.g., from SutonnyMJ font).

    Returns:
        Unicode Bengali text string.
    """
    if not text or not text.strip():
        return text

    try:
        return _converter.convertBijoyToUnicode(text)
    except Exception:
        # If conversion fails, return original text rather than crash
        return text
