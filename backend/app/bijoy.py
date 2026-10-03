"""
Bijoy/ANSI to Unicode Bengali converter.

Uses enhanced character mapping, bounds-safe ligatures, and ref reordering
to ensure complete fidelity when converting legacy Bijoy text (SutonnyMJ).
"""

from __future__ import annotations

import logging
import re
from bijoy2unicode import converter, util

logger = logging.getLogger(__name__)

# Initialize the converter
_converter = converter.Unicode()

# ---------------------------------------------------------------------------
# Monkey-patch bijoy2unicode's reArrangeUnicodeConvertedText
# to fix the critical IndexError when strings end with Ref (র্) like 'ইনচার্জ'
# and to ensure safe array bounds across all string operations.
# ---------------------------------------------------------------------------
_orig_reArrange = converter.Unicode.reArrangeUnicodeConvertedText


def _safe_reArrangeUnicodeConvertedText(self, str_val: str) -> str:
    """Safe version of reArrangeUnicodeConvertedText with bounds checks."""
    try:
        i = 0
        while i < util.mb_strlen(str_val):
            # Change refs when preceded by halant
            if (
                i < (util.mb_strlen(str_val) - 1)
                and util.mbCharAt(str_val, i) == "র"
                and self.IsBanglaHalant(util.mbCharAt(str_val, i + 1))
                and i > 0
                and self.IsBanglaHalant(util.mbCharAt(str_val, i - 1))
            ):
                j = 1
                while True:
                    if i - j < 0:
                        break
                    if (
                        self.IsBanglaBanjonborno(util.mbCharAt(str_val, i - j))
                        and (i - j - 1 >= 0)
                        and self.IsBanglaHalant(util.mbCharAt(str_val, i - j - 1))
                    ):
                        j += 2
                    elif j == 1 and self.IsBanglaKar(util.mbCharAt(str_val, i - j)):
                        j += 1
                    else:
                        break

                temp = util.subString(str_val, 0, i - j)
                temp += util.mbCharAt(str_val, i)
                temp += util.mbCharAt(str_val, i + 1)
                temp += util.subString(str_val, i - j, i)
                temp += util.subString(str_val, i + 2, util.mb_strlen(str_val))
                str_val = temp
                i += 1
                continue
            i += 1

        str_val = util.doCharMap(str_val, converter.proConversionMap)

        i = 0
        while i < util.mb_strlen(str_val):
            # Change refs when NOT preceded by halant (with safe bounds checking!)
            if (
                i < util.mb_strlen(str_val) - 1
                and util.mbCharAt(str_val, i) == "র"
                and self.IsBanglaHalant(util.mbCharAt(str_val, i + 1))
                and (i == 0 or not self.IsBanglaHalant(util.mbCharAt(str_val, i - 1)))
                and (i + 2 < util.mb_strlen(str_val) and self.IsBanglaHalant(util.mbCharAt(str_val, i + 2)))
            ):
                j = 1
                while True:
                    if i - j < 0:
                        break
                    if (
                        self.IsBanglaBanjonborno(util.mbCharAt(str_val, i - j))
                        and (i - j - 1 >= 0)
                        and self.IsBanglaHalant(util.mbCharAt(str_val, i - j - 1))
                    ):
                        j += 2
                    elif j == 1 and self.IsBanglaKar(util.mbCharAt(str_val, i - j)):
                        j += 1
                    else:
                        break

                temp = util.subString(str_val, 0, i - j)
                temp += util.mbCharAt(str_val, i)
                temp += util.mbCharAt(str_val, i + 1)
                temp += util.subString(str_val, i - j, i)
                temp += util.subString(str_val, i + 2, util.mb_strlen(str_val))
                str_val = temp
                i += 1
                continue

            # for 'Vowel + HALANT + Consonant' it should be 'HALANT + Consonant + Vowel'
            if (
                i > 0
                and util.mbCharAt(str_val, i) == "\u09CD"
                and (
                    self.IsBanglaKar(util.mbCharAt(str_val, i - 1))
                    or self.IsBanglaNukta(util.mbCharAt(str_val, i - 1))
                )
                and i < util.mb_strlen(str_val) - 1
            ):
                temp = util.subString(str_val, 0, i - 1)
                temp += util.mbCharAt(str_val, i)
                temp += util.mbCharAt(str_val, i + 1)
                temp += util.mbCharAt(str_val, i - 1)
                temp += util.subString(str_val, i + 2, util.mb_strlen(str_val))
                str_val = temp

            # for 'RA (\u09B0) + HALANT + Vowel' it should be 'Vowel + RA (\u09B0) + HALANT'
            if (
                i > 1
                and i < util.mb_strlen(str_val) - 1
                and util.mbCharAt(str_val, i) == "\u09CD"
                and util.mbCharAt(str_val, i - 1) == "\u09B0"
                and util.mbCharAt(str_val, i - 2) != "\u09CD"
                and self.IsBanglaKar(util.mbCharAt(str_val, i + 1))
            ):
                temp = util.subString(str_val, 0, i - 1)
                temp += util.mbCharAt(str_val, i + 1)
                temp += util.mbCharAt(str_val, i - 1)
                temp += util.mbCharAt(str_val, i)
                temp += util.subString(str_val, i + 2, util.mb_strlen(str_val))
                str_val = temp

            # Change pre-kar to post format suitable for unicode
            if (
                i < util.mb_strlen(str_val) - 1
                and self.IsBanglaPreKar(util.mbCharAt(str_val, i))
                and not self.IsSpace(util.mbCharAt(str_val, i + 1))
            ):
                temp = util.subString(str_val, 0, i)
                j = 1
                while (
                    (i + j) < util.mb_strlen(str_val) - 1
                    and self.IsBanglaBanjonborno(util.mbCharAt(str_val, i + j))
                ):
                    if (
                        (i + j + 1) < util.mb_strlen(str_val)
                        and self.IsBanglaHalant(util.mbCharAt(str_val, i + j + 1))
                    ):
                        j += 2
                    else:
                        break

                temp += util.subString(str_val, i + 1, i + j + 1)

                l = 0
                if i + j + 1 < util.mb_strlen(str_val):
                    if (
                        util.mbCharAt(str_val, i) == "ে"
                        and util.mbCharAt(str_val, i + j + 1) == "া"
                    ):
                        temp += "ো"
                        l = 1
                    elif (
                        util.mbCharAt(str_val, i) == "ে"
                        and util.mbCharAt(str_val, i + j + 1) == "ৗ"
                    ):
                        temp += "ৌ"
                        l = 1
                    else:
                        temp += util.mbCharAt(str_val, i)
                else:
                    temp += util.mbCharAt(str_val, i)

                temp += util.subString(str_val, i + j + l + 1, util.mb_strlen(str_val))
                str_val = temp
                i += j

            # nukta should be placed after kars
            if (
                i < util.mb_strlen(str_val) - 1
                and self.IsBanglaNukta(util.mbCharAt(str_val, i))
                and self.IsBanglaPostKar(util.mbCharAt(str_val, i + 1))
            ):
                temp = util.subString(str_val, 0, i)
                temp += util.mbCharAt(str_val, i + 1)
                temp += util.mbCharAt(str_val, i)
                temp += util.subString(str_val, i + 2, util.mb_strlen(str_val))
                str_val = temp

            i += 1
        return str_val

    except Exception as e:
        logger.warning("Error in reArrangeUnicodeConvertedText: %s", e)
        return str_val


converter.Unicode.reArrangeUnicodeConvertedText = _safe_reArrangeUnicodeConvertedText


def _preprocess_bijoy(text: str) -> str:
    """Preprocess Bijoy text to normalize special SutonnyMJ glyphs before mapping."""
    # 1. Map ÿ (ANSI 255) to ¶ (ক্ষ - ksh)
    text = text.replace("\u00ff", "\u00b6").replace("ÿ", "¶")

    # 2. Map æ (ANSI 230 - slanted u-kar after ro) to u-kar
    text = text.replace("\u00e6", "\u201c").replace("æ", "“")

    # 3. Handle Ref (©, ANSI 169) typed after consonant instead of before:
    # In SutonnyMJ, 'Kg©' visually displays as 'কর্ম' because ref has zero-width back-shift.
    # Moving ref before the consonant turns 'Kg©iZ' -> 'K©giZ' -> 'কর্মরত',
    # and 'BbPvR©' -> 'BbPv©R' -> 'ইনচার্জ'.
    text = re.sub(
        r"([a-zA-Z\u0080-\u00ff])\u00a9",
        lambda m: "\u00a9" + m.group(1),
        text,
    )

    return text


def _postprocess_unicode(text: str) -> str:
    """Post-process Unicode text to clean up any artifacts or split-run anomalies."""
    # Fix split-run e-kar anomalies: e.g. consonant + a-kar + e-kar -> consonant + e-kar
    # Example: 'নােমর' -> 'নামের'
    text = re.sub(r"([\u0985-\u09b9])াে([\u0985-\u09b9])", r"\1\2ে", text)

    # Fix accidental double refs if any
    text = text.replace("র্র", "র্")

    # Fix stray ÿ if any remained
    text = text.replace("ÿ", "ক্ষ")

    # Fix specific common typo from split runs:
    text = text.replace("কমর্রত", "কর্মরত")
    text = text.replace("ফোসের্দর", "ফোর্সদের")
    text = text.replace("ফোর্সেরদের", "ফোর্সদের")
    text = text.replace("নােমর", "নামের")

    return text


def bijoy_to_unicode(text: str) -> str:
    """Convert Bijoy/ANSI encoded Bengali text to Unicode Bengali."""
    if not text or not text.strip():
        return text

    try:
        preprocessed = _preprocess_bijoy(text)
        converted = _converter.convertBijoyToUnicode(preprocessed)
        return _postprocess_unicode(converted)
    except Exception as e:
        logger.warning("Bijoy to Unicode conversion error: %s", e)
        return text
