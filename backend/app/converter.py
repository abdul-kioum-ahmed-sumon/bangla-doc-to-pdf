"""
DOCX to PDF converter with Bijoy/ANSI Bengali support and Dual-Engine Rendering.

Supports two conversion engines:
1. Microsoft Word Native (Windows) — 100% pixel-perfect fidelity ("like a screenshot").
   Uses Word COM automation to render documents with identical layout, kerning,
   margins, and installed fonts (including SutonnyMJ and Unicode).
2. LibreOffice Headless (Linux / Docker / Fallback) — Cross-platform server conversion.
   Converts Bijoy text to Unicode with high-fidelity matching fonts (Kalpurush/Nikosh).
"""

from __future__ import annotations

import asyncio
import logging
import os
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .bijoy import bijoy_to_unicode
from .detector import (
    FontDetectionResult,
    contains_unicode_bengali,
    detect_fonts,
    is_legacy_bijoy_font,
    is_legacy_bijoy_run,
    _get_document_default_font,
    _get_paragraph_default_font,
    _get_run_font_name,
)

logger = logging.getLogger(__name__)

# Conversion timeout in seconds
CONVERSION_TIMEOUT = int(os.environ.get("CONVERSION_TIMEOUT", "120"))

# Conversion engine preference: "auto", "word", "libreoffice"
CONVERSION_ENGINE = os.environ.get("CONVERSION_ENGINE", "auto").lower()

# Check if SutonnyMJ fonts are available on the system
SUTONNYMJ_AVAILABLE = False
_WORD_AVAILABLE: bool | None = None


def check_font_availability() -> dict[str, bool]:
    """Check which Bengali fonts are available on the system.

    Checks fc-list (Linux), system fonts (Windows), and the local fonts/ directory.

    Returns:
        Dictionary mapping font names to availability status.
    """
    global SUTONNYMJ_AVAILABLE

    fonts_to_check = [
        "sutonnymj",
        "sutonnyomj",
        "kalpurush",
        "nikosh",
        "siyamrupali",
        "noto sans bengali",
        "noto serif bengali",
    ]
    fonts_status: dict[str, bool] = {name: False for name in fonts_to_check}

    # 1. Check local repo fonts/ directory
    repo_fonts_dir = Path(__file__).resolve().parent.parent.parent / "fonts"
    repo_font_files: list[str] = []
    if repo_fonts_dir.exists():
        repo_font_files = [f.name.lower() for f in repo_fonts_dir.glob("*.*")]

    # 2. Check Windows Fonts if on Windows
    win_font_files: list[str] = []
    if os.name == "nt":
        win_fonts_dir = Path("C:/Windows/Fonts")
        if win_fonts_dir.exists():
            try:
                win_font_files = [f.name.lower() for f in win_fonts_dir.glob("*.*")]
            except Exception:
                pass

    all_local_files = repo_font_files + win_font_files
    for font_name in fonts_to_check:
        clean_name = font_name.replace(" ", "").replace("-", "")
        if any(clean_name in f.replace(" ", "").replace("-", "") for f in all_local_files):
            fonts_status[font_name] = True

    # 3. Check fc-list on Linux if available
    try:
        result = subprocess.run(
            ["fc-list", ":lang=bn", "family"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        available_fonts = result.stdout.lower()
        for font_name in fonts_to_check:
            if font_name in available_fonts:
                fonts_status[font_name] = True
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    SUTONNYMJ_AVAILABLE = fonts_status.get("sutonnymj", False) or fonts_status.get("sutonnyomj", False)
    return fonts_status


def get_unicode_replacement_font() -> str:
    """Determine the best available Unicode replacement font for Bijoy text.

    Instead of defaulting to generic Noto Sans, we prioritize fonts that match
    SutonnyMJ's classic typography (Kalpurush, Nikosh, Siyam Rupali).

    Returns:
        Font family name string.
    """
    env_font = os.environ.get("UNICODE_REPLACEMENT_FONT")
    if env_font:
        return env_font

    available = check_font_availability()
    # Preference order for typography matching SutonnyMJ
    if available.get("kalpurush"):
        return "Kalpurush"
    if available.get("nikosh"):
        return "Nikosh"
    if available.get("siyamrupali"):
        return "Siyam Rupali"
    if available.get("noto serif bengali"):
        return "Noto Serif Bengali"
    return "Noto Sans Bengali"


def is_word_available() -> bool:
    """Check if Microsoft Word COM automation is available on Windows.

    Returns:
        True if Word can be dispatched via COM, False otherwise.
    """
    global _WORD_AVAILABLE
    if _WORD_AVAILABLE is not None:
        return _WORD_AVAILABLE

    if os.name != "nt":
        _WORD_AVAILABLE = False
        return False

    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.Quit(SaveChanges=False)
            _WORD_AVAILABLE = True
            logger.info("Microsoft Word COM is available. Native high-fidelity rendering enabled.")
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        logger.info("Microsoft Word COM not available: %s", e)
        _WORD_AVAILABLE = False

    return _WORD_AVAILABLE


_BENGALI_DIGITS = str.maketrans("0123456789", "০১২৩৪৫৬৭৮৯")


def to_bengali_number(n: int) -> str:
    """Convert an integer to Bengali numeral string."""
    return str(n).translate(_BENGALI_DIGITS)


def _set_run_font(run: Run, replacement_font: str) -> None:
    """Set the font of a run across OpenXML font attributes and synchronize bold/italic."""
    run.font.name = replacement_font
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            for attr in ["w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"]:
                rfonts.set(qn(attr), replacement_font)
        else:
            rfonts = rpr.makeelement(
                qn("w:rFonts"),
                {
                    qn("w:ascii"): replacement_font,
                    qn("w:hAnsi"): replacement_font,
                    qn("w:cs"): replacement_font,
                },
            )
            rpr.insert(0, rfonts)

        # Synchronize complex script bold (w:bCs) with Latin bold (w:b).
        # In legacy Bijoy documents (SutonnyMJ), Word uses w:b for Latin/ANSI fonts.
        # Word often automatically inserts phantom <w:bCs/> even when the text is regular.
        # When converted to Unicode Bengali (a Complex Script), <w:bCs/> causes regular text
        # to render as bold! Therefore, w:bCs must match w:b.
        b_elem = rpr.find(qn("w:b"))
        is_bold = b_elem is not None and b_elem.get(qn("w:val"), "true").lower() not in ("0", "false", "off")
        bcs_elem = rpr.find(qn("w:bCs"))
        if not is_bold and bcs_elem is not None:
            rpr.remove(bcs_elem)
        elif is_bold and bcs_elem is None:
            rpr.append(rpr.makeelement(qn("w:bCs")))

        # Same for italic (w:iCs vs w:i)
        i_elem = rpr.find(qn("w:i"))
        is_italic = i_elem is not None and i_elem.get(qn("w:val"), "true").lower() not in ("0", "false", "off")
        ics_elem = rpr.find(qn("w:iCs"))
        if not is_italic and ics_elem is not None:
            rpr.remove(ics_elem)
        elif is_italic and ics_elem is None:
            rpr.append(rpr.makeelement(qn("w:iCs")))


def _clean_phantom_complex_script_styles(paragraphs: list[Paragraph]) -> None:
    """Ensure complex script bold/italic match the intended style across all runs."""
    for p in paragraphs:
        for run in p.runs:
            rpr = run._element.find(qn("w:rPr"))
            if rpr is not None:
                b_elem = rpr.find(qn("w:b"))
                is_bold = b_elem is not None and b_elem.get(qn("w:val"), "true").lower() not in ("0", "false", "off")
                bcs_elem = rpr.find(qn("w:bCs"))
                if not is_bold and bcs_elem is not None:
                    rpr.remove(bcs_elem)

                i_elem = rpr.find(qn("w:i"))
                is_italic = i_elem is not None and i_elem.get(qn("w:val"), "true").lower() not in ("0", "false", "off")
                ics_elem = rpr.find(qn("w:iCs"))
                if not is_italic and ics_elem is not None:
                    rpr.remove(ics_elem)


def _convert_run_text(
    run: Run,
    font_name: str | None,
    doc_default: str | None,
    replacement_font: str,
) -> bool:
    """Convert Bijoy text in a single run to Unicode.

    Modifies the run in-place: converts text and updates font.
    Safely skips text that already contains Unicode Bengali code points.

    Args:
        run: The run to process.
        font_name: The detected font name for this run.
        doc_default: The document's default font name.
        replacement_font: Unicode font name to assign.

    Returns:
        True if the run was converted, False otherwise.
    """
    effective_font = font_name or doc_default

    if not is_legacy_bijoy_run(run, effective_font):
        return False

    original_text = run.text
    converted_text = bijoy_to_unicode(original_text)

    if converted_text == original_text:
        return False

    run.text = converted_text
    _set_run_font(run, replacement_font)
    return True


def _process_paragraphs(
    paragraphs: list[Paragraph],
    doc_default: str | None,
    replacement_font: str,
) -> int:
    """Process all paragraphs, grouping adjacent Bijoy runs to preserve multi-run ligatures."""
    converted_count = 0
    for paragraph in paragraphs:
        para_font = _get_paragraph_default_font(paragraph)
        runs = paragraph.runs
        if not runs:
            continue

        i = 0
        while i < len(runs):
            run = runs[i]
            font_name = _get_run_font_name(run) or para_font
            effective_font = font_name or doc_default
            if not is_legacy_bijoy_run(run, effective_font):
                i += 1
                continue

            # Group adjacent Bijoy runs to prevent vowel-sign / pre-kar splits
            group = [run]
            j = i + 1
            while j < len(runs):
                next_run = runs[j]
                next_font = _get_run_font_name(next_run) or para_font
                next_effective = next_font or doc_default
                if is_legacy_bijoy_run(next_run, next_effective):
                    group.append(next_run)
                    j += 1
                else:
                    break

            combined_text = "".join(r.text for r in group)
            converted_text = bijoy_to_unicode(combined_text)

            group[0].text = converted_text
            _set_run_font(group[0], replacement_font)
            converted_count += 1

            for extra_run in group[1:]:
                extra_run.text = ""
                paragraph._p.remove(extra_run._r)

            i = j

    return converted_count


def convert_bijoy_in_docx(
    input_path: Path,
    output_path: Path,
    replacement_font: str | None = None,
) -> dict:
    """Convert all Bijoy text in a DOCX to Unicode and save."""
    doc = Document(str(input_path))
    detection = detect_fonts(doc)

    if not detection.has_legacy_fonts:
        shutil.copy2(input_path, output_path)
        return {
            "bijoy_detected": False,
            "legacy_fonts": [],
            "runs_converted": 0,
            "message": "No legacy Bijoy fonts detected. Document is already Unicode.",
        }

    target_font = replacement_font or get_unicode_replacement_font()
    doc_default = _get_document_default_font(doc)
    total_converted = 0

    # Body paragraphs
    total_converted += _process_paragraphs(doc.paragraphs, doc_default, target_font)

    # Tables
    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            # Handle serial number column (cell 0) with automatic w:numPr numbering
            if r_idx > 0 and len(row.cells) > 0:
                cell0 = row.cells[0]
                for p in cell0.paragraphs:
                    pPr = p._p.pPr
                    has_numPr = pPr is not None and pPr.find(qn("w:numPr")) is not None
                    clean_p_text = p.text.strip()
                    is_serial_num = clean_p_text.isdigit() or (
                        clean_p_text.endswith(".") and clean_p_text[:-1].isdigit()
                    )
                    if has_numPr or is_serial_num:
                        if pPr is not None:
                            numPr = pPr.find(qn("w:numPr"))
                            if numPr is not None:
                                pPr.remove(numPr)
                        p.text = f"{to_bengali_number(r_idx)}."
                        for r in p.runs:
                            _set_run_font(r, target_font)

            for cell in row.cells:
                total_converted += _process_paragraphs(cell.paragraphs, doc_default, target_font)

    # Headers and Footers
    for section in doc.sections:
        if section.header:
            total_converted += _process_paragraphs(section.header.paragraphs, doc_default, target_font)
        if section.footer:
            total_converted += _process_paragraphs(section.footer.paragraphs, doc_default, target_font)

    # Clean up phantom complex script bold/italic attributes across the entire document
    # so regular text never renders as bold
    _clean_phantom_complex_script_styles(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _clean_phantom_complex_script_styles(cell.paragraphs)
    for section in doc.sections:
        if section.header:
            _clean_phantom_complex_script_styles(section.header.paragraphs)
        if section.footer:
            _clean_phantom_complex_script_styles(section.footer.paragraphs)

    doc.save(str(output_path))
    return {
        "bijoy_detected": True,
        "legacy_fonts": list(detection.legacy_font_names),
        "runs_converted": total_converted,
        "replacement_font": target_font,
        "message": f"Converted {total_converted} text runs to Unicode using '{target_font}'.",
    }


def _convert_docx_to_pdf_word_sync(input_docx_path: Path, output_pdf_path: Path) -> None:
    """Synchronous worker that exports a DOCX to PDF using Microsoft Word COM.

    Ensures 100% exact rendering matching what Microsoft Word displays on screen.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        doc = word.Documents.Open(
            str(input_docx_path.resolve()),
            ReadOnly=True,
            ConfirmConversions=False,
            Visible=False,
        )
        try:
            # wdExportFormatPDF = 17
            # wdExportOptimizeForPrint = 0
            doc.ExportAsFixedFormat(
                OutputFileName=str(output_pdf_path.resolve()),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                BitmapMissingFonts=True,
                DocStructureTags=True,
            )
        finally:
            doc.Close(SaveChanges=False)
    finally:
        if word:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


async def convert_docx_to_pdf_word(input_docx_path: Path, output_pdf_path: Path) -> Path:
    """Convert a DOCX file to PDF using Microsoft Word COM automation (Windows).

    Yields 100% pixel-perfect fidelity, identical to taking a screenshot in Word.

    Args:
        input_docx_path: Path to DOCX.
        output_pdf_path: Path to destination PDF.

    Returns:
        Path to output PDF.
    """
    logger.info("Converting via Microsoft Word COM (Native High Fidelity): %s", input_docx_path.name)
    await asyncio.to_thread(_convert_docx_to_pdf_word_sync, input_docx_path, output_pdf_path)

    if not output_pdf_path.exists() or output_pdf_path.stat().st_size == 0:
        raise RuntimeError("Microsoft Word did not produce a valid PDF file.")

    logger.info("PDF generated via Word: %s (%d bytes)", output_pdf_path, output_pdf_path.stat().st_size)
    return output_pdf_path


async def convert_docx_to_pdf_libreoffice(
    input_docx_path: Path,
    output_dir: Path,
) -> Path:
    """Convert a DOCX file to PDF using LibreOffice headless.

    Used on Linux/Docker environments or as a fallback.
    """
    libreoffice_bin = _find_libreoffice()

    cmd = [
        libreoffice_bin,
        "--headless",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(input_docx_path),
    ]

    logger.info("Running LibreOffice conversion: %s", " ".join(cmd))

    try:
        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            env={
                **os.environ,
                "HOME": str(output_dir),
            },
        )

        stdout, stderr = await asyncio.wait_for(
            process.communicate(),
            timeout=CONVERSION_TIMEOUT,
        )

        if process.returncode != 0:
            error_msg = stderr.decode("utf-8", errors="replace")
            logger.error("LibreOffice conversion failed: %s", error_msg)
            raise RuntimeError(
                f"LibreOffice conversion failed (exit code {process.returncode}): {error_msg}"
            )

    except asyncio.TimeoutError:
        logger.error("LibreOffice conversion timed out after %d seconds", CONVERSION_TIMEOUT)
        try:
            process.kill()
        except Exception:
            pass
        raise RuntimeError(
            f"Conversion timed out after {CONVERSION_TIMEOUT} seconds. "
            "The document may be too large or complex."
        )

    pdf_name = input_docx_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name

    if not pdf_path.exists():
        pdf_files = list(output_dir.glob("*.pdf"))
        if pdf_files:
            pdf_path = pdf_files[0]
        else:
            raise RuntimeError("LibreOffice did not produce a PDF file.")

    logger.info("PDF generated via LibreOffice: %s (%d bytes)", pdf_path, pdf_path.stat().st_size)
    return pdf_path


async def convert_docx_to_pdf(
    input_docx_path: Path,
    output_dir: Path,
) -> Path:
    """Convert DOCX to PDF choosing the best available engine.

    Automatically uses Microsoft Word COM when available on Windows
    for 100% screenshot-like fidelity, otherwise uses LibreOffice.
    """
    use_word = False
    if CONVERSION_ENGINE == "word":
        use_word = True
    elif CONVERSION_ENGINE == "libreoffice":
        use_word = False
    else:  # "auto"
        use_word = is_word_available()

    if use_word:
        try:
            pdf_path = output_dir / f"{input_docx_path.stem}.pdf"
            return await convert_docx_to_pdf_word(input_docx_path, pdf_path)
        except Exception as e:
            logger.warning("Word conversion failed (%s). Attempting LibreOffice fallback...", e)
            try:
                return await convert_docx_to_pdf_libreoffice(input_docx_path, output_dir)
            except Exception:
                raise e

    return await convert_docx_to_pdf_libreoffice(input_docx_path, output_dir)


async def full_convert(
    input_path: Path,
    output_dir: Path,
    exact_mode: bool = True,
) -> tuple[Path, dict]:
    """Full conversion pipeline with exact screenshot fidelity support.

    Args:
        input_path: Path to the uploaded DOCX file.
        output_dir: Temporary directory for output files.
        exact_mode: If True and running on Windows with Word and SutonnyMJ,
                   renders the original document directly for 100% exact
                   screenshot-like fidelity.

    Returns:
        Tuple of (pdf_path, conversion_info).
    """
    doc = Document(str(input_path))
    detection = detect_fonts(doc)

    word_ready = is_word_available()
    check_font_availability()

    # Case 1: 100% Exact Screenshot Mode on Windows with Word + SutonnyMJ
    if word_ready and exact_mode and SUTONNYMJ_AVAILABLE and detection.has_legacy_fonts:
        pdf_path = output_dir / f"{input_path.stem}.pdf"
        await convert_docx_to_pdf_word(input_path, pdf_path)
        return pdf_path, {
            "engine": "Microsoft Word (Native - 100% Screenshot Match)",
            "exact_screenshot_mode": True,
            "bijoy_detected": True,
            "legacy_fonts": list(detection.legacy_font_names),
            "runs_converted": 0,
            "message": "Rendered with Microsoft Word native engine preserving exact SutonnyMJ typography and layout.",
        }

    # Case 2: Document has Bijoy fonts, convert to high-fidelity Unicode
    if detection.has_legacy_fonts:
        replacement_font = get_unicode_replacement_font()
        converted_docx_path = output_dir / "converted.docx"
        conversion_info = convert_bijoy_in_docx(input_path, converted_docx_path, replacement_font=replacement_font)
        pdf_path = await convert_docx_to_pdf(converted_docx_path, output_dir)
        engine_name = "Microsoft Word (Native)" if word_ready else "LibreOffice Headless"
        conversion_info["engine"] = engine_name
        conversion_info["exact_screenshot_mode"] = word_ready
        return pdf_path, conversion_info

    # Case 3: Document is already Unicode Bengali or standard text
    pdf_path = await convert_docx_to_pdf(input_path, output_dir)
    engine_name = "Microsoft Word (Native)" if word_ready else "LibreOffice Headless"
    return pdf_path, {
        "engine": engine_name,
        "exact_screenshot_mode": word_ready,
        "bijoy_detected": False,
        "legacy_fonts": [],
        "runs_converted": 0,
        "message": "Document is already Unicode. Rendered directly preserving layout.",
    }


def _find_libreoffice() -> str:
    """Find the LibreOffice binary on the system."""
    candidates = [
        "libreoffice",
        "soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for candidate in candidates:
        if shutil.which(candidate):
            return candidate

    raise RuntimeError(
        "Neither Microsoft Word nor LibreOffice is available on the system. "
        "Please install LibreOffice or run on Windows with Microsoft Word installed."
    )
