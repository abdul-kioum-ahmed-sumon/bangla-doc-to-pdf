"""
DOCX to PDF converter with Bijoy/ANSI Bengali support.

This module handles the complete conversion pipeline:
1. Open the DOCX file
2. Detect legacy Bijoy/ANSI Bengali fonts
3. Convert Bijoy text to Unicode (preserving formatting)
4. Replace legacy font names with Unicode-compatible fonts
5. Save the modified DOCX
6. Convert to PDF using LibreOffice headless

The conversion preserves all document formatting including:
- Text styling (bold, italic, underline, font size)
- Paragraph formatting (alignment, spacing, indentation)
- Tables, images, headers, footers
- Page layout (margins, page size, page breaks)
"""

from __future__ import annotations

import asyncio
import logging
import os
import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .bijoy import bijoy_to_unicode
from .detector import (
    FontDetectionResult,
    detect_fonts,
    is_legacy_bijoy_font,
    _get_run_font_name,
    _get_paragraph_default_font,
    _get_document_default_font,
)

logger = logging.getLogger(__name__)

# Conversion timeout in seconds
CONVERSION_TIMEOUT = int(os.environ.get("CONVERSION_TIMEOUT", "120"))

# Unicode font to use as replacement for legacy Bijoy fonts
UNICODE_REPLACEMENT_FONT = os.environ.get(
    "UNICODE_REPLACEMENT_FONT", "Noto Sans Bengali"
)

# Check if SutonnyMJ fonts are available on the system
SUTONNYMJ_AVAILABLE = False


def check_font_availability() -> dict[str, bool]:
    """Check which Bengali fonts are available on the system.

    Returns:
        Dictionary mapping font names to availability status.
    """
    global SUTONNYMJ_AVAILABLE

    fonts_status: dict[str, bool] = {}

    try:
        result = subprocess.run(
            ["fc-list", ":lang=bn", "family"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        available_fonts = result.stdout.lower()

        for font_name in ["sutonnymj", "sutonnyomj", "nikosh",
                          "noto sans bengali", "noto serif bengali"]:
            fonts_status[font_name] = font_name.lower() in available_fonts

        SUTONNYMJ_AVAILABLE = fonts_status.get("sutonnymj", False)
    except (subprocess.TimeoutExpired, FileNotFoundError):
        logger.warning("fc-list not available; cannot check font availability")

    return fonts_status


def _convert_run_text(run: Run, font_name: str | None, doc_default: str | None) -> bool:
    """Convert Bijoy text in a single run to Unicode.

    Modifies the run in-place: converts text and updates font.

    Args:
        run: The run to process.
        font_name: The detected font name for this run.
        doc_default: The document's default font name.

    Returns:
        True if the run was converted, False otherwise.
    """
    effective_font = font_name or doc_default

    if not effective_font or not is_legacy_bijoy_font(effective_font):
        return False

    if not run.text or not run.text.strip():
        return False

    # Convert the text from Bijoy to Unicode
    original_text = run.text
    converted_text = bijoy_to_unicode(original_text)

    if converted_text == original_text:
        return False

    # Update the run text
    run.text = converted_text

    # Update the font to a Unicode-compatible font
    # We replace the legacy font with the Unicode replacement
    run.font.name = UNICODE_REPLACEMENT_FONT

    # Also update the rFonts XML element for full compatibility
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            for attr in ["w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"]:
                if rfonts.get(qn(attr)):
                    rfonts.set(qn(attr), UNICODE_REPLACEMENT_FONT)
        else:
            # Create rFonts element
            rfonts = rpr.makeelement(qn("w:rFonts"), {
                qn("w:ascii"): UNICODE_REPLACEMENT_FONT,
                qn("w:hAnsi"): UNICODE_REPLACEMENT_FONT,
                qn("w:cs"): UNICODE_REPLACEMENT_FONT,
            })
            rpr.insert(0, rfonts)

    logger.debug("Converted run: '%s' → '%s'", original_text[:50], converted_text[:50])
    return True


def _process_paragraphs(
    paragraphs: list[Paragraph],
    doc_default: str | None,
) -> int:
    """Process all paragraphs, converting Bijoy text in each run.

    Args:
        paragraphs: List of paragraphs to process.
        doc_default: The document's default font name.

    Returns:
        Number of runs converted.
    """
    converted_count = 0

    for paragraph in paragraphs:
        para_font = _get_paragraph_default_font(paragraph)

        for run in paragraph.runs:
            font_name = _get_run_font_name(run) or para_font
            if _convert_run_text(run, font_name, doc_default):
                converted_count += 1

    return converted_count


def convert_bijoy_in_docx(input_path: Path, output_path: Path) -> dict:
    """Convert all Bijoy text in a DOCX to Unicode and save.

    Opens the DOCX, detects legacy fonts, converts text run by run,
    and saves the modified document. This preserves all formatting
    because we only modify the text content and font name of each run.

    Args:
        input_path: Path to the input DOCX file.
        output_path: Path to save the converted DOCX file.

    Returns:
        Dictionary with conversion statistics.

    Raises:
        ValueError: If legacy fonts are detected but required fonts are missing.
    """
    doc = Document(str(input_path))

    # Detect fonts
    detection = detect_fonts(doc)
    logger.info(
        "Font detection: legacy=%s (%s), unicode=%s (%s), total_runs=%d",
        detection.has_legacy_fonts,
        detection.legacy_font_names,
        detection.has_unicode_bengali,
        detection.unicode_font_names,
        detection.total_runs,
    )

    if not detection.has_legacy_fonts:
        # No legacy fonts — just copy the file as-is
        shutil.copy2(input_path, output_path)
        return {
            "bijoy_detected": False,
            "legacy_fonts": [],
            "runs_converted": 0,
            "message": "No legacy Bijoy fonts detected. Document is already Unicode.",
        }

    # Check if we can convert (need either SutonnyMJ font or we do text conversion)
    doc_default = _get_document_default_font(doc)
    total_converted = 0

    # Process body paragraphs
    total_converted += _process_paragraphs(doc.paragraphs, doc_default)

    # Process table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total_converted += _process_paragraphs(cell.paragraphs, doc_default)

    # Process headers and footers
    for section in doc.sections:
        if section.header:
            total_converted += _process_paragraphs(
                section.header.paragraphs, doc_default
            )
        if section.footer:
            total_converted += _process_paragraphs(
                section.footer.paragraphs, doc_default
            )

    # Save the modified document
    doc.save(str(output_path))

    return {
        "bijoy_detected": True,
        "legacy_fonts": list(detection.legacy_font_names),
        "runs_converted": total_converted,
        "message": f"Converted {total_converted} text runs from Bijoy to Unicode.",
    }


async def convert_docx_to_pdf(
    input_docx_path: Path,
    output_dir: Path,
) -> Path:
    """Convert a DOCX file to PDF using LibreOffice headless.

    This is the core conversion function that uses LibreOffice's
    built-in rendering engine to produce a high-fidelity PDF.

    Args:
        input_docx_path: Path to the DOCX file to convert.
        output_dir: Directory where the PDF will be saved.

    Returns:
        Path to the generated PDF file.

    Raises:
        RuntimeError: If LibreOffice conversion fails.
        asyncio.TimeoutError: If conversion exceeds the timeout.
    """
    # Find LibreOffice binary
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
            # Set environment to avoid LibreOffice profile conflicts
            env={
                **os.environ,
                "HOME": str(output_dir),  # Use temp dir as home
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

        logger.info("LibreOffice stdout: %s", stdout.decode("utf-8", errors="replace"))

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

    # Find the output PDF
    pdf_name = input_docx_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name

    if not pdf_path.exists():
        # LibreOffice might have created it with a slightly different name
        pdf_files = list(output_dir.glob("*.pdf"))
        if pdf_files:
            pdf_path = pdf_files[0]
        else:
            raise RuntimeError(
                "LibreOffice did not produce a PDF file. "
                "Check that the document is a valid DOCX file."
            )

    logger.info("PDF generated: %s (%d bytes)", pdf_path, pdf_path.stat().st_size)
    return pdf_path


async def full_convert(input_path: Path, output_dir: Path) -> tuple[Path, dict]:
    """Full conversion pipeline: Bijoy detection/conversion + PDF generation.

    Args:
        input_path: Path to the uploaded DOCX file.
        output_dir: Temporary directory for intermediate files.

    Returns:
        Tuple of (pdf_path, conversion_info).
    """
    # Step 1: Convert Bijoy text if needed
    converted_docx_path = output_dir / "converted.docx"
    conversion_info = convert_bijoy_in_docx(input_path, converted_docx_path)

    logger.info("Bijoy conversion result: %s", conversion_info)

    # Step 2: Convert to PDF using LibreOffice
    pdf_path = await convert_docx_to_pdf(converted_docx_path, output_dir)

    return pdf_path, conversion_info


def _find_libreoffice() -> str:
    """Find the LibreOffice binary on the system.

    Returns:
        Path to the LibreOffice binary.

    Raises:
        RuntimeError: If LibreOffice is not found.
    """
    # Common locations
    candidates = [
        "libreoffice",
        "soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for candidate in candidates:
        if shutil.which(candidate):
            return candidate

    raise RuntimeError(
        "LibreOffice is not installed or not found in PATH. "
        "Please install LibreOffice: apt-get install libreoffice"
    )
