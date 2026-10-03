"""
Temporary file cleanup utilities.

Provides safe cleanup of temporary directories and files created
during DOCX → PDF conversion. Ensures no uploaded documents or
generated PDFs persist on disk after the response is sent.
"""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
from contextlib import asynccontextmanager
from pathlib import Path
from typing import AsyncGenerator

logger = logging.getLogger(__name__)


@asynccontextmanager
async def temporary_conversion_dir() -> AsyncGenerator[Path, None]:
    """Create a temporary directory for conversion and clean up after use.

    Creates a secure temporary directory with a random name. The directory
    and all its contents are deleted when the context manager exits,
    regardless of whether an exception occurred.

    Yields:
        Path to the temporary directory.

    Example:
        async with temporary_conversion_dir() as tmp_dir:
            input_path = tmp_dir / "input.docx"
            output_path = tmp_dir / "output.pdf"
            # ... do conversion ...
    """
    tmp_dir = Path(tempfile.mkdtemp(prefix="bangla_convert_"))
    try:
        yield tmp_dir
    finally:
        cleanup_directory(tmp_dir)


def cleanup_directory(dir_path: Path) -> None:
    """Safely remove a directory and all its contents.

    Args:
        dir_path: Path to the directory to remove.
    """
    try:
        if dir_path.exists():
            shutil.rmtree(dir_path, ignore_errors=False)
            logger.debug("Cleaned up temporary directory: %s", dir_path)
    except OSError as e:
        logger.warning("Failed to clean up directory %s: %s", dir_path, e)
        # Try harder with ignore_errors
        try:
            shutil.rmtree(dir_path, ignore_errors=True)
        except Exception:
            pass


def cleanup_file(file_path: Path) -> None:
    """Safely remove a single file.

    Args:
        file_path: Path to the file to remove.
    """
    try:
        if file_path.exists():
            os.remove(file_path)
            logger.debug("Cleaned up file: %s", file_path)
    except OSError as e:
        logger.warning("Failed to clean up file %s: %s", file_path, e)


def secure_filename(filename: str) -> str:
    """Generate a safe filename from the original.

    Removes path separators, null bytes, and other dangerous characters
    to prevent path traversal attacks.

    Args:
        filename: The original filename.

    Returns:
        A sanitized filename safe for filesystem use.
    """
    # Remove path separators and null bytes
    filename = filename.replace("/", "_").replace("\\", "_").replace("\0", "")

    # Remove leading dots (hidden files)
    filename = filename.lstrip(".")

    # If empty after sanitization, use a default
    if not filename:
        filename = "document.docx"

    return filename
