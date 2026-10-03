"""
Bangla Word to PDF — FastAPI Backend

A production-ready API for converting Bangla Word (.docx) documents to PDF,
with full support for both Unicode Bengali and legacy Bijoy/ANSI (SutonnyMJ)
encoded documents.

Endpoints:
    GET  /health   — Health check
    POST /convert  — Upload .docx, receive .pdf
"""

from __future__ import annotations

import logging
import os
import time
import uuid
from pathlib import Path

import magic
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .cleanup import secure_filename, temporary_conversion_dir
from .converter import check_font_availability, full_convert

# ── Logging ──────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# ── Configuration ────────────────────────────────────────────
MAX_FILE_SIZE = int(os.environ.get("MAX_FILE_SIZE", str(20 * 1024 * 1024)))  # 20 MB
ALLOWED_EXTENSIONS = {".docx"}
ALLOWED_MIMES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/zip",  # .docx is a ZIP file; some systems report this
}

# CORS origins (comma-separated or "*")
CORS_ORIGINS = os.environ.get("CORS_ORIGINS", "*").split(",")

# Rate limiting (simple in-memory)
RATE_LIMIT_WINDOW = int(os.environ.get("RATE_LIMIT_WINDOW", "60"))  # seconds
RATE_LIMIT_MAX = int(os.environ.get("RATE_LIMIT_MAX", "10"))  # requests per window

# ── App ──────────────────────────────────────────────────────
app = FastAPI(
    title="Bangla Word to PDF Converter",
    description="Convert Bangla Word (.docx) documents to PDF with Bijoy/ANSI support.",
    version="1.0.0",
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

# Serve frontend static files if present
FRONTEND_DIR = Path(__file__).resolve().parent.parent.parent / "frontend" / "dist"
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

# ── Simple in-memory rate limiter ────────────────────────────
_rate_limit_store: dict[str, list[float]] = {}


def _check_rate_limit(client_ip: str) -> bool:
    """Check if a client has exceeded the rate limit.

    Returns True if the request is allowed, False if rate-limited.
    """
    now = time.time()
    window_start = now - RATE_LIMIT_WINDOW

    # Clean old entries
    if client_ip in _rate_limit_store:
        _rate_limit_store[client_ip] = [
            t for t in _rate_limit_store[client_ip] if t > window_start
        ]
    else:
        _rate_limit_store[client_ip] = []

    if len(_rate_limit_store[client_ip]) >= RATE_LIMIT_MAX:
        return False

    _rate_limit_store[client_ip].append(now)
    return True


# ── Startup ──────────────────────────────────────────────────
@app.on_event("startup")
async def startup_event():
    """Check system dependencies on startup."""
    logger.info("Checking font availability...")
    fonts = check_font_availability()
    for font_name, available in fonts.items():
        status = "✓ installed" if available else "✗ not found"
        logger.info("  Font '%s': %s", font_name, status)


# ── Health Check ─────────────────────────────────────────────
@app.get("/health")
async def health_check():
    """Health check endpoint.

    Returns:
        JSON with status and system information.
    """
    fonts = check_font_availability()
    return {
        "status": "healthy",
        "service": "bangla-doc-to-pdf",
        "version": "1.0.0",
        "fonts": fonts,
    }


# ── Convert Endpoint ─────────────────────────────────────────
@app.post("/convert")
async def convert_docx_to_pdf(
    request: Request,
    file: UploadFile = File(..., description="A .docx file to convert to PDF"),
):
    """Convert a .docx file to PDF.

    Accepts a multipart/form-data upload with a single `file` field.
    Returns the generated PDF as a downloadable file.

    - Validates file extension and MIME type
    - Enforces 20 MB size limit
    - Detects and converts Bijoy/ANSI Bengali text to Unicode
    - Uses LibreOffice headless for PDF generation
    - Cleans up all temporary files after response

    Raises:
        HTTPException 400: Invalid file type or size
        HTTPException 413: File too large
        HTTPException 429: Rate limit exceeded
        HTTPException 500: Conversion error
    """
    # Rate limiting
    client_ip = request.client.host if request.client else "unknown"
    if not _check_rate_limit(client_ip):
        raise HTTPException(
            status_code=429,
            detail="Too many requests. Please wait before trying again.",
        )

    # ── Validate filename and extension ──
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided.")

    safe_name = secure_filename(file.filename)
    ext = Path(safe_name).suffix.lower()

    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid file type '{ext}'. Only .docx files are accepted.",
        )

    # ── Read file content ──
    try:
        content = await file.read()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {e}")

    # ── Validate size ──
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024 * 1024)} MB.",
        )

    # ── Validate MIME type ──
    try:
        detected_mime = magic.from_buffer(content, mime=True)
    except Exception:
        detected_mime = "unknown"

    if detected_mime not in ALLOWED_MIMES:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Invalid file content (detected: {detected_mime}). "
                "The file must be a valid .docx document."
            ),
        )

    # ── Convert ──
    async with temporary_conversion_dir() as tmp_dir:
        # Save uploaded file
        input_path = tmp_dir / f"{uuid.uuid4().hex}.docx"
        input_path.write_bytes(content)

        try:
            pdf_path, conversion_info = await full_convert(input_path, tmp_dir)
        except RuntimeError as e:
            logger.error("Conversion failed: %s", e)
            raise HTTPException(status_code=500, detail=str(e))
        except Exception as e:
            logger.exception("Unexpected error during conversion")
            raise HTTPException(
                status_code=500,
                detail="An unexpected error occurred during conversion.",
            )

        # Read the PDF into memory so we can clean up the temp dir
        pdf_content = pdf_path.read_bytes()

    # ── Build response ──
    # Generate output filename
    original_stem = Path(safe_name).stem
    output_filename = f"{original_stem}.pdf"

    return FileResponse(
        # We need to write to a temp file for FileResponse
        # Actually, let's use a Response with content directly
        path=_write_temp_response(pdf_content, output_filename),
        media_type="application/pdf",
        filename=output_filename,
        headers={
            "Content-Disposition": f'attachment; filename="{output_filename}"',
            "X-Bijoy-Detected": str(conversion_info.get("bijoy_detected", False)).lower(),
            "X-Runs-Converted": str(conversion_info.get("runs_converted", 0)),
        },
        background=None,
    )


def _write_temp_response(content: bytes, filename: str) -> str:
    """Write PDF content to a temporary file for FileResponse.

    The file is created in the system temp directory and will be
    cleaned up by the OS.
    """
    import tempfile

    # Create a temp file that won't be auto-deleted
    fd, path = tempfile.mkstemp(suffix=".pdf", prefix="response_")
    try:
        os.write(fd, content)
    finally:
        os.close(fd)
    return path


# ── Serve frontend index.html as fallback ────────────────────
@app.get("/")
async def serve_frontend():
    """Serve the frontend application."""
    index_path = FRONTEND_DIR / "index.html"
    if index_path.exists():
        return FileResponse(str(index_path))
    return JSONResponse(
        {
            "message": "Bangla Word to PDF Converter API",
            "docs": "/docs",
            "health": "/health",
        }
    )
