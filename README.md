# Bangla Word to PDF

**বাংলা ওয়ার্ড টু পিডিএফ** — Convert Bangla Word (.docx) documents to PDF while preserving the original formatting.

Supports both **Unicode Bengali** and **legacy Bijoy/ANSI** (SutonnyMJ, SutonnyOMJ, Nikosh) encoded documents.

![Python](https://img.shields.io/badge/Python-3.12-blue)
![FastAPI](https://img.shields.io/badge/FastAPI-0.115-green)
![LibreOffice](https://img.shields.io/badge/LibreOffice-headless-orange)
![Docker](https://img.shields.io/badge/Docker-ready-blue)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/abdul-kioum-ahmed-sumon/bangla-doc-to-pdf)


---

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Features](#features)
4. [Quick Start with Docker](#quick-start-with-docker)
5. [Local Development Setup](#local-development-setup)
6. [Adding SutonnyMJ Fonts](#adding-sutonnymj-fonts)
7. [Running the Backend](#running-the-backend)
8. [Running the Frontend](#running-the-frontend)
9. [Docker Build](#docker-build)
10. [Docker Compose](#docker-compose)
11. [Testing](#testing)
12. [API Endpoints](#api-endpoints)
13. [Environment Variables](#environment-variables)
14. [Deployment to Render](#deployment-to-render)
15. [Security](#security)
16. [Privacy](#privacy)
17. [Technical Limitations](#technical-limitations)

---

## Overview

Many Bangladeshi government offices, legal firms, and businesses use **Bijoy Bayanno** with **SutonnyMJ** fonts to create Word documents. These documents store Bengali text using ANSI encoding, which displays correctly only when the SutonnyMJ font is installed. Without proper handling, converting such documents to PDF produces **corrupted, unreadable text** like:

```
µwgK  GmAvB  wewc  w`bvRcyi  e¸ov
```

This project solves this problem by:

1. **Detecting** legacy Bijoy/ANSI fonts (SutonnyMJ, SutonnyOMJ, etc.) in the DOCX
2. **Converting** Bijoy-encoded text to proper Unicode Bengali — run by run, preserving all formatting
3. **Rendering** the PDF using LibreOffice headless for high-fidelity output

Documents already using Unicode Bengali pass through without text conversion.

---

## Architecture

```
GitHub Repository
        ↓
    Docker Image
        ↓
  Python + FastAPI (Backend)
        ↓
    LibreOffice Headless
        ↓
   DOCX → PDF Conversion
        ↓
     Public API
        ↓
   Web Frontend (Vite)
        ↓
    Download PDF
```

### How It Works

1. User uploads a `.docx` file via the web UI or API
2. Backend validates the file (extension, MIME type, size)
3. `detector.py` scans all text runs for legacy Bijoy fonts
4. If Bijoy fonts are found, `bijoy.py` converts each affected text run to Unicode
5. Font names are updated to Unicode-compatible fonts (Noto Sans Bengali)
6. The modified DOCX is converted to PDF by LibreOffice `--headless`
7. PDF is returned to the user; all temporary files are deleted

---

## Features

- ✅ **Bijoy/SutonnyMJ detection and conversion** — automatic, per-run
- ✅ **Unicode Bengali pass-through** — no unnecessary conversion
- ✅ **Full formatting preservation** — bold, italic, underline, font size, alignment, spacing, tables, images, headers, footers, page breaks, margins
- ✅ **LibreOffice headless** for production-quality PDF
- ✅ **Drag & drop web UI** with progress and download
- ✅ **REST API** for programmatic access
- ✅ **Docker-ready** with all dependencies
- ✅ **Security** — file validation, size limits, rate limiting, temporary file cleanup
- ✅ **Privacy** — no files are stored; everything is cleaned up after conversion

---

## Quick Start with Docker

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/bangla-doc-to-pdf.git
cd bangla-doc-to-pdf

# (Optional) Add your SutonnyMJ fonts
cp /path/to/SutonnyMJ.ttf fonts/
cp /path/to/SutonnyOMJ.ttf fonts/

# Build and run
docker compose up --build

# Open http://localhost:8000
```

---

## Local Development Setup

### Prerequisites

- **Python 3.10+**
- **Node.js 18+** (for frontend)
- **LibreOffice** (for PDF conversion)

### Install LibreOffice

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install libreoffice libreoffice-writer fonts-noto
```

**macOS:**
```bash
brew install --cask libreoffice
```

**Windows:**
Download and install from [libreoffice.org](https://www.libreoffice.org/download/)

### Install Python Dependencies

```bash
cd backend
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### Install Frontend Dependencies

```bash
cd frontend
npm install
```

---

## Adding SutonnyMJ Fonts

> **⚠️ Important:** SutonnyMJ is a proprietary font. It is NOT included in this repository.

If your documents use SutonnyMJ, you must provide your legally obtained font files:

1. Place font files in the `fonts/` directory:
   ```
   fonts/
   ├── SutonnyMJ.ttf
   ├── SutonnyMJ Bold.ttf
   ├── SutonnyOMJ.ttf
   └── SutonnyOMJ Bold.ttf
   ```

2. **For local development**, install the fonts on your system:
   - **Linux:** Copy to `~/.fonts/` and run `fc-cache -fv`
   - **macOS:** Double-click each font to install via Font Book
   - **Windows:** Right-click → Install

3. **For Docker**, the fonts are automatically installed during the build.

### Without SutonnyMJ Fonts

The converter will still work without SutonnyMJ fonts:
- Bijoy text is converted to Unicode
- Noto Sans Bengali is used as the display font
- Text will be **readable** but may look slightly different from the original SutonnyMJ rendering

For pixel-perfect reproduction, provide the SutonnyMJ fonts.

---

## Running the Backend

```bash
cd backend

# Activate virtual environment
source venv/bin/activate  # Windows: venv\Scripts\activate

# Start the server
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

The API is now available at `http://localhost:8000`.

- API docs: `http://localhost:8000/docs`
- Health check: `http://localhost:8000/health`

---

## Running the Frontend

```bash
cd frontend
npm run dev
```

The frontend dev server starts at `http://localhost:3000` and proxies API requests to `http://localhost:8000`.

### Building for Production

```bash
cd frontend
npm run build
```

The built files are output to `frontend/dist/` and served by the backend.

---

## Docker Build

### Build the Image

```bash
docker build -t bangla-doc-to-pdf .
```

### Run the Container

```bash
docker run -p 8000:8000 bangla-doc-to-pdf
```

### With Custom Fonts

Place your fonts in the `fonts/` directory before building:

```bash
cp /path/to/SutonnyMJ.ttf fonts/
docker build -t bangla-doc-to-pdf .
docker run -p 8000:8000 bangla-doc-to-pdf
```

---

## Docker Compose

```bash
# Start
docker compose up --build

# Start in background
docker compose up --build -d

# View logs
docker compose logs -f

# Stop
docker compose down
```

---

## Testing

```bash
cd backend

# Install test dependencies
pip install pytest httpx

# Run tests
pytest tests/ -v

# Run with coverage
pip install pytest-cov
pytest tests/ -v --cov=app --cov-report=html
```

### Manual API Testing

```bash
# Health check
curl http://localhost:8000/health

# Convert a file
curl -X POST http://localhost:8000/convert \
  -F "file=@document.docx" \
  -o output.pdf

# Test with the sample document
curl -X POST http://localhost:8000/convert \
  -F "file=@অফিসার ও ফোর্সের তালিকা.docx" \
  -o output.pdf
```

---

## API Endpoints

### `GET /health`

Health check endpoint.

**Response:**
```json
{
  "status": "healthy",
  "service": "bangla-doc-to-pdf",
  "version": "1.0.0",
  "fonts": {
    "sutonnymj": false,
    "sutonnyomj": false,
    "noto sans bengali": true,
    "noto serif bengali": true
  }
}
```

### `POST /convert`

Convert a .docx file to PDF.

**Request:**
- Content-Type: `multipart/form-data`
- Field: `file` (the .docx file)
- Max size: 20 MB

**Response:**
- Content-Type: `application/pdf`
- Headers:
  - `X-Bijoy-Detected`: `true` or `false`
  - `X-Runs-Converted`: number of text runs converted

**Error Responses:**
| Status | Description |
|--------|-------------|
| 400 | Invalid file type, MIME type, or empty file |
| 413 | File exceeds 20 MB limit |
| 429 | Rate limit exceeded |
| 500 | Conversion error (LibreOffice failure, etc.) |

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `CONVERTER_URL` | (empty) | Base URL for the converter API. Used by frontend. |
| `VITE_CONVERTER_URL` | (empty) | Same as above, used during Vite build. |
| `MAX_FILE_SIZE` | `20971520` | Maximum upload size in bytes (default 20 MB) |
| `CONVERSION_TIMEOUT` | `120` | LibreOffice conversion timeout in seconds |
| `CORS_ORIGINS` | `*` | Comma-separated list of allowed CORS origins |
| `RATE_LIMIT_WINDOW` | `60` | Rate limit window in seconds |
| `RATE_LIMIT_MAX` | `10` | Maximum requests per window |
| `UNICODE_REPLACEMENT_FONT` | `Noto Sans Bengali` | Font used to replace legacy Bijoy fonts |

---

## Deployment to Render

### Option 1: Docker Deployment (Recommended)

1. Push your repository to GitHub
2. Go to [Render Dashboard](https://dashboard.render.com/)
3. Click **New** → **Web Service**
4. Connect your GitHub repository
5. Configure:
   - **Environment:** Docker
   - **Region:** Choose your preferred region
   - **Instance Type:** Standard or higher (LibreOffice needs RAM)
   - **Health Check Path:** `/health`
6. Add environment variables:
   - `CORS_ORIGINS`: Your frontend domain
7. Click **Create Web Service**

### Option 2: Native Python Deployment

1. Create a `render.yaml` in your repo root:
   ```yaml
   services:
     - type: web
       name: bangla-doc-to-pdf
       env: docker
       dockerfilePath: ./Dockerfile
       healthCheckPath: /health
       envVars:
         - key: CORS_ORIGINS
           value: "*"
   ```

2. Push to GitHub and connect on Render.

### Frontend Deployment

If deploying the frontend separately (e.g., on Vercel, Netlify):

```bash
cd frontend
VITE_CONVERTER_URL=https://your-backend.onrender.com npm run build
```

Set `VITE_CONVERTER_URL` (or `CONVERTER_URL`) to your backend's public URL.

---

## Security

- **File Extension Validation:** Only `.docx` files are accepted
- **MIME Type Validation:** Content is checked using `python-magic` (libmagic)
- **Size Limit:** 20 MB maximum (configurable)
- **Safe Filenames:** Path traversal and null bytes are sanitized
- **No Shell Injection:** LibreOffice is invoked via `subprocess` with argument lists, not shell strings
- **Conversion Timeout:** Default 120 seconds to prevent hanging processes
- **Temporary File Cleanup:** All uploaded and generated files are deleted after response
- **Rate Limiting:** In-memory per-IP rate limiting (10 requests per 60 seconds)
- **CORS Configuration:** Configurable allowed origins
- **No Secrets in Frontend:** The frontend only contains the API URL

---

## Privacy

- 🔒 **No files are permanently stored.** All uploads and generated PDFs are deleted immediately after the response is sent.
- 🔒 **No logging of file contents.** Only metadata (filename, size, conversion status) is logged.
- 🔒 **No analytics or tracking.** The application does not collect any user data.
- 🔒 **Temporary directories** are created per-request and cleaned up on both success and failure.

---

## Technical Limitations

1. **Bijoy Conversion Accuracy:** The Bijoy-to-Unicode conversion uses the `bijoy2unicode` library which handles the vast majority of SutonnyMJ encoded text. However, extremely rare or non-standard Bijoy conjuncts may not convert perfectly. If you encounter specific characters that don't convert correctly, please open an issue.

2. **Font Substitution:** When SutonnyMJ font files are not provided, the converted text uses Noto Sans Bengali. While the text will be readable and correct, the visual appearance (metrics, spacing, weight) may differ from the original SutonnyMJ rendering.

3. **Complex Documents:** Very large documents (100+ pages) or documents with embedded OLE objects may take longer to convert. The default timeout is 120 seconds.

4. **LibreOffice Dependency:** PDF generation requires LibreOffice. This is a large dependency (~500 MB in Docker). There is no way to avoid this while maintaining formatting fidelity.

5. **Concurrent Conversions:** LibreOffice does not handle concurrent instances well. The Docker image uses 2 uvicorn workers, but heavy concurrent load may require horizontal scaling.

---

## Project Structure

```
bangla-doc-to-pdf/
│
├── backend/
│   ├── app/
│   │   ├── __init__.py
│   │   ├── main.py          # FastAPI application
│   │   ├── converter.py     # DOCX → PDF conversion pipeline
│   │   ├── bijoy.py         # Bijoy/ANSI → Unicode converter
│   │   ├── detector.py      # Legacy font detector
│   │   └── cleanup.py       # Temporary file management
│   ├── requirements.txt
│   └── tests/
│       └── test_converter.py
│
├── frontend/
│   ├── src/
│   │   ├── index.html
│   │   ├── style.css
│   │   └── app.js
│   ├── package.json
│   └── vite.config.js
│
├── fonts/
│   └── README.md             # Instructions for SutonnyMJ fonts
│
├── Dockerfile
├── docker-compose.yml
├── .dockerignore
├── .gitignore
└── README.md
```

---

## License

MIT

---

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request