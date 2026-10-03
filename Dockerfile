# ============================================================
# Bangla Word to PDF Converter — Production Dockerfile
# ============================================================
# Multi-stage build:
#   Stage 1: Build the frontend (Vite)
#   Stage 2: Production image with Python + LibreOffice
# ============================================================

# ── Stage 1: Build Frontend ─────────────────────────────────
FROM node:20-slim AS frontend-build

WORKDIR /app/frontend
COPY frontend/package.json frontend/package-lock.json* ./
RUN npm install
COPY frontend/ ./
RUN npm run build

# ── Stage 2: Production Image ───────────────────────────────
FROM python:3.12-slim

# Metadata
LABEL maintainer="bangla-doc-to-pdf"
LABEL description="Bangla Word (.docx) to PDF converter with Bijoy/ANSI support"

# Environment
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    DEBIAN_FRONTEND=noninteractive

# Install system dependencies + LibreOffice + fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-noto-cjk \
    fonts-noto \
    fonts-noto-extra \
    fonts-lohit-beng-bengali \
    fontconfig \
    libmagic1 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install Noto Sans Bengali and Noto Serif Bengali specifically
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    || true

# Create app directory
WORKDIR /app

# Copy and install Python dependencies
COPY backend/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy user-provided fonts (if any)
COPY fonts/ /app/fonts/
RUN mkdir -p /usr/local/share/fonts/bangla && \
    find /app/fonts -name "*.ttf" -o -name "*.otf" -o -name "*.TTF" -o -name "*.OTF" | \
    xargs -I {} cp {} /usr/local/share/fonts/bangla/ 2>/dev/null || true && \
    fc-cache -fv

# Copy backend code
COPY backend/ ./backend/

# Copy built frontend
COPY --from=frontend-build /app/frontend/dist ./frontend/dist/

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request, os; port = os.environ.get('PORT', '8000'); urllib.request.urlopen(f'http://localhost:{port}/health')" || exit 1

# Run with uvicorn (reads $PORT from cloud provider or defaults to 8000)
CMD ["sh", "-c", "uvicorn backend.app.main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 2"]

