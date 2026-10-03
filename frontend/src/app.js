/**
 * Bangla Word to PDF — Frontend Application
 *
 * Handles file upload (drag-and-drop + click), conversion request,
 * progress display, and PDF download.
 */

// ── Configuration ─────────────────────────────────────────────
// CONVERTER_URL can be set as an environment variable during build
// or defaults to the current origin (same-origin API)
const CONVERTER_URL = import.meta.env.VITE_CONVERTER_URL || '';
const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20 MB
const ALLOWED_EXTENSIONS = ['.docx'];

// ── DOM Elements ──────────────────────────────────────────────
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const fileSize = document.getElementById('file-size');
const removeFileBtn = document.getElementById('remove-file');
const convertBtn = document.getElementById('convert-btn');
const convertBtnText = document.getElementById('convert-btn-text');
const progress = document.getElementById('progress');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const successResult = document.getElementById('success-result');
const successDetail = document.getElementById('success-detail');
const downloadBtn = document.getElementById('download-btn');
const convertAnother = document.getElementById('convert-another');
const errorResult = document.getElementById('error-result');
const errorDetail = document.getElementById('error-detail');
const tryAgain = document.getElementById('try-again');

// ── State ─────────────────────────────────────────────────────
let selectedFile = null;
let pdfBlob = null;
let pdfFilename = '';

// ── Utility Functions ─────────────────────────────────────────

/**
 * Format bytes into a human-readable string.
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Get file extension from filename.
 */
function getExtension(filename) {
  const idx = filename.lastIndexOf('.');
  return idx >= 0 ? filename.substring(idx).toLowerCase() : '';
}

/**
 * Show an element (remove hidden attribute).
 */
function show(el) {
  el.hidden = false;
}

/**
 * Hide an element (set hidden attribute).
 */
function hide(el) {
  el.hidden = true;
}

// ── File Selection ────────────────────────────────────────────

/**
 * Handle a selected file — validate and display info.
 */
function handleFileSelect(file) {
  // Validate extension
  const ext = getExtension(file.name);
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    showError(`Invalid file type "${ext}". Only .docx files are accepted.`);
    return;
  }

  // Validate size
  if (file.size > MAX_FILE_SIZE) {
    showError(`File too large (${formatFileSize(file.size)}). Maximum size is 20 MB.`);
    return;
  }

  selectedFile = file;
  pdfBlob = null;
  pdfFilename = '';

  // Update UI
  fileName.textContent = file.name;
  fileSize.textContent = formatFileSize(file.size);

  hide(dropzone);
  show(fileInfo);
  hide(successResult);
  hide(errorResult);
  hide(progress);

  convertBtn.disabled = false;
  convertBtnText.textContent = 'Convert to PDF';
  convertBtn.classList.remove('btn--loading');
}

/**
 * Remove the selected file and reset UI.
 */
function resetUI() {
  selectedFile = null;
  pdfBlob = null;
  pdfFilename = '';

  show(dropzone);
  hide(fileInfo);
  hide(successResult);
  hide(errorResult);
  hide(progress);

  convertBtn.disabled = true;
  convertBtnText.textContent = 'Convert to PDF';
  convertBtn.classList.remove('btn--loading');
  show(convertBtn);

  // Reset file input
  fileInput.value = '';
}

// ── Drag and Drop ─────────────────────────────────────────────

dropzone.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dropzone.classList.add('dropzone--active');
});

dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('dropzone--active');
});

dropzone.addEventListener('dragleave', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dropzone--active');
});

dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dropzone--active');

  const files = e.dataTransfer.files;
  if (files.length > 0) {
    handleFileSelect(files[0]);
  }
});

// Click to upload
dropzone.addEventListener('click', () => {
  fileInput.click();
});

dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener('change', () => {
  if (fileInput.files.length > 0) {
    handleFileSelect(fileInput.files[0]);
  }
});

// Remove file
removeFileBtn.addEventListener('click', () => {
  resetUI();
});

// ── Conversion ────────────────────────────────────────────────

convertBtn.addEventListener('click', async () => {
  if (!selectedFile) return;

  // Show loading state
  convertBtn.disabled = true;
  convertBtn.classList.add('btn--loading');
  convertBtnText.textContent = 'Converting...';
  hide(errorResult);
  hide(successResult);
  show(progress);
  progressFill.style.width = '0%';
  progressFill.classList.add('progress__fill--indeterminate');
  progressText.textContent = 'Uploading and converting your document...';

  try {
    const formData = new FormData();
    formData.append('file', selectedFile);

    const response = await fetch(`${CONVERTER_URL}/convert`, {
      method: 'POST',
      body: formData,
    });

    progressFill.classList.remove('progress__fill--indeterminate');

    if (!response.ok) {
      let errorMessage;
      try {
        const errorData = await response.json();
        errorMessage = errorData.detail || `Server error (${response.status})`;
      } catch {
        errorMessage = `Server error (${response.status})`;
      }
      throw new Error(errorMessage);
    }

    // Success — get the PDF blob
    progressFill.style.width = '90%';
    progressText.textContent = 'Preparing your PDF...';

    pdfBlob = await response.blob();

    // Extract filename from Content-Disposition header
    const disposition = response.headers.get('Content-Disposition');
    if (disposition) {
      const match = disposition.match(/filename="?([^";\n]+)"?/);
      if (match) {
        pdfFilename = match[1];
      }
    }
    if (!pdfFilename) {
      const stem = selectedFile.name.replace(/\.docx$/i, '');
      pdfFilename = `${stem}.pdf`;
    }

    // Show success
    progressFill.style.width = '100%';
    progressText.textContent = 'Done!';

    // Check conversion headers
    const bijoyDetected = response.headers.get('X-Bijoy-Detected') === 'true';
    const runsConverted = parseInt(response.headers.get('X-Runs-Converted') || '0', 10);

    let detailText = 'Your PDF is ready to download.';
    if (bijoyDetected) {
      detailText = `Bijoy text detected and converted (${runsConverted} text runs). Your PDF is ready.`;
    }
    successDetail.textContent = detailText;

    setTimeout(() => {
      hide(progress);
      hide(convertBtn);
      show(successResult);
    }, 500);

  } catch (err) {
    progressFill.classList.remove('progress__fill--indeterminate');
    hide(progress);
    showError(err.message);
    convertBtn.disabled = false;
    convertBtn.classList.remove('btn--loading');
    convertBtnText.textContent = 'Convert to PDF';
  }
});

// ── Download ──────────────────────────────────────────────────

downloadBtn.addEventListener('click', () => {
  if (!pdfBlob) return;

  const url = URL.createObjectURL(pdfBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = pdfFilename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// ── Convert Another / Try Again ───────────────────────────────

convertAnother.addEventListener('click', resetUI);
tryAgain.addEventListener('click', () => {
  hide(errorResult);
  show(convertBtn);
  convertBtn.disabled = false;
  convertBtnText.textContent = 'Convert to PDF';
  convertBtn.classList.remove('btn--loading');
});

// ── Error Display ─────────────────────────────────────────────

function showError(message) {
  errorDetail.textContent = message;
  show(errorResult);
}

// ── Health Check (optional background check) ──────────────────

async function checkHealth() {
  try {
    const res = await fetch(`${CONVERTER_URL}/health`);
    if (res.ok) {
      console.log('✓ Backend is healthy');
    }
  } catch {
    console.warn('Backend health check failed — ensure the server is running');
  }
}

// Run health check on load
checkHealth();
